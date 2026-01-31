#include <windows.h>
#include <comdef.h>
#include <WbemIdl.h>
#include <oleauto.h>
#include <fstream>
#include <ctime>
#include <vector>
#include <shellapi.h>
#include <string>
#include <math.h>
#include <comip.h>
#include <taskschd.h>
#include <shlobj.h>
#include <exdisp.h>
#include "resource.h"

#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Shell32.lib")

#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 1001
#define ID_TRAY_CONFIG 2001
#define ID_TRAY_LOG 3001
#define ID_TRAY_ABOUT 4001

const wchar_t* APP_VERSION = L"v2.15";
const wchar_t* LOG_FILENAME = L"debug.log";

struct Config {
	wchar_t sensorID[256];
	int tempLow, tempMed, tempHigh;
	// Nuevos campos de color
	COLORREF colorLow, colorMed, colorHigh;
};

Config g_cfg;
const wchar_t* INI_FILE = L".\\config.ini";

typedef int (*LPMLAPI_Initialize)();
typedef int (*LPMLAPI_GetDeviceInfo)(SAFEARRAY** pDevType, SAFEARRAY** pLedCount);
typedef int (*LPMLAPI_GetDeviceNameEx)(BSTR type, DWORD index, BSTR* pDevName);
typedef int (*LPMLAPI_GetLedInfo)(BSTR, DWORD, BSTR*, SAFEARRAY**);
typedef int (*LPMLAPI_SetLedColor)(BSTR type, DWORD index, DWORD R, DWORD G, DWORD B);
typedef int (*LPMLAPI_SetLedStyle)(BSTR type, DWORD index, BSTR style);
typedef int (*LPMLAPI_SetLedSpeed)(BSTR type, DWORD index, DWORD level);
typedef int (*LPMLAPI_Release)();

// SMART POINTERS DEFINITIONS
_COM_SMARTPTR_TYPEDEF(IWbemLocator, __uuidof(IWbemLocator));
_COM_SMARTPTR_TYPEDEF(IWbemServices, __uuidof(IWbemServices));
_COM_SMARTPTR_TYPEDEF(IEnumWbemClassObject, __uuidof(IEnumWbemClassObject));
_COM_SMARTPTR_TYPEDEF(IWbemClassObject, __uuidof(IWbemClassObject));
_COM_SMARTPTR_TYPEDEF(ITaskService, __uuidof(ITaskService));
_COM_SMARTPTR_TYPEDEF(ITaskFolder, __uuidof(ITaskFolder));
_COM_SMARTPTR_TYPEDEF(IRegisteredTask, __uuidof(IRegisteredTask));
_COM_SMARTPTR_TYPEDEF(IRunningTask, __uuidof(IRunningTask));
_COM_SMARTPTR_TYPEDEF(ITaskDefinition, __uuidof(ITaskDefinition));
_COM_SMARTPTR_TYPEDEF(IActionCollection, __uuidof(IActionCollection));
_COM_SMARTPTR_TYPEDEF(IAction, __uuidof(IAction));
_COM_SMARTPTR_TYPEDEF(IExecAction, __uuidof(IExecAction));
_COM_SMARTPTR_TYPEDEF(ITaskSettings, __uuidof(ITaskSettings));
_COM_SMARTPTR_TYPEDEF(IPrincipal, __uuidof(IPrincipal));
_COM_SMARTPTR_TYPEDEF(ITriggerCollection, __uuidof(ITriggerCollection));
_COM_SMARTPTR_TYPEDEF(ITrigger, __uuidof(ITrigger));
_COM_SMARTPTR_TYPEDEF(ILogonTrigger, __uuidof(ILogonTrigger));

// VARIABLES GLOBALES 
bool g_Running = true;
bool g_LedsEnabled = true;
_bstr_t g_bstrQuery=L"";

// Variables globales de tiempo y estado
ULONGLONG g_LastResetAttempt = 0;
ULONGLONG g_ResetTimer = 0;
bool g_Resetting_sdk = false;
int g_ResetStage = 0;
ULONGLONG lastRetry = 0;
ULONGLONG lastForceRefresh = 0;
const ULONGLONG REFRESH_INTERVAL = 60000;

IWbemServicesPtr g_pSvc = nullptr;
IWbemLocatorPtr g_pLoc = nullptr;

DWORD lastR = 999, lastG = 999, lastB = 999;

BSTR g_deviceName = NULL;
HMODULE g_hLibrary = NULL;
int g_totalLeds = 0;
LPMLAPI_Initialize lpMLAPI_Initialize = nullptr;
LPMLAPI_GetDeviceInfo lpMLAPI_GetDeviceInfo = nullptr;
LPMLAPI_SetLedColor lpMLAPI_SetLedColor = nullptr;
LPMLAPI_SetLedStyle lpMLAPI_SetLedStyle = nullptr;
LPMLAPI_SetLedSpeed lpMLAPI_SetLedSpeed = nullptr;
LPMLAPI_Release lpMLAPI_Release = nullptr;
LPMLAPI_GetDeviceNameEx lpMLAPI_GetDeviceNameEx = nullptr;
LPMLAPI_GetLedInfo lpMLAPI_GetLedInfo = nullptr;
HANDLE g_hMutex = NULL;


static void RunShellNonAdmin(const wchar_t* path) {
	// Usamos explorer.exe como "escudo". 
	// Al ser un proceso que ya corre como usuario estándar, 
	// cualquier cosa que abra perderá nuestros privilegios de Admin.
	ShellExecuteW(NULL, L"open", L"explorer.exe", path, NULL, SW_SHOWNORMAL);
}

static bool IsValidHex(const wchar_t* hex) {
	if (!hex || wcslen(hex) != 7 || hex[0] != L'#') return false;
	for (int i = 1; i < 7; i++) {
		if (!iswxdigit(hex[i])) return false;
	}
	return true;
}

// Convierte una cadena L"#RRGGBB" a COLORREF (0x00BBGGRR)
static COLORREF HexToColor(const wchar_t* hex) {
	if (hex[0] == L'#') hex++;
	unsigned int r = 0, g = 0, b = 0;
	if (swscanf_s(hex, L"%02x%02x%02x", &r, &g, &b) == 3) {
		return RGB(r, g, b);
	}
	return RGB(0, 0, 0); // Negro por defecto si falla
}

// Convierte un COLORREF a una cadena L"#RRGGBB"
static void ColorToHex(COLORREF color, wchar_t* out, size_t size) {
	swprintf_s(out, size, L"#%02X%02X%02X", GetRValue(color), GetGValue(color), GetBValue(color));
}

// --- GESTIÓN SEGURA DE MEMORIA ---
static wchar_t* HeapDupString(const wchar_t* src) {
	if (!src) return nullptr;
	size_t size = (wcslen(src) + 1) * sizeof(wchar_t);
	wchar_t* dest = (wchar_t*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, size);
	if (dest) wcscpy_s(dest, wcslen(src) + 1, src);
	return dest;
}

static void ClearComboHeapData(HWND hCombo) {
	int count = (int)SendMessage(hCombo, CB_GETCOUNT, 0, 0);
	for (int i = 0; i < count; i++) {
		wchar_t* ptr = (wchar_t*)SendMessage(hCombo, CB_GETITEMDATA, i, 0);
		if (ptr && ptr != (wchar_t*)CB_ERR) HeapFree(GetProcessHeap(), 0, ptr);
	}
	SendMessage(hCombo, CB_RESETCONTENT, 0, 0);
}



static bool InitWMI() {
	// Al asignar nullptr, los Smart Pointers liberan cualquier referencia previa de forma segura
	g_pSvc = nullptr;
	g_pLoc = nullptr;

	// Usamos el método CreateInstance del propio Smart Pointer
	HRESULT hr = g_pLoc.CreateInstance(CLSID_WbemLocator, nullptr, CLSCTX_INPROC_SERVER);
	if (FAILED(hr)) return false;

	// Conexión al servidor WMI
	// Nota: ConnectServer devuelve el puntero directamente al Smart Pointer g_pSvc
	hr = g_pLoc->ConnectServer(
		_bstr_t(L"ROOT\\LibreHardwareMonitor"),
		nullptr, nullptr, nullptr, 0, nullptr, nullptr,
		&g_pSvc
	);
	if (FAILED(hr)) return false;

	// Configurar la seguridad del proxy
	hr = CoSetProxyBlanket(
		g_pSvc,
		RPC_C_AUTHN_WINNT,
		RPC_C_AUTHZ_NONE,
		nullptr,
		RPC_C_AUTHN_LEVEL_CALL,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr,
		EOAC_NONE
	);

	return SUCCEEDED(hr);
}

// --- LOGGING OPTIMIZADO ---
static void Log(const char* text) {
	
	time_t now = time(0);
	char dt[26];
	ctime_s(dt, sizeof(dt), &now);
	dt[24] = '\0';

	// Abrir en modo append (ios::app) es O(1), no depende del tamaño del archivo
	std::ofstream out(LOG_FILENAME, std::ios::app);
	if (out) {
		out << "[" << dt << "] " << text << "\n";
	}
}

// Llama a esto SOLO UNA VEZ en WinMain al iniciar
static void TrimLogFile() {
	WIN32_FILE_ATTRIBUTE_DATA fad;
	if (!GetFileAttributesExW(LOG_FILENAME, GetFileExInfoStandard, &fad)) return;

	// Si el log pesa menos de 1MB, no nos molestamos en recortar
	LARGE_INTEGER size;
	size.LowPart = fad.nFileSizeLow;
	size.HighPart = fad.nFileSizeHigh;
	if (size.QuadPart < 1024 * 1024) return;

	std::vector<std::string> lines;
	std::string line;
	{
		std::ifstream in(LOG_FILENAME);
		if (!in.is_open()) return;
		while (std::getline(in, line)) {
			lines.push_back(line);
		}
	}

	if (lines.size() > 500) {
		std::ofstream out(LOG_FILENAME, std::ios::trunc);
		if (out.is_open()) {
			// Guardamos solo las últimas 500 líneas
			for (size_t i = lines.size() - 500; i < lines.size(); ++i) {
				out << lines[i] << "\n";
			}
		}
	}
}

static HRESULT ControlScheduledTask(const wchar_t* taskName, bool start) {
	try {
		ITaskServicePtr pService;
		HRESULT hr = pService.CreateInstance(__uuidof(TaskScheduler), NULL, CLSCTX_INPROC_SERVER);
		if (FAILED(hr)) return hr;

		pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());

		ITaskFolderPtr pRootFolder;
		pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);

		IRegisteredTaskPtr pTask;
		hr = pRootFolder->GetTask(_bstr_t(taskName), &pTask);
		if (FAILED(hr)) return hr;

		if (start) {
			IRunningTaskPtr pRunningTask; // Ahora el identificador sí existe
			return pTask->Run(_variant_t(), &pRunningTask);
		}
		else {
			return pTask->Stop(0);
		}
	}
	catch (const _com_error& e) {
		return e.Error();
	}
}

static void PrepareLHMSensorWMIQuery() {
	if (wcslen(g_cfg.sensorID) == 0) return;

	std::wstring q = L"SELECT Value FROM Sensor WHERE Identifier = '" + std::wstring(g_cfg.sensorID) + L"'";
	g_bstrQuery = q.c_str(); // El operador = de _bstr_t maneja la memoria por ti

	char sensorLog[512];
	snprintf(sensorLog, sizeof(sensorLog), "[MysticFight] WMI Sensor updated to: %ls", g_cfg.sensorID);
	Log(sensorLog);
}

static void SetRunAtStartup(bool run) {
	wchar_t szPath[MAX_PATH], szDir[MAX_PATH];
	if (GetModuleFileNameW(NULL, szPath, MAX_PATH) == 0) return;
	wcscpy_s(szDir, szPath);
	wchar_t* lastSlash = wcsrchr(szDir, L'\\');
	if (lastSlash) *lastSlash = L'\0';

	try {
		ITaskServicePtr pService;
		HRESULT hr = pService.CreateInstance(__uuidof(TaskScheduler), NULL, CLSCTX_INPROC_SERVER);
		if (FAILED(hr)) throw _com_error(hr);

		pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());

		ITaskFolderPtr pRootFolder;
		pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);

		// Siempre intentar borrar la tarea previa
		pRootFolder->DeleteTask(_bstr_t(L"MysticFight"), 0);
		if (!run) return;

		ITaskDefinitionPtr pTask;
		pService->NewTask(0, &pTask);

		// --- CONFIGURACIÓN DE SEGURIDAD ---
		IPrincipalPtr pPrincipal;
		pTask->get_Principal(&pPrincipal); // Usamos get_ para obtener la interfaz
		if (pPrincipal != nullptr) {
			pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);
			pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
		}

		// --- CONFIGURACIÓN DE SETTINGS ---
		ITaskSettingsPtr pSettings;
		pTask->get_Settings(&pSettings); // Usamos get_
		if (pSettings != nullptr) {
			pSettings->put_StartWhenAvailable(VARIANT_TRUE);
			pSettings->put_DisallowStartIfOnBatteries(VARIANT_FALSE);
			pSettings->put_StopIfGoingOnBatteries(VARIANT_FALSE);
			pSettings->put_ExecutionTimeLimit(_bstr_t(L"PT0S"));
			pSettings->put_AllowHardTerminate(VARIANT_TRUE);
		}

		// --- GATILLO (Trigger) ---
		ITriggerCollectionPtr pTriggers;
		pTask->get_Triggers(&pTriggers); // Usamos get_
		ITriggerPtr pTrigger;
		pTriggers->Create(TASK_TRIGGER_LOGON, &pTrigger);

		ILogonTriggerPtr pLogonTrigger = pTrigger; // Esto sí es automático (QueryInterface)
		if (pLogonTrigger != nullptr) {
			pLogonTrigger->put_Delay(_bstr_t(L"PT30S"));
		}

		// --- ACCIÓN ---
		IActionCollectionPtr pActions;
		pTask->get_Actions(&pActions); // Usamos get_
		IActionPtr pAction;
		pActions->Create(TASK_ACTION_EXEC, &pAction);

		IExecActionPtr pExecAction = pAction; // Esto sí es automático (QueryInterface)
		if (pExecAction != nullptr) {
			pExecAction->put_Path(_bstr_t(szPath));
			pExecAction->put_WorkingDirectory(_bstr_t(szDir));
		}

		// --- REGISTRO FINAL ---
		IRegisteredTaskPtr pRegisteredTask;
		pRootFolder->RegisterTaskDefinition(
			_bstr_t(L"MysticFight"),
			pTask,
			TASK_CREATE_OR_UPDATE,
			_variant_t(), _variant_t(),
			TASK_LOGON_INTERACTIVE_TOKEN,
			_variant_t(L""),
			&pRegisteredTask
		);

		Log("[MysticFight] Task created OK");
	}
	catch (const _com_error& e) {
		char buffer[256];
		snprintf(buffer, sizeof(buffer), "[MysticFight] TaskScheduler ERROR: HRESULT 0x%08X", e.Error());
		Log(buffer);
	}
}

static bool IsTaskValid() {
	wchar_t szCurrentPath[MAX_PATH];
	if (GetModuleFileNameW(NULL, szCurrentPath, MAX_PATH) == 0) return false;

	try {
		ITaskServicePtr pService;
		HRESULT hr = pService.CreateInstance(__uuidof(TaskScheduler), NULL, CLSCTX_INPROC_SERVER);
		if (FAILED(hr)) return false;

		hr = pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
		if (FAILED(hr)) return false;

		ITaskFolderPtr pRootFolder;
		if (FAILED(pService->GetFolder(_bstr_t(L"\\"), &pRootFolder))) return false;

		IRegisteredTaskPtr pRegisteredTask;
		if (FAILED(pRootFolder->GetTask(_bstr_t(L"MysticFight"), &pRegisteredTask))) return false;

		ITaskDefinitionPtr pDefinition;
		if (FAILED(pRegisteredTask->get_Definition(&pDefinition))) return false;

		// --- ACCESO CORREGIDO A ACTIONS ---
		IActionCollectionPtr pActions;
		if (FAILED(pDefinition->get_Actions(&pActions))) return false;

		IActionPtr pAction;
		// Obtenemos la primera acción (el índice en Task Scheduler suele empezar en 1)
		if (FAILED(pActions->get_Item(1, &pAction))) return false;

		// --- QUERY INTERFACE AUTOMÁTICO ---
		IExecActionPtr pExecAction = pAction;
		if (pExecAction == nullptr) return false;

		BSTR bstrPath = NULL;
		if (SUCCEEDED(pExecAction->get_Path(&bstrPath))) {
			// Comparamos la ruta actual con la guardada en la tarea (sin distinguir mayúsculas)
			bool match = (_wcsicmp(szCurrentPath, bstrPath) == 0);
			SysFreeString(bstrPath);
			return match;
		}
	}
	catch (const _com_error&) {
		return false;
	}

	return false;
}

void SaveSettings() {
	WritePrivateProfileStringW(L"Settings", L"SensorID", g_cfg.sensorID, INI_FILE);
	WritePrivateProfileStringW(L"Settings", L"TempLow", std::to_wstring(g_cfg.tempLow).c_str(), INI_FILE);
	WritePrivateProfileStringW(L"Settings", L"TempMed", std::to_wstring(g_cfg.tempMed).c_str(), INI_FILE);
	WritePrivateProfileStringW(L"Settings", L"TempHigh", std::to_wstring(g_cfg.tempHigh).c_str(), INI_FILE);

	wchar_t hL[10], hM[10], hH[10];
	ColorToHex(g_cfg.colorLow, hL, 10);
	ColorToHex(g_cfg.colorMed, hM, 10);
	ColorToHex(g_cfg.colorHigh, hH, 10);

	WritePrivateProfileStringW(L"Settings", L"ColorLow", hL, INI_FILE);
	WritePrivateProfileStringW(L"Settings", L"ColorMed", hM, INI_FILE);
	WritePrivateProfileStringW(L"Settings", L"ColorHigh", hH, INI_FILE);
}

void LoadSettings() {
	bool needsReset = false;

	int tL = GetPrivateProfileIntW(L"Settings", L"TempLow", 50, INI_FILE);
	int tM = GetPrivateProfileIntW(L"Settings", L"TempMed", 70, INI_FILE);
	int tH = GetPrivateProfileIntW(L"Settings", L"TempHigh", 90, INI_FILE);

	wchar_t hL[10], hM[10], hH[10];
	GetPrivateProfileStringW(L"Settings", L"ColorLow", L"#00FF00", hL, 10, INI_FILE);
	GetPrivateProfileStringW(L"Settings", L"ColorMed", L"#FFFF00", hM, 10, INI_FILE);
	GetPrivateProfileStringW(L"Settings", L"ColorHigh", L"#FF0000", hH, 10, INI_FILE);

	// Validación lógica: Low < Med < High
	if (tL < 0 || tH > 110 || tL >= tM || tM >= tH ||
		!IsValidHex(hL) || !IsValidHex(hM) || !IsValidHex(hH)) {
		needsReset = true;
	}

	if (needsReset) {
		g_cfg.tempLow = 50; g_cfg.tempMed = 70; g_cfg.tempHigh = 90;
		g_cfg.colorLow = RGB(0, 255, 0); g_cfg.colorMed = RGB(255, 255, 0); g_cfg.colorHigh = RGB(255, 0, 0);
		SaveSettings();
		MessageBoxW(NULL, L"Configuration was corrupted. Factory defaults restored.", L"MysticFight", MB_OK | MB_ICONINFORMATION);
	}
	else {
		g_cfg.tempLow = tL; g_cfg.tempMed = tM; g_cfg.tempHigh = tH;
		g_cfg.colorLow = HexToColor(hL); g_cfg.colorMed = HexToColor(hM); g_cfg.colorHigh = HexToColor(hH);
	}
	GetPrivateProfileStringW(L"Settings", L"SensorID", L"", g_cfg.sensorID, 256, INI_FILE);
}

void PopulateSensorList(HWND hDlg) {
	HWND hCombo = GetDlgItem(hDlg, IDC_SENSOR_ID);
	ClearComboHeapData(hCombo);

	if (!g_pSvc) InitWMI();
	if (g_pSvc) {
		IEnumWbemClassObjectPtr pEnumerator = NULL;
		bstr_t query("SELECT Name, Identifier FROM Sensor WHERE SensorType = 'Temperature'");
		if (SUCCEEDED(g_pSvc->ExecQuery(bstr_t("WQL"), query, WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator))) {
			IWbemClassObjectPtr pclsObj = NULL;
			ULONG uReturn = 0;
			bool found = false;
			while (pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn) == S_OK) {
				_variant_t vtName, vtID;
				if (SUCCEEDED(pclsObj->Get(L"Name", 0, &vtName, 0, 0)) && SUCCEEDED(pclsObj->Get(L"Identifier", 0, &vtID, 0, 0))) {
					if (vtName.vt == VT_BSTR && vtID.vt == VT_BSTR) {
						int idx = (int)SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)vtName.bstrVal);
						if (idx != CB_ERR) {
							SendMessage(hCombo, CB_SETITEMDATA, idx, (LPARAM)HeapDupString(vtID.bstrVal));
							if (wcscmp(vtID.bstrVal, g_cfg.sensorID) == 0) {
								SendMessage(hCombo, CB_SETCURSEL, idx, 0);
								found = true;
							}
						}
					}
				}
			}
			if (!found && SendMessage(hCombo, CB_GETCOUNT, 0, 0) > 0) SendMessage(hCombo, CB_SETCURSEL, 0, 0);
		}
	}
}


// Función para decidir si el texto debe ser blanco o negro (Contraste)
static bool IsColorDark(COLORREF col) {
	double brightness = (GetRValue(col) * 299 + GetGValue(col) * 587 + GetBValue(col) * 114) / 1000.0;
	return brightness < 128.0;
}
// --- PROCEDIMIENTO DE DIÁLOGO DE CONFIGURACIÓN ---

INT_PTR CALLBACK SettingsDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
	static HBRUSH hBrushLow = NULL, hBrushMed = NULL, hBrushHigh = NULL;

	switch (message) {
	case WM_CTLCOLOREDIT: {
		HDC hdc = (HDC)wParam;
		HWND hCtrl = (HWND)lParam;
		int id = GetDlgCtrlID(hCtrl);

		HBRUSH hSelectedBrush = NULL;

		if (id == IDC_HEX_LOW)  hSelectedBrush = hBrushLow;
		if (id == IDC_HEX_MED)  hSelectedBrush = hBrushMed;
		if (id == IDC_HEX_HIGH) hSelectedBrush = hBrushHigh;

		if (hSelectedBrush) {
			wchar_t buf[10];
			GetWindowTextW(hCtrl, buf, 10);
			if (IsValidHex(buf)) {
				COLORREF c = HexToColor(buf);
				SetBkColor(hdc, c);
				// Si el fondo es oscuro, ponemos el texto blanco
				if (IsColorDark(c)) SetTextColor(hdc, RGB(255, 255, 255));
				else SetTextColor(hdc, RGB(0, 0, 0));
				return (INT_PTR)hSelectedBrush;
			}
		}
		return (INT_PTR)GetStockObject(WHITE_BRUSH);
	}
	case WM_INITDIALOG: {
		// 1. Estética: Iconos y centrar ventana
		HICON hIcon = (HICON)LoadImage(GetModuleHandle(NULL), MAKEINTRESOURCE(101), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
		if (hIcon) {
			SendMessage(hDlg, WM_SETICON, ICON_SMALL, (LPARAM)hIcon);
			SendMessage(hDlg, WM_SETICON, ICON_BIG, (LPARAM)hIcon);
		}

		if (IsTaskValid()) {
			CheckDlgButton(hDlg, IDC_CHK_STARTUP, BST_CHECKED);
		}
		else {
			CheckDlgButton(hDlg, IDC_CHK_STARTUP, BST_UNCHECKED);
		}

		hBrushLow = CreateSolidBrush(g_cfg.colorLow);
		hBrushMed = CreateSolidBrush(g_cfg.colorMed);
		hBrushHigh = CreateSolidBrush(g_cfg.colorHigh);

		RECT rcOwner, rcDlg;
		GetWindowRect(GetDesktopWindow(), &rcOwner);
		GetWindowRect(hDlg, &rcDlg);
		int x = ((rcOwner.right - rcOwner.left) - (rcDlg.right - rcDlg.left)) / 2;
		int y = ((rcOwner.bottom - rcOwner.top) - (rcDlg.bottom - rcDlg.top)) / 2;
		SetWindowPos(hDlg, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE);

		// 2. Rellenar lista y cargar valores actuales de temperatura
		PopulateSensorList(hDlg);
		SetDlgItemInt(hDlg, IDC_TEMP_LOW, g_cfg.tempLow, TRUE);
		SetDlgItemInt(hDlg, IDC_TEMP_MED, g_cfg.tempMed, TRUE);
		SetDlgItemInt(hDlg, IDC_TEMP_HIGH, g_cfg.tempHigh, TRUE);

		// 3. Cargar los colores actuales en los cajetines HEX
		SendMessage(GetDlgItem(hDlg, IDC_HEX_LOW), EM_SETLIMITTEXT, 7, 0);
		SendMessage(GetDlgItem(hDlg, IDC_HEX_MED), EM_SETLIMITTEXT, 7, 0);
		SendMessage(GetDlgItem(hDlg, IDC_HEX_HIGH), EM_SETLIMITTEXT, 7, 0);

		wchar_t hexBuf[10];
		ColorToHex(g_cfg.colorLow, hexBuf, 10);
		SetDlgItemTextW(hDlg, IDC_HEX_LOW, hexBuf);

		ColorToHex(g_cfg.colorMed, hexBuf, 10);
		SetDlgItemTextW(hDlg, IDC_HEX_MED, hexBuf);

		ColorToHex(g_cfg.colorHigh, hexBuf, 10);
		SetDlgItemTextW(hDlg, IDC_HEX_HIGH, hexBuf);

		return (INT_PTR)TRUE;
	}

	case WM_COMMAND: {
		int id = LOWORD(wParam);
		int code = HIWORD(wParam);

		// --- 1. LÓGICA DE CAMBIO EN TIEMPO REAL (COLOR PREVIEW) ---
		if (code == EN_CHANGE) {
			wchar_t buf[10];
			GetDlgItemTextW(hDlg, id, buf, 10);

			// Si el usuario termina de escribir un Hex válido (#RRGGBB)
			if (IsValidHex(buf)) {
				COLORREF newCol = HexToColor(buf);
				HBRUSH newBrush = CreateSolidBrush(newCol);

				if (id == IDC_HEX_LOW) {
					if (hBrushLow) DeleteObject(hBrushLow);
					hBrushLow = newBrush;
				}
				else if (id == IDC_HEX_MED) {
					if (hBrushMed) DeleteObject(hBrushMed);
					hBrushMed = newBrush;
				}
				else if (id == IDC_HEX_HIGH) {
					if (hBrushHigh) DeleteObject(hBrushHigh);
					hBrushHigh = newBrush;
				}

				// Forzamos a la caja de texto a repintarse con el nuevo color
				InvalidateRect(GetDlgItem(hDlg, id), NULL, TRUE);
			}
			return TRUE;
		}

		switch (id) {
			// --- 2. BOTÓN RESET DEFAULTS ---
		case IDC_BTN_RESET: {
			// Temperaturas por defecto
			SetDlgItemInt(hDlg, IDC_TEMP_LOW, 50, TRUE);
			SetDlgItemInt(hDlg, IDC_TEMP_MED, 70, TRUE);
			SetDlgItemInt(hDlg, IDC_TEMP_HIGH, 90, TRUE);

			// Colores por defecto (esto disparará EN_CHANGE automáticamente)
			SetDlgItemTextW(hDlg, IDC_HEX_LOW, L"#00FF00"); // Green
			SetDlgItemTextW(hDlg, IDC_HEX_MED, L"#FFFF00"); // Yellow
			SetDlgItemTextW(hDlg, IDC_HEX_HIGH, L"#FF0000"); // Red

			Log("[MysticFight] UI restored to factory defaults.");
			return TRUE;
		}

		// --- 3. BOTÓN SAVE (IDOK) ---
		case IDOK: {
			// A. Validar Sensor
			HWND hCombo = GetDlgItem(hDlg, IDC_SENSOR_ID);
			int idx = (int)SendMessage(hCombo, CB_GETCURSEL, 0, 0);
			if (idx == CB_ERR) {
				MessageBoxW(hDlg, L"Please select a sensor to continue.", L"Configuration Error", MB_ICONWARNING);
				return TRUE;
			}

			// B. Recoger y Validar Colores Hex
			wchar_t hL[10], hM[10], hH[10];
			GetDlgItemTextW(hDlg, IDC_HEX_LOW, hL, 10);
			GetDlgItemTextW(hDlg, IDC_HEX_MED, hM, 10);
			GetDlgItemTextW(hDlg, IDC_HEX_HIGH, hH, 10);

			if (!IsValidHex(hL) || !IsValidHex(hM) || !IsValidHex(hH)) {
				MessageBoxW(hDlg, L"Invalid color codes.\nFormat must be #RRGGBB (e.g., #FF0000).", L"Validation Error", MB_ICONERROR);
				return TRUE;
			}

			// C. Recoger y Validar Temperaturas
			int tL = GetDlgItemInt(hDlg, IDC_TEMP_LOW, NULL, TRUE);
			int tM = GetDlgItemInt(hDlg, IDC_TEMP_MED, NULL, TRUE);
			int tH = GetDlgItemInt(hDlg, IDC_TEMP_HIGH, NULL, TRUE);

			if (tL < 0 || tH > 110 || tL >= tM || tM >= tH) {
				MessageBoxW(hDlg, L"Invalid temperature range.\nOrder must be: Low < Med < High.", L"Validation Error", MB_ICONWARNING);
				return TRUE;
			}

			// D. Guardado Final en la estructura Global
			wchar_t* selectedID = (wchar_t*)SendMessage(hCombo, CB_GETITEMDATA, idx, 0);
			if (selectedID && selectedID != (wchar_t*)CB_ERR) {
				wcscpy_s(g_cfg.sensorID, selectedID);
			}

			g_cfg.tempLow = tL;  g_cfg.tempMed = tM;  g_cfg.tempHigh = tH;
			g_cfg.colorLow = HexToColor(hL);
			g_cfg.colorMed = HexToColor(hM);
			g_cfg.colorHigh = HexToColor(hH);

			char config_new[512];
			snprintf(config_new, sizeof(config_new),
				"[MysticFight] Config Updated - Sensor: %ls | Low: %dºC (%ls) | Med: %dºC (%ls) | High: %dºC (%ls)",
				g_cfg.sensorID,
				g_cfg.tempLow, hL,
				g_cfg.tempMed, hM,
				g_cfg.tempHigh, hH
			);

			Log(config_new);

			PrepareLHMSensorWMIQuery();

			lastR = 999;

			bool wantStartup = (IsDlgButtonChecked(hDlg, IDC_CHK_STARTUP) == BST_CHECKED);

			SetRunAtStartup(wantStartup);

			WritePrivateProfileStringW(L"Settings", L"BootTaskAsked", L"1", INI_FILE);

			SaveSettings();

			EndDialog(hDlg, IDOK);

			return TRUE;
		}

				 // --- 4. BOTÓN CANCEL ---
		case IDCANCEL: {
			EndDialog(hDlg, IDCANCEL);
			return TRUE;
		}
		}
		break;
	}

	case WM_DESTROY: {
		// BORRAR PINCELES PARA EVITAR FUGAS DE MEMORIA (GDI Leaks)
		if (hBrushLow)  DeleteObject(hBrushLow);
		if (hBrushMed)  DeleteObject(hBrushMed);
		if (hBrushHigh) DeleteObject(hBrushHigh);

		hBrushLow = hBrushMed = hBrushHigh = NULL;

		// Liberamos los strings del Heap del ComboBox (lo que ya tenías)
		ClearComboHeapData(GetDlgItem(hDlg, IDC_SENSOR_ID));
		break;
	}
	}
	return (INT_PTR)FALSE;
}
INT_PTR CALLBACK AboutDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
	static HFONT hFontLink = NULL;

	switch (message) {
	case WM_INITDIALOG: {
		// --- LÓGICA DE CENTRADO ---
		RECT rcDlg;
		GetWindowRect(hDlg, &rcDlg); // Obtenemos el tamaño del diálogo

		int dwWidth = rcDlg.right - rcDlg.left;
		int dwHeight = rcDlg.bottom - rcDlg.top;

		// Centrar respecto a la pantalla completa
		int x = (GetSystemMetrics(SM_CXSCREEN) - dwWidth) / 2;
		int y = (GetSystemMetrics(SM_CYSCREEN) - dwHeight) / 2;

		SetWindowPos(hDlg, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE);
		// Cargar icono estándar de 48x48
		HICON hIcon = (HICON)LoadImage(GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_ICON1), IMAGE_ICON, 48, 48, LR_SHARED);
		SendMessage(GetDlgItem(hDlg, IDC_ABOUT_ICON), STM_SETIMAGE, IMAGE_ICON, (LPARAM)hIcon);

		std::wstring versionStr = L"MysticFight " + std::wstring(APP_VERSION);
		SetDlgItemTextW(hDlg, IDC_ABOUT_VERSION, versionStr.c_str());

		// Crear fuente subrayada para el link
		HFONT hFont = (HFONT)SendMessage(GetDlgItem(hDlg, IDC_GITHUB_LINK), WM_GETFONT, 0, 0);
		LOGFONT lf;
		GetObject(hFont, sizeof(LOGFONT), &lf);
		lf.lfUnderline = TRUE;
		hFontLink = CreateFontIndirect(&lf);
		SendMessage(GetDlgItem(hDlg, IDC_GITHUB_LINK), WM_SETFONT, (WPARAM)hFontLink, TRUE);
		return (INT_PTR)TRUE;
	}

	case WM_CTLCOLORSTATIC: {
		// Pintar el link de azul
		if ((HWND)lParam == GetDlgItem(hDlg, IDC_GITHUB_LINK)) {
			SetTextColor((HDC)wParam, RGB(0, 102, 204));
			SetBkMode((HDC)wParam, TRANSPARENT);
			return (INT_PTR)GetStockObject(NULL_BRUSH);
		}
		break;
	}

	case WM_SETCURSOR: {
		// Cambiar cursor a mano en el link
		if ((HWND)wParam == GetDlgItem(hDlg, IDC_GITHUB_LINK)) {
			SetCursor(LoadCursor(NULL, IDC_HAND));
			return TRUE;
		}
		break;
	}

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) {
			if (hFontLink) {
				DeleteObject(hFontLink);
				hFontLink = NULL;
			}
			
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		if (LOWORD(wParam) == IDC_GITHUB_LINK) {
			RunShellNonAdmin(L"https://github.com/tonikelope/MysticFight");
		}
		break;
	case WM_DESTROY: // CORRECCIÓN: Borrado garantizado al destruir la ventana
		if (hFontLink) {
			DeleteObject(hFontLink);
			hFontLink = NULL;
		}
		break;
	}
	return (INT_PTR)FALSE;
}

static void FinalCleanup(HWND hWnd) {
	static bool cleaned = false;
	if (cleaned) return;
	cleaned = true;

	Log("[MysticFight] Starting cleaning...");

	// 1. APAGADO DE HARDWARE
	if (g_hLibrary) {
		// Apagar LEDs si el dispositivo es válido
		if (g_deviceName && lpMLAPI_SetLedStyle) {
			_bstr_t bstrOff(L"Off");
			lpMLAPI_SetLedStyle(g_deviceName, 0, bstrOff);

			Log("[MysticFight] LEDs power off");
		}

		
		if (lpMLAPI_Release) {
			lpMLAPI_Release();
			Log("[MysticFight] MSI SDK Released");
		}

		FreeLibrary(g_hLibrary);
		g_hLibrary = NULL;
	}

	// 2. LIMPIEZA DE BSTRs GLOBALES
	if (g_deviceName) {
		SysFreeString(g_deviceName);
		g_deviceName = NULL;
	}
	g_bstrQuery = L""; // Limpiar el wrapper de BSTR antes de CoUninitialize

	// 3. LIMPIEZA DE INTERFAZ
	if (hWnd) {
		NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA), hWnd, 1 };
		Shell_NotifyIcon(NIM_DELETE, &nid);
		UnregisterHotKey(hWnd, 1);
	}

	// 4. DESTRUCCIÓN DE SMART POINTERS (WMI)
	g_pSvc = nullptr;
	g_pLoc = nullptr;

	// 5. SINCRONIZACIÓN
	if (g_hMutex) {
		ReleaseMutex(g_hMutex);
		CloseHandle(g_hMutex);
		g_hMutex = NULL;
	}

	// 6. CERRAR COM
	CoUninitialize();

	Log("[MysticFight] Cleaning finished.");
}

static LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
	switch (message) {
	case WM_HOTKEY :
	if (wParam == 1) {
		g_LedsEnabled = !g_LedsEnabled;
		MessageBeep(g_LedsEnabled ? MB_OK : MB_ICONHAND);
		lastR = 999;
	}
	break;
	case WM_TRAYICON:
		if (lParam == WM_RBUTTONUP) {
			POINT curPoint; GetCursorPos(&curPoint);
			HMENU hMenu = CreatePopupMenu();
			if (hMenu) {
				InsertMenu(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_CONFIG, L"Settings");
				InsertMenu(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_ABOUT, L"About MysticFight");
				InsertMenu(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_LOG, L"View Debug Log");
				InsertMenu(hMenu, -1, MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
				InsertMenu(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_EXIT, L"Exit MysticFight");
				SetForegroundWindow(hWnd);
				TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, curPoint.x, curPoint.y, 0, hWnd, NULL);
				DestroyMenu(hMenu);
			}
		}
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == ID_TRAY_EXIT) g_Running = false;
		
		if (LOWORD(wParam) == ID_TRAY_LOG) {
			wchar_t fullLogPath[MAX_PATH];
			GetFullPathNameW(LOG_FILENAME, MAX_PATH, fullLogPath, NULL);
			RunShellNonAdmin(fullLogPath);
		}
		
		
		if (LOWORD(wParam) == ID_TRAY_CONFIG) {
			DialogBoxW(GetModuleHandle(NULL), MAKEINTRESOURCE(IDD_SETTINGS), hWnd, SettingsDlgProc);
		}

		if (LOWORD(wParam) == ID_TRAY_ABOUT) { 
			DialogBoxW(GetModuleHandle(NULL), MAKEINTRESOURCE(IDD_ABOUT), hWnd, AboutDlgProc);
		}
		break;
	case WM_QUERYENDSESSION:
		Log("[MysticFight] Windows Shutdown detected...");
		ShutdownBlockReasonCreate(hWnd, L"Releasing MSI Hardware...");
		g_Running = false;
		FinalCleanup(hWnd);
		ShutdownBlockReasonDestroy(hWnd);
		return TRUE;
	case WM_CLOSE:
	case WM_DESTROY:
		g_Running = false;
		FinalCleanup(hWnd);
		PostQuitMessage(0);
		return 0;
	}
	return DefWindowProc(hWnd, message, wParam, lParam);
}

static float GetCPUTempFast() {
	if (!g_pSvc || g_bstrQuery.length() == 0) return -1.0f;

	IEnumWbemClassObjectPtr pEnum = nullptr;
	HRESULT hr = g_pSvc->ExecQuery(_bstr_t(L"WQL"), g_bstrQuery,
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnum);

	if (SUCCEEDED(hr) && pEnum) {
		IWbemClassObjectPtr pObj = nullptr;
		ULONG uRet = 0;

		if (pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet) == S_OK && uRet > 0) {
			_variant_t vtVal;
			if (SUCCEEDED(pObj->Get(L"Value", 0, &vtVal, 0, 0))) {
				if (vtVal.vt == VT_R4) return vtVal.fltVal;
				if (vtVal.vt == VT_R8) return (float)vtVal.dblVal;
			}
		}
	}
	return -1.0f;
}

static void DebugListAllSensors() {
	if (!g_pSvc) return;
	IEnumWbemClassObjectPtr pEnum = nullptr;

	// Quitamos el WHERE para ver TODO lo que LHM está publicando
	HRESULT hr = g_pSvc->ExecQuery(_bstr_t(L"WQL"), _bstr_t(L"SELECT * FROM Sensor"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnum);

	if (SUCCEEDED(hr)) {
		IWbemClassObjectPtr pObj = nullptr;
		ULONG uRet = 0;
		Log("--- DUMPING ALL LHM SENSORS ---");
		while (pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet) == S_OK) {
			_variant_t vtID, vtName, vtType;
			pObj->Get(L"Identifier", 0, &vtID, 0, 0);
			pObj->Get(L"Name", 0, &vtName, 0, 0);
			pObj->Get(L"SensorType", 0, &vtType, 0, 0); // Queremos ver qué tipo dice ser

			char buf[512];
			snprintf(buf, sizeof(buf), "ID: %ls | Name: %ls | Type: %ls",
				vtID.bstrVal, vtName.bstrVal, vtType.bstrVal);
			Log(buf);
		}
		Log("--- END OF DUMP ---");
	}
}

static void ShowNotification(HWND hWnd, HINSTANCE hInstance, const wchar_t* title, const wchar_t* info) {
	NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA), hWnd, 1, NIF_INFO };
	wcscpy_s(nid.szInfoTitle, title);
	wcscpy_s(nid.szInfo, info);
	nid.dwInfoFlags = NIIF_USER | NIIF_LARGE_ICON;
	Shell_NotifyIcon(NIM_MODIFY, &nid);
}
// --- EXTRACCIÓN SEGURA DE DATOS DEL SDK ---
static int GetIntFromSafeArray(void* pRawData, VARTYPE vt, int index) {
	if (!pRawData) return 0;

	switch (vt) {
	case VT_BSTR: {
		BSTR* pBstrArray = (BSTR*)pRawData;
		return _wtoi(pBstrArray[index]);
	}
	case VT_I4:
	case VT_INT: {
		long* pLongArray = (long*)pRawData;
		return (int)pLongArray[index];
	}
	case VT_UI4:
	case VT_UINT: {
		unsigned int* pULongArray = (unsigned int*)pRawData;
		return (int)pULongArray[index];
	}
	case VT_I2: {
		short* pShortArray = (short*)pRawData;
		return (int)pShortArray[index];
	}
	default:
		return 0; // Tipo no soportado
	}
}

static bool ExtractMSIDLL() {
	wchar_t szPath[MAX_PATH];
	GetModuleFileNameW(NULL, szPath, MAX_PATH);
	wchar_t* lastSlash = wcsrchr(szPath, L'\\');
	if (lastSlash) *(lastSlash + 1) = L'\0';

	std::wstring dllPath = std::wstring(szPath) + L"MysticLight_SDK.dll";

	// Si ya existe, no hacemos nada para ahorrar tiempo y escrituras
	if (GetFileAttributesW(dllPath.c_str()) != INVALID_FILE_ATTRIBUTES) return true;

	HMODULE hModule = GetModuleHandle(NULL);
	HRSRC hRes = FindResourceW(hModule, MAKEINTRESOURCE(IDR_MSI_DLL), L"BINARY");
	if (!hRes) return false;

	HGLOBAL hData = LoadResource(hModule, hRes);
	DWORD size = SizeofResource(hModule, hRes);
	void* pData = LockResource(hData);

	HANDLE hFile = CreateFileW(dllPath.c_str(), GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) return false;

	DWORD written;
	bool success = WriteFile(hFile, pData, size, &written, NULL);
	CloseHandle(hFile);

	return success;
}



int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {

	// --- A. PRIORIDADES Y COM ---
	SetPriorityClass(GetCurrentProcess(), HIGH_PRIORITY_CLASS);
	SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);

	HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
	if (FAILED(hr)) {
		Log("[MysticFight] Critical: COM Initialization failed.");
		return 1;
	}

	// --- B. CONTROL DE INSTANCIA ÚNICA ---
	g_hMutex = CreateMutex(NULL, TRUE, L"Global\\MysticFight_Unique_Mutex");
	if (g_hMutex == NULL || GetLastError() == ERROR_ALREADY_EXISTS) {
		if (g_hMutex) {
			MessageBoxW(NULL, L"MysticFight is already running.", L"Information", MB_OK | MB_ICONINFORMATION);
			CloseHandle(g_hMutex);
		}
		CoUninitialize();
		return 0;
	}

	// --- C. PREPARAR COMPONENTES (DLL) ---
	if (!ExtractMSIDLL() || !(g_hLibrary = LoadLibrary(L"MysticLight_SDK.dll"))) {
		Log("[MysticFight] FATAL: Component preparation failed.");
		MessageBoxW(NULL, L"Critical Error: Required components (DLL) could not be prepared.", L"MysticFight - Error", MB_OK | MB_ICONERROR);

		if (g_hMutex) {
			ReleaseMutex(g_hMutex);
			CloseHandle(g_hMutex);
		}
		CoUninitialize();
		return 1;
	}

	// --- D. INICIALIZACIÓN DE LOGS Y CONFIGURACIÓN ---
	TrimLogFile();

	char versionMsg[128];
	snprintf(versionMsg, sizeof(versionMsg), "[MysticFight] MysticFight %ls started", APP_VERSION);
	Log(versionMsg);

	LoadSettings();

	wchar_t hL[10], hH[10], hA[10];
	ColorToHex(g_cfg.colorLow, hL, 10);
	ColorToHex(g_cfg.colorMed, hH, 10);
	ColorToHex(g_cfg.colorHigh, hA, 10);

	char startupCfg[512];
	snprintf(startupCfg, sizeof(startupCfg),
		"[MysticFight] Config Loaded - Sensor: %ls | Low: %dºC (%ls) | Med: %dºC (%ls) | High: %dºC (%ls)",
		g_cfg.sensorID,
		g_cfg.tempLow, hL,
		g_cfg.tempMed, hH,
		g_cfg.tempHigh, hA
	);
	Log(startupCfg);

	// --- E. REGISTRO DE VENTANA E ICONO (¡ESTO DEBE IR ANTES DEL SDK!) ---
	wchar_t windowTitle[100];
	swprintf_s(windowTitle, L"MysticFight %s (by tonikelope)", APP_VERSION);

	WNDCLASS wc = { 0 };
	wc.lpfnWndProc = WndProc;
	wc.hInstance = hInstance;
	wc.lpszClassName = L"MysticFight_Class";
	wc.hIcon = (HICON)LoadImage(hInstance, MAKEINTRESOURCE(101), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);

	RegisterClass(&wc);

	HWND hWnd = CreateWindowEx(0, wc.lpszClassName, windowTitle, 0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);

	// Registramos el icono del tray inmediatamente para que sea visible si el SDK tarda
	NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA), hWnd, 1, NIF_ICON | NIF_MESSAGE | NIF_TIP, WM_TRAYICON };
	nid.hIcon = (HICON)LoadImage(hInstance, MAKEINTRESOURCE(101), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
	swprintf_s(nid.szTip, L"MysticFight %s (by tonikelope)", APP_VERSION);
	Shell_NotifyIcon(NIM_ADD, &nid);

	RegisterHotKey(hWnd, 1, MOD_CONTROL | MOD_SHIFT | MOD_ALT, 0x4C);
	ShowNotification(hWnd, hInstance, windowTitle, L"Let's dance baby");

	// --- F. VINCULAR FUNCIONES DE LA DLL ---
	lpMLAPI_Initialize = (LPMLAPI_Initialize)GetProcAddress(g_hLibrary, "MLAPI_Initialize");
	lpMLAPI_GetDeviceInfo = (LPMLAPI_GetDeviceInfo)GetProcAddress(g_hLibrary, "MLAPI_GetDeviceInfo");
	lpMLAPI_SetLedColor = (LPMLAPI_SetLedColor)GetProcAddress(g_hLibrary, "MLAPI_SetLedColor");
	lpMLAPI_SetLedStyle = (LPMLAPI_SetLedStyle)GetProcAddress(g_hLibrary, "MLAPI_SetLedStyle");
	lpMLAPI_SetLedSpeed = (LPMLAPI_SetLedSpeed)GetProcAddress(g_hLibrary, "MLAPI_SetLedSpeed");
	lpMLAPI_Release = (LPMLAPI_Release)GetProcAddress(g_hLibrary, "MLAPI_Release");
	lpMLAPI_GetDeviceNameEx = (LPMLAPI_GetDeviceNameEx)GetProcAddress(g_hLibrary, "MLAPI_GetDeviceNameEx");
	lpMLAPI_GetLedInfo = (LPMLAPI_GetLedInfo)GetProcAddress(g_hLibrary, "MLAPI_GetLedInfo");

	// --- G. INICIALIZACIÓN DEL SDK (Con tus reintentos originales) ---
	Log("[MysticFight] Attempting to initialize SDK...");
	if (!lpMLAPI_Initialize || lpMLAPI_Initialize() != 0) {
		bool initialized = false;
		for (int i = 1; i <= 10; i++) {
			char retryMsg[100];
			snprintf(retryMsg, sizeof(retryMsg), "[MysticFight] Attempt %d/10 failed. Retrying in 5s...", i);
			Log(retryMsg);

			if (lpMLAPI_Initialize && lpMLAPI_Initialize() == 0) {
				Log("[MysticFight] SDK Initialized successfully on retry.");
				initialized = true;
				break;
			}

			if (i == 10) {
				Log("[MysticFight] FATAL: SDK could not be initialized after 10 attempts.");
				MessageBox(NULL, L"MSI SDK Initialization failed. Ensure MSI Center/Dragon Center is installed.", L"Critical Error", MB_OK | MB_ICONERROR);
				FinalCleanup(hWnd);
				return 1;
			}
			Sleep(5000);
		}
	}
	else {
		Log("[MysticFight] SDK Initialized successfully at first attempt.");
	}

	// --- H. BLOQUE DE DETECCIÓN DETALLADO DE HARDWARE ---
	SAFEARRAY* pDevType = nullptr;
	SAFEARRAY* pLedCount = nullptr;

	if (lpMLAPI_GetDeviceInfo && lpMLAPI_GetDeviceInfo(&pDevType, &pLedCount) == 0 && pDevType && pLedCount) {
		BSTR* pTypes = nullptr;
		void* pCountsRaw = nullptr;
		VARTYPE vtCount;
		SafeArrayGetVartype(pLedCount, &vtCount);

		if (SUCCEEDED(SafeArrayAccessData(pDevType, (void**)&pTypes)) &&
			SUCCEEDED(SafeArrayAccessData(pLedCount, &pCountsRaw))) {

			long lBound, uBound;
			SafeArrayGetLBound(pDevType, 1, &lBound);
			SafeArrayGetUBound(pDevType, 1, &uBound);
			long count = uBound - lBound + 1;

			if (count > 0 && pTypes != nullptr && pTypes[0] != nullptr) {
				if (g_deviceName) SysFreeString(g_deviceName);
				g_deviceName = SysAllocString(pTypes[0]);
				g_totalLeds = GetIntFromSafeArray(pCountsRaw, vtCount, 0);

				BSTR friendlyName = NULL;
				if (lpMLAPI_GetDeviceNameEx) {
					lpMLAPI_GetDeviceNameEx(g_deviceName, 0, &friendlyName);
				}

				char devInfo[512];
				snprintf(devInfo, sizeof(devInfo), "[MysticFight] %ls (Type: %ls) | Logical Areas: %d",
					(friendlyName ? friendlyName : L"Unknown Device"),
					g_deviceName, g_totalLeds);
				Log(devInfo);
				if (friendlyName) SysFreeString(friendlyName);

				Log("[MysticFight] Listing styles per area...");
				for (DWORD i = 0; i < (DWORD)g_totalLeds; i++) {
					BSTR ledName = nullptr;
					SAFEARRAY* pStyles = nullptr;
					int resInfo = lpMLAPI_GetLedInfo(g_deviceName, i, &ledName, &pStyles);

					if (resInfo == 0) {
						char ledLine[512];
						snprintf(ledLine, sizeof(ledLine), "    [INDEX %lu] LED: %ls", i, (ledName ? ledName : L"Unknown"));
						Log(ledLine);

						if (pStyles) {
							long sLB, sUB;
							SafeArrayGetLBound(pStyles, 1, &sLB);
							SafeArrayGetUBound(pStyles, 1, &sUB);
							for (long k = sLB; k <= sUB; k++) {
								BSTR sName = nullptr;
								SafeArrayGetElement(pStyles, &k, &sName);
								if (sName) {
									char styleLine[256];
									snprintf(styleLine, sizeof(styleLine), "      |-- Style %ld: %ls", k, sName);
									Log(styleLine);
									SysFreeString(sName);
								}
							}
							SafeArrayDestroy(pStyles);
						}
						if (ledName) SysFreeString(ledName);
					}
				}
			}
			SafeArrayUnaccessData(pDevType);
			SafeArrayUnaccessData(pLedCount);
		}
		SafeArrayDestroy(pDevType);
		SafeArrayDestroy(pLedCount);
	}

	// --- I. CONEXIÓN A LIBREHARDWAREMONITOR ---
	Log("[MysticFight] Connecting to LibreHardwareMonitor...");
	bool lhmAlive = false;
	for (int j = 1; j <= 10; j++) {
		if (InitWMI()) {
			Log("[MysticFight] Connected to WMI namespace successfully.");
			lhmAlive = true;
			break;
		}
		char wmiRetry[100];
		snprintf(wmiRetry, sizeof(wmiRetry), "[MysticFight] WMI Connection attempt %d/10 failed.", j);
		Log(wmiRetry);

		if (j == 10) {
			Log("[MysticFight] FATAL: Could not connect to WMI.");
			MessageBox(NULL, L"Could not connect to LibreHardwareMonitor.\nPlease ensure it is open and 'WMI Server' is enabled.", L"WMI Error", MB_OK | MB_ICONWARNING);
			FinalCleanup(hWnd);
			return 1;
		}
		Sleep(5000);
	}

	DebugListAllSensors();

	// --- J. PREPARAR VARIABLES DEL BUCLE ---
	_bstr_t bstrOff(L"Off");
	_bstr_t bstrBreath(L"Breath");
	_bstr_t bstrSteady(L"Steady");

	PrepareLHMSensorWMIQuery();
	if (lpMLAPI_SetLedStyle) lpMLAPI_SetLedStyle(g_deviceName, 0, bstrSteady);

	DWORD nR = 0, nG = 0, nB = 0;
	lastR = 999; lastG = 999; lastB = 999;
	MSG msg = { 0 };

	// --- BUCLE PRINCIPAL (v2.17 - BLINDADO) ---
	while (g_Running) {
		ULONGLONG currentTime = GetTickCount64();

		// 1. PROCESAR MENSAJES
		while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
			if (msg.message == WM_QUIT) { g_Running = false; break; }
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}

		if (!g_Running) break;

		// 2. MÁQUINA DE ESTADOS DEL RESET (WATCHDOG)
		if (g_Resetting_sdk) {
			if (currentTime >= g_ResetTimer) {
				switch (g_ResetStage) {
				case 0:
					Log("[MysticFight] Reset Stage 0: Cleaning MSI...");
					ControlScheduledTask(L"MSI Task Host - LEDKeeper2_Host", false);
					system("taskkill /F /IM LEDKeeper2.exe /T > nul 2>&1");
					g_ResetTimer = currentTime + 1000; g_ResetStage = 1;
					break;
				case 1:
					Log("[MysticFight] Reset Stage 1: Restarting MSI...");
					ControlScheduledTask(L"MSI Task Host - LEDKeeper2_Host", true);
					g_ResetTimer = currentTime + 4500; g_ResetStage = 2;
					break;
				case 2:
					Log("[MysticFight] Reset Stage 2: Re-initializing SDK...");
					if (lpMLAPI_Initialize && lpMLAPI_Initialize() == 0) {
						Log("[MysticFight] SDK RE-INITIALIZED.");
						ShowNotification(hWnd, hInstance, L"MysticFight - Recovered", L"Hardware communication restored.");
						RegisterHotKey(hWnd, 1, MOD_CONTROL | MOD_SHIFT | MOD_ALT, 0x4C);
						if (lpMLAPI_SetLedStyle) lpMLAPI_SetLedStyle(g_deviceName, 0, bstrSteady);
						lastR = 999;
					}
					g_Resetting_sdk = false;
					break;
				}
			}
		}
		// 3. LÓGICA DE SENSORES Y HARDWARE
		else if (g_LedsEnabled) {
			float rawTemp = GetCPUTempFast();

			if (rawTemp < 0.0f) {
				if (lhmAlive) {
					Log("[MysticFight] LHM disconnected!");
					ShowNotification(hWnd, hInstance, L"MysticFight - Connection Lost", L"LHM monitoring is not responding.");
					if (lpMLAPI_SetLedStyle) lpMLAPI_SetLedStyle(g_deviceName, 0, bstrBreath);
					lhmAlive = false;
				}
				if (currentTime - lastRetry > 3000) { lastRetry = currentTime; InitWMI(); }
			}
			else {
				if (!lhmAlive) {
					lhmAlive = true; PrepareLHMSensorWMIQuery();
					if (lpMLAPI_SetLedStyle) lpMLAPI_SetLedStyle(g_deviceName, 0, bstrSteady);
					lastR = 999;
				}

				// CÁLCULO DE MEZCLA DE COLORES
				float temp = floorf(rawTemp * 4.0f + 0.5f) / 4.0f;
				float ratio = 0.0f;
				if (temp <= (float)g_cfg.tempLow) {
					nR = GetRValue(g_cfg.colorLow); nG = GetGValue(g_cfg.colorLow); nB = GetBValue(g_cfg.colorLow);
				}
				else if (temp <= (float)g_cfg.tempMed) {
					ratio = (temp - (float)g_cfg.tempLow) / ((float)g_cfg.tempMed - (float)g_cfg.tempLow);
					nR = GetRValue(g_cfg.colorLow) + (DWORD)(ratio * (int(GetRValue(g_cfg.colorMed)) - int(GetRValue(g_cfg.colorLow))));
					nG = GetGValue(g_cfg.colorLow) + (DWORD)(ratio * (int(GetGValue(g_cfg.colorMed)) - int(GetGValue(g_cfg.colorLow))));
					nB = GetBValue(g_cfg.colorLow) + (DWORD)(ratio * (int(GetBValue(g_cfg.colorMed)) - int(GetBValue(g_cfg.colorLow))));
				}
				else {
					ratio = (temp - (float)g_cfg.tempMed) / ((float)g_cfg.tempHigh - (float)g_cfg.tempMed);
					if (ratio > 1.0f) ratio = 1.0f;
					nR = GetRValue(g_cfg.colorMed) + (DWORD)(ratio * (int(GetRValue(g_cfg.colorHigh)) - int(GetRValue(g_cfg.colorMed))));
					nG = GetGValue(g_cfg.colorMed) + (DWORD)(ratio * (int(GetGValue(g_cfg.colorHigh)) - int(GetGValue(g_cfg.colorMed))));
					nB = GetBValue(g_cfg.colorMed) + (DWORD)(ratio * (int(GetBValue(g_cfg.colorHigh)) - int(GetBValue(g_cfg.colorMed))));
				}

				// --- ENVÍO AL HARDWARE (CRÍTICO) ---
				if (nR != lastR || nG != lastG || nB != lastB || (currentTime - lastForceRefresh > REFRESH_INTERVAL)) {
					int status = 0;

					// Si venimos de estar apagados (lastR == 1000) o forzados (999), ponemos STEADY primero
					if (lastR >= 999 || (currentTime - lastForceRefresh > REFRESH_INTERVAL)) {
						if (lpMLAPI_SetLedStyle) status = lpMLAPI_SetLedStyle(g_deviceName, 0, bstrSteady);
					}

					if (status == 0 && lpMLAPI_SetLedColor) {
						for (int i = 0; i < g_totalLeds; i++) {
							status = lpMLAPI_SetLedColor(g_deviceName, i, nR, nG, nB);
							if (status != 0) break;
						}
					}

					if (status != 0) {
						if (currentTime - g_LastResetAttempt > 15000) {
							Log("[MysticFight] SDK Call failed. Triggering Reset...");
							g_Resetting_sdk = true; g_ResetStage = 0; g_ResetTimer = 0;
							g_LastResetAttempt = currentTime;
						}
					}
					else {
						lastR = nR; lastG = nG; lastB = nB; lastForceRefresh = currentTime;
					}
				}
			}
		}
		else {
			// MODO APAGADO: Solo entramos si no estábamos ya en modo "1000" (Apagado)
			if (lastR != 1000 || (currentTime - lastForceRefresh > REFRESH_INTERVAL)) {
				if (lpMLAPI_SetLedStyle) lpMLAPI_SetLedStyle(g_deviceName, 0, bstrOff);
				lastR = 1000; // Centinela: Indica que el hardware está en modo OFF
				lastForceRefresh = currentTime;
			}
		}

		MsgWaitForMultipleObjects(0, NULL, FALSE, 500, QS_ALLINPUT);
	}

	// --- 4. LIMPIEZA FINAL ---
	FinalCleanup(hWnd);
	Log("[MysticFight] BYE BYE");

	return 0;
}