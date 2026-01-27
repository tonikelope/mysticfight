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
#include "resource.h"

#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Shell32.lib")

#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 1001
#define ID_TRAY_CONFIG 2001
#define ID_TRAY_LOG 3001

const wchar_t* APP_VERSION = L"v2.0";

struct Config {
	wchar_t sensorID[256];
	int tempLow, tempHigh, tempAlert;
};

Config g_cfg;
const wchar_t* INI_FILE = L".\\config.ini";

typedef int (*LPMLAPI_Initialize)();
typedef int (*LPMLAPI_GetDeviceInfo)(SAFEARRAY** pDevType, SAFEARRAY** pLedCount);
typedef int (*LPMLAPI_SetLedColor)(BSTR type, DWORD index, DWORD R, DWORD G, DWORD B);
typedef int (*LPMLAPI_SetLedStyle)(BSTR type, DWORD index, BSTR style);
typedef int (*LPMLAPI_SetLedSpeed)(BSTR type, DWORD index, DWORD level);

// SMART POINTERS DEFINITIONS
_COM_SMARTPTR_TYPEDEF(IWbemLocator, __uuidof(IWbemLocator));
_COM_SMARTPTR_TYPEDEF(IWbemServices, __uuidof(IWbemServices));
_COM_SMARTPTR_TYPEDEF(IEnumWbemClassObject, __uuidof(IEnumWbemClassObject));
_COM_SMARTPTR_TYPEDEF(IWbemClassObject, __uuidof(IWbemClassObject));
_COM_SMARTPTR_TYPEDEF(ITaskService, __uuidof(ITaskService));
_COM_SMARTPTR_TYPEDEF(ITaskFolder, __uuidof(ITaskFolder));
_COM_SMARTPTR_TYPEDEF(IRegisteredTask, __uuidof(IRegisteredTask));
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

IWbemServicesPtr g_pSvc = nullptr;
IWbemLocatorPtr g_pLoc = nullptr;

BSTR g_deviceName = NULL;
HMODULE g_hLibrary = NULL;
int g_totalLeds = 0;
LPMLAPI_SetLedColor lpMLAPI_SetLedColor = nullptr;
LPMLAPI_SetLedStyle lpMLAPI_SetLedStyle = nullptr;
LPMLAPI_SetLedSpeed lpMLAPI_SetLedSpeed = nullptr;
HANDLE g_hMutex = NULL;

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
	const char* filename = "debug_log.txt";

	time_t now = time(0);
	char dt[26];
	ctime_s(dt, sizeof(dt), &now);
	dt[24] = '\0';

	// Abrir en modo append (ios::app) es O(1), no depende del tamaño del archivo
	std::ofstream out(filename, std::ios::app);
	if (out) {
		out << "[" << dt << "] " << text << "\n";
	}
}

// Llama a esto SOLO UNA VEZ en WinMain al iniciar
static void TrimLogFile() {
	const char* filename = "debug_log.txt";
	std::vector<std::string> lines;
	std::string line;

	{
		std::ifstream in(filename);
		while (std::getline(in, line)) lines.push_back(line);
	}

	if (lines.size() > 500) {
		std::ofstream out(filename, std::ios::trunc);
		for (size_t i = lines.size() - 500; i < lines.size(); ++i) {
			out << lines[i] << "\n";
		}
	}
}

static void PrepareLHMSensorWMIQuery() {
	if (wcslen(g_cfg.sensorID) == 0) return;

	std::wstring q = L"SELECT Value FROM Sensor WHERE Identifier = '" + std::wstring(g_cfg.sensorID) + L"'";
	g_bstrQuery = q.c_str(); // El operador = de _bstr_t maneja la memoria por ti

	char sensorLog[512];
	snprintf(sensorLog, sizeof(sensorLog), "[MysticFight] Target Sensor updated to: %ls", g_cfg.sensorID);
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

// --- CONFIGURATION MANAGEMENT ---
void LoadSettings() {
	GetPrivateProfileStringW(L"Settings", L"SensorID", L"", g_cfg.sensorID, 256, INI_FILE);
	g_cfg.tempLow = GetPrivateProfileIntW(L"Settings", L"TempLow", 50, INI_FILE);
	g_cfg.tempHigh = GetPrivateProfileIntW(L"Settings", L"TempHigh", 70, INI_FILE);
	g_cfg.tempAlert = GetPrivateProfileIntW(L"Settings", L"TempAlert", 90, INI_FILE);

}

void SaveSettings() {
	WritePrivateProfileStringW(L"Settings", L"SensorID", g_cfg.sensorID, INI_FILE);
	WritePrivateProfileStringW(L"Settings", L"TempLow", std::to_wstring(g_cfg.tempLow).c_str(), INI_FILE);
	WritePrivateProfileStringW(L"Settings", L"TempHigh", std::to_wstring(g_cfg.tempHigh).c_str(), INI_FILE);
	WritePrivateProfileStringW(L"Settings", L"TempAlert", std::to_wstring(g_cfg.tempAlert).c_str(), INI_FILE);

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



// --- PROCEDIMIENTO DE DIÁLOGO DE CONFIGURACIÓN ---

INT_PTR CALLBACK SettingsDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
	switch (message) {
	case WM_INITDIALOG: {
		// 1. Estética: Iconos y centrar ventana
		HICON hIcon = (HICON)LoadImage(GetModuleHandle(NULL), MAKEINTRESOURCE(101), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
		if (hIcon) {
			SendMessage(hDlg, WM_SETICON, ICON_SMALL, (LPARAM)hIcon);
			SendMessage(hDlg, WM_SETICON, ICON_BIG, (LPARAM)hIcon);
		}

		RECT rcOwner, rcDlg;
		GetWindowRect(GetDesktopWindow(), &rcOwner);
		GetWindowRect(hDlg, &rcDlg);
		int x = ((rcOwner.right - rcOwner.left) - (rcDlg.right - rcDlg.left)) / 2;
		int y = ((rcOwner.bottom - rcOwner.top) - (rcDlg.bottom - rcDlg.top)) / 2;
		SetWindowPos(hDlg, HWND_TOP, x, y, 0, 0, SWP_NOSIZE);

		// 2. Rellenar lista y cargar valores actuales
		PopulateSensorList(hDlg);
		SetDlgItemInt(hDlg, IDC_TEMP_LOW, g_cfg.tempLow, TRUE);
		SetDlgItemInt(hDlg, IDC_TEMP_HIGH, g_cfg.tempHigh, TRUE);
		SetDlgItemInt(hDlg, IDC_TEMP_ALERT, g_cfg.tempAlert, TRUE);
		return (INT_PTR)TRUE;
	}

	case WM_COMMAND:
		switch (LOWORD(wParam)) {
		case IDOK: {
			HWND hCombo = GetDlgItem(hDlg, IDC_SENSOR_ID);
			int idx = (int)SendMessage(hCombo, CB_GETCURSEL, 0, 0);

			if (idx == CB_ERR) {
				MessageBoxW(hDlg, L"Please select a sensor to continue.", L"Configuration Error", MB_ICONWARNING);
				return TRUE;
			}

			// Sacamos el ID del sensor (el puntero al Heap)
			wchar_t* selectedID = (wchar_t*)SendMessage(hCombo, CB_GETITEMDATA, idx, 0);

			int tLow = GetDlgItemInt(hDlg, IDC_TEMP_LOW, NULL, TRUE);
			int tHigh = GetDlgItemInt(hDlg, IDC_TEMP_HIGH, NULL, TRUE);
			int tAlert = GetDlgItemInt(hDlg, IDC_TEMP_ALERT, NULL, TRUE);

			// Validación lógica
			if (tLow < 0 || tAlert > 110 || tLow >= tHigh || tHigh >= tAlert) {
				MessageBoxW(hDlg, L"Invalid temperature range. Order must be: Green < Yellow < Red.", L"Validation Error", MB_ICONWARNING);
				return TRUE;
			}

			// Guardar configuración
			if (selectedID && selectedID != (wchar_t*)CB_ERR) wcscpy_s(g_cfg.sensorID, selectedID);
			g_cfg.tempLow = tLow;
			g_cfg.tempHigh = tHigh;
			g_cfg.tempAlert = tAlert;

			SaveSettings();
			EndDialog(hDlg, IDOK);
			return (INT_PTR)TRUE;
		}

		case IDCANCEL:
			EndDialog(hDlg, IDCANCEL);
			return (INT_PTR)TRUE;
		}
		break;

	case WM_DESTROY:
		// --- AQUÍ ESTÁ LA MAGIA ---
		// Liberamos los strings del Heap antes de que el ComboBox desaparezca
		ClearComboHeapData(GetDlgItem(hDlg, IDC_SENSOR_ID));
		break;
	}
	return (INT_PTR)FALSE;
}


static void FinalCleanup(HWND hWnd) {
	static bool cleaned = false;
	if (cleaned) return;
	cleaned = true;

	Log("[MysticFight] Starting cleaning...");

	// 2. APAGADO DE HARDWARE (Prioridad Máxima)
	// Se hace mientras la DLL y los BSTRs siguen siendo válidos en memoria.
	if (g_hLibrary && g_deviceName) {
		// Volver al estilo estático y apagar todos los LEDs
		
		_bstr_t bstrSteady(L"Steady");

		lpMLAPI_SetLedStyle(g_deviceName, 0, bstrSteady);
		
		for (int i = 0; i < g_totalLeds; i++) {
			lpMLAPI_SetLedColor(g_deviceName, i, 0, 0, 0);
		}
		
		Log("[MysticFight] LEDs power off");
	}

	// 3. LIMPIEZA DE INTERFAZ DE USUARIO (Shell y Hotkeys)
	if (hWnd) {
		NOTIFYICONDATA nid = {};
		nid.cbSize = sizeof(NOTIFYICONDATA);
		nid.hWnd = hWnd;
		nid.uID = 1;
		Shell_NotifyIcon(NIM_DELETE, &nid);

		UnregisterHotKey(hWnd, 1);
	}

	// 4. LIBERACIÓN DE MEMORIA BSTR (MSI SDK)
	// Es vital poner a NULL después de liberar para evitar "Dangling Pointers"
	if (g_deviceName) {
		SysFreeString(g_deviceName);
		g_deviceName = NULL;
	}
	
	// 5. LIBERACIÓN DE LA LIBRERÍA (DLL)
	if (g_hLibrary) {
		FreeLibrary(g_hLibrary);
		g_hLibrary = NULL;
		// Limpiamos los punteros a funciones por seguridad
		lpMLAPI_SetLedColor = nullptr;
		lpMLAPI_SetLedStyle = nullptr;
	}

	// 6. DESTRUCCIÓN DE SMART POINTERS (CRÍTICO)
	g_pSvc = nullptr;
	g_pLoc = nullptr;

	// 7. SINCRONIZACIÓN Y COM
	if (g_hMutex) {
		ReleaseMutex(g_hMutex);
		CloseHandle(g_hMutex);
		g_hMutex = NULL;
	}

	// Último paso: Cerrar el entorno COM de Windows
	CoUninitialize();

	Log("[MysticFight] Cleaning finished.");
}

static LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
	switch (message) {
	case WM_TRAYICON:
		if (lParam == WM_RBUTTONUP) {
			POINT curPoint; GetCursorPos(&curPoint);
			HMENU hMenu = CreatePopupMenu();
			if (hMenu) {
				InsertMenu(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_CONFIG, L"Settings");
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
		if (LOWORD(wParam) == ID_TRAY_LOG) { // 👈 Acción para el Log
			ShellExecuteW(NULL, L"open", L"debug_log.txt", NULL, NULL, SW_SHOW);
		}
		// DENTRO DE WndProc -> switch (message) -> case WM_COMMAND
		if (LOWORD(wParam) == ID_TRAY_CONFIG) {
			if (DialogBoxW(GetModuleHandle(NULL), MAKEINTRESOURCE(IDD_SETTINGS), hWnd, SettingsDlgProc) == IDOK) {
				PrepareLHMSensorWMIQuery();
				Log("[MysticFight] Sensor updated.");
			}
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
	// Si no hay query preparada o no hay WMI, no hacemos nada
	if (!g_pSvc || g_bstrQuery.length() == 0) return 0.0f;

	IEnumWbemClassObjectPtr pEnum = nullptr;

	// USAMOS g_bstrQuery (YA CACHEADA)
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
	return 0.0f;
}

void DebugListAllSensors() {
	if (!g_pSvc) return;
	IEnumWbemClassObjectPtr pEnum = nullptr;
	HRESULT hr = g_pSvc->ExecQuery(_bstr_t(L"WQL"), _bstr_t(L"SELECT Identifier, Name FROM Sensor WHERE SensorType='Temperature'"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnum);

	if (SUCCEEDED(hr)) {
		IWbemClassObjectPtr pObj = nullptr;
		ULONG uRet = 0;
		Log("AVAILABLE SENSORS FROM LHM:");
		while (pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet) == S_OK) {
			_variant_t vtID, vtName;
			pObj->Get(L"Identifier", 0, &vtID, 0, 0);
			pObj->Get(L"Name", 0, &vtName, 0, 0);
			char buf[512];
			snprintf(buf, sizeof(buf), "ID: %ls | Nombre: %ls", vtID.bstrVal, vtName.bstrVal);
			Log(buf);
		}
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


int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	SetPriorityClass(GetCurrentProcess(), HIGH_PRIORITY_CLASS);
	SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);
	
	HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

	TrimLogFile();
	
	if (FAILED(hr)) {
		// Critical error: cannot use WMI or Task Scheduler without COM
		char errorMsg[256];
		snprintf(errorMsg, sizeof(errorMsg), "[MysticFight] Failed to initialize COM: HRESULT 0x%08X", hr);
		Log(errorMsg);

		MessageBoxW(NULL, L"Critical Error: Could not initialize Windows COM system.", L"MysticFight - Error", MB_OK | MB_ICONERROR);
		return 1;
	}

	char versionMsg[128];
	snprintf(versionMsg, sizeof(versionMsg), "[MysticFight] Starting MysticFight %ls application...", APP_VERSION);
	Log(versionMsg);

	LoadSettings();

	char startupCfg[512];
	snprintf(startupCfg, sizeof(startupCfg),
		"[MysticFight] Config Loaded - Sensor: %ls | Low: %d | High: %d | Alert: %d",
		g_cfg.sensorID, g_cfg.tempLow, g_cfg.tempHigh, g_cfg.tempAlert);
	Log(startupCfg);

	PrepareLHMSensorWMIQuery();

	g_hMutex = CreateMutex(NULL, TRUE, L"Global\\MysticFight_Unique_Mutex");
	if (g_hMutex == NULL || GetLastError() == ERROR_ALREADY_EXISTS) {
		if (g_hMutex) {
			MessageBox(NULL, L"Error: An instance of MysticFight is already running.", L"Error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
			CloseHandle(g_hMutex);
		}
		CoUninitialize();
		return 0;
	}

	// --- 2. VERIFICACIÓN DE AUTO-INICIO (Task Scheduler) ---
	// Comprobamos si la tarea existe y si la ruta coincide
	// --- STARTUP TASK CHECK ---
	if (!IsTaskValid()) {
		int msgBoxID = MessageBoxW(NULL,
			L"MysticFight is not configured to run at startup (or the executable path has changed).\n\n"
			L"Would you like to create a scheduled task to start automatically?",
			L"Startup Configuration - MysticFight",
			MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON1 | MB_SETFOREGROUND);

		if (msgBoxID == IDYES) {
			SetRunAtStartup(true); // Your COM-based function
		}
	}

	if (wcslen(g_cfg.sensorID) == 0) {
		if (DialogBoxW(hInstance, MAKEINTRESOURCE(IDD_SETTINGS), NULL, SettingsDlgProc) == IDCANCEL) {
			if (g_hMutex) { ReleaseMutex(g_hMutex); CloseHandle(g_hMutex); }
			CoUninitialize();
			return 0;
		}
	}
	wchar_t windowTitle[100];
	swprintf_s(windowTitle, L"MysticFight %s (by tonikelope)", APP_VERSION);
	WNDCLASS wc = { 0 };
	wc.lpfnWndProc = WndProc;
	wc.hInstance = hInstance;
	wc.lpszClassName = L"MysticFight_Class";
	wc.hIcon = (HICON)LoadImage(hInstance, MAKEINTRESOURCE(101), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
	RegisterClass(&wc);
	HWND hWnd = CreateWindowEx(0, wc.lpszClassName, windowTitle, 0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);

	g_hLibrary = LoadLibrary(L"MysticLight_SDK.dll");
	if (!g_hLibrary) {
		Log("[MysticFight] FATAL: Could not find 'MysticLight_SDK.dll'.");
		MessageBox(NULL, L"Could not find 'MysticLight_SDK.dll'.", L"Error", MB_OK | MB_ICONERROR);
		FinalCleanup(hWnd);
		return 1;
	}

	auto fnMLAPI_Initialize = (LPMLAPI_Initialize)GetProcAddress(g_hLibrary, "MLAPI_Initialize");
	auto fnMLAPI_GetDeviceInfo = (LPMLAPI_GetDeviceInfo)GetProcAddress(g_hLibrary, "MLAPI_GetDeviceInfo");
	lpMLAPI_SetLedColor = (LPMLAPI_SetLedColor)GetProcAddress(g_hLibrary, "MLAPI_SetLedColor");
	lpMLAPI_SetLedStyle = (LPMLAPI_SetLedStyle)GetProcAddress(g_hLibrary, "MLAPI_SetLedStyle");
	lpMLAPI_SetLedSpeed = (LPMLAPI_SetLedSpeed)GetProcAddress(g_hLibrary, "MLAPI_SetLedSpeed");

	Log("[MysticFight] Attempting to initialize SDK...");
	if (!fnMLAPI_Initialize || fnMLAPI_Initialize() != 0) {
		bool initialized = false;
		for (int i = 1; i <= 10; i++) {
			char retryMsg[100];
			snprintf(retryMsg, sizeof(retryMsg), "[MysticFight] Attempt %d/10 failed. Retrying in 5s...", i);
			Log(retryMsg);

			if (fnMLAPI_Initialize && fnMLAPI_Initialize() == 0) {
				Log("[MysticFight] SDK Initialized successfully on retry.");
				initialized = true;
				break;
			}

			if (i == 10) {
				Log("[MysticFight] FATAL: SDK could not be initialized after 10 attempts.");
				MessageBox(NULL, L"MSI SDK Initialization failed. Ensure MSI Center/Dragon Center is installed and service is running.", L"Critical Error", MB_OK | MB_ICONERROR);
				FinalCleanup(hWnd);
				return 1;
			}
			Sleep(5000);
		}
	}
	else {
		Log("[MysticFight] SDK Initialized successfully at first attempt.");
	}

	// --- REEMPLAZO SEGURO ---
	SAFEARRAY* pDevType = nullptr;
	SAFEARRAY* pLedCount = nullptr;

	if (fnMLAPI_GetDeviceInfo && fnMLAPI_GetDeviceInfo(&pDevType, &pLedCount) == 0 && pDevType && pLedCount) {
		BSTR* pNames = nullptr;
		void* pCountsRaw = nullptr;
		VARTYPE vtCount;

		// Obtenemos el tipo real del array de conteo
		SafeArrayGetVartype(pLedCount, &vtCount);

		if (SUCCEEDED(SafeArrayAccessData(pDevType, (void**)&pNames)) &&
			SUCCEEDED(SafeArrayAccessData(pLedCount, &pCountsRaw))) {

			long lBound, uBound;
			SafeArrayGetLBound(pDevType, 1, &lBound);
			SafeArrayGetUBound(pDevType, 1, &uBound);
			long count = uBound - lBound + 1;

			if (count > 0) {
				if (g_deviceName) SysFreeString(g_deviceName);
				
				g_deviceName = SysAllocString(pNames[0]);

				g_totalLeds = GetIntFromSafeArray(pCountsRaw, vtCount, 0);
				
				char devInfo[256];
				
				snprintf(devInfo, sizeof(devInfo), "[MysticFight] Device: %ls | LEDs: %d",
					g_deviceName, g_totalLeds);
				
				Log(devInfo);
			}

			SafeArrayUnaccessData(pDevType);
			SafeArrayUnaccessData(pLedCount);
		}
		SafeArrayDestroy(pDevType);
		SafeArrayDestroy(pLedCount);
	}

	Log("[MysticFight] Connecting to LibreHardwareMonitor...");

	bool wmiConnected = false;
	for (int j = 1; j <= 10; j++) {
		if (InitWMI()) {
			Log("[MysticFight] Connected to WMI namespace successfully.");
			wmiConnected = true;
			break;
		}
		char wmiRetry[100];
		snprintf(wmiRetry, sizeof(wmiRetry), "[MysticFight] WMI Connection attempt %d/10 failed.", j);
		Log(wmiRetry);

		if (j == 10) {
			Log("[MysticFight] FATAL: Could not connect to WMI. Is LibreHardwareMonitor running?");
			MessageBox(NULL, L"Could not connect to LibreHardwareMonitor.\nPlease ensure it is open and 'WMI Server' is enabled.", L"WMI Error", MB_OK | MB_ICONWARNING);
			FinalCleanup(hWnd);
			return 1;
		}
		Sleep(5000);
	}

	DebugListAllSensors();

	RegisterHotKey(hWnd, 1, MOD_CONTROL | MOD_SHIFT | MOD_ALT, 0x4C);

	NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA), hWnd, 1, NIF_ICON | NIF_MESSAGE | NIF_TIP, WM_TRAYICON };
	nid.hIcon = (HICON)LoadImage(hInstance, MAKEINTRESOURCE(101), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
	swprintf_s(nid.szTip, L"MysticFight %s (by tonikelope)", APP_VERSION);
	Shell_NotifyIcon(NIM_ADD, &nid);

	ShowNotification(hWnd, hInstance, windowTitle, L"Let's dance baby");

	_bstr_t bstrSteady(L"Steady");

	DWORD R = 0, G = 0, B = 0;
	DWORD lastR = 999, lastG = 999;
	int wmiRetryCounter = 0;
	const int WMI_RETRY_DELAY = 10;
	MSG msg = { 0 };

	while (g_Running) {
		// 1. PROCESAR MENSAJES
		while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
			if (msg.message == WM_QUIT) { g_Running = false; break; }
			if (msg.message == WM_HOTKEY) {
				g_LedsEnabled = !g_LedsEnabled;
				lastR = 999; lastG = 999; // Forzar actualización de color
				if (g_LedsEnabled) MessageBeep(MB_OK);
				else {
					MessageBeep(MB_ICONHAND);
					
					if (lpMLAPI_SetLedStyle && g_deviceName) {

						lpMLAPI_SetLedStyle(g_deviceName, 0, bstrSteady);

						for (int i = 0; i < g_totalLeds; i++)
							if (lpMLAPI_SetLedColor) lpMLAPI_SetLedColor(g_deviceName, i, 0, 0, 0);
					}
				}
			}
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}

		if (!g_Running) break;

		// 2. LÓGICA DE HARDWARE
		if (g_LedsEnabled) {
			float rawTemp = GetCPUTempFast();

			if (rawTemp > 0.0f) {
				
				// Lógica de colores normal
				float temp = floorf(rawTemp * 4.0f + 0.5f) / 4.0f;
				float ratio;
				
				if (temp <= (float)g_cfg.tempLow) {
					R = 0; G = 255; B = 0;
				}
				else if (temp <= (float)g_cfg.tempHigh) {
					float divisor = (float)g_cfg.tempHigh - (float)g_cfg.tempLow;
					ratio = (temp - (float)g_cfg.tempLow) / divisor;
					if (ratio > 1.0f) ratio = 1.0f; // Seguridad
					R = (DWORD)(255 * ratio); G = 255; B = 0;
				}
				else {
					float divisor = (float)g_cfg.tempAlert - (float)g_cfg.tempHigh;
					ratio = (temp - (float)g_cfg.tempHigh) / divisor;

					// Si ratio es 1.0 o más, ya estamos en zona de alerta máxima
					if (ratio >= 1.0f) {
						R = 255; G = 0; B = 0;
					}
					else {
						R = 255; G = (DWORD)(255 * (1.0f - ratio)); B = 0;
					}
				}

				if (R != lastR || G != lastG) {
					
					lpMLAPI_SetLedStyle(g_deviceName, 0, bstrSteady);
					
					for (int i = 0; i < g_totalLeds; i++) {
						if (lpMLAPI_SetLedColor) lpMLAPI_SetLedColor(g_deviceName, i, R, G, B);
					}

					lastR = R; lastG = G;
				}
			}
		}

		MsgWaitForMultipleObjects(0, NULL, FALSE, 500, QS_ALLINPUT);
	}

	FinalCleanup(hWnd);

	Log("[MysticFight] BYE BYE");

	return 0;
}
