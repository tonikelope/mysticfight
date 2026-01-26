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
#include "resource.h"

#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Shell32.lib")

#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 1001
#define ID_TRAY_CONFIG 2001

const wchar_t* APP_VERSION = L"v1.1";

// CONFIG STRUCTURE
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

bool g_Running = true;
bool g_LedsEnabled = true;
IWbemServices* g_pSvc = NULL;
IEnumWbemClassObject* g_pEnumTemperatura = NULL;
IWbemLocator* g_pLoc = NULL;
BSTR g_deviceName = NULL;
BSTR g_styleSteady = NULL;
HMODULE g_hLibrary = NULL;
int g_totalLeds = 0;
LPMLAPI_SetLedColor lpMLAPI_SetLedColor = nullptr;
LPMLAPI_SetLedStyle lpMLAPI_SetLedStyle = nullptr;
HANDLE g_hMutex = NULL;

static bool InitWMI();

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

	// 1. Limpieza GARANTIZADA de la memoria anterior
	int count = (int)SendMessage(hCombo, CB_GETCOUNT, 0, 0);
	for (int i = 0; i < count; i++) {
		wchar_t* ptr = (wchar_t*)SendMessage(hCombo, CB_GETITEMDATA, i, 0);
		if (ptr && ptr != (wchar_t*)CB_ERR) {
			free(ptr); // Liberamos el string duplicado
		}
	}
	SendMessage(hCombo, CB_RESETCONTENT, 0, 0);

	if (!g_pSvc) InitWMI();

	if (g_pSvc) {
		IEnumWbemClassObject* pEnumerator = NULL;
		// Optimizamos la query pidiendo solo lo estrictamente necesario
		bstr_t query("SELECT Name, Identifier FROM Sensor WHERE SensorType = 'Temperature'");

		if (SUCCEEDED(g_pSvc->ExecQuery(bstr_t("WQL"), query,
			WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator))) {

			IWbemClassObject* pclsObj = NULL;
			ULONG uReturn = 0;
			bool found = false;

			while (pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn) == S_OK) {
				VARIANT vtName, vtID;
				pclsObj->Get(L"Name", 0, &vtName, 0, 0);
				pclsObj->Get(L"Identifier", 0, &vtID, 0, 0);

				int idx = (int)SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)vtName.bstrVal);
				if (idx != CB_ERR) {
					// Guardamos el Identifier oculto en el item
					wchar_t* persistentID = _wcsdup(vtID.bstrVal);
					SendMessage(hCombo, CB_SETITEMDATA, idx, (LPARAM)persistentID);

					// Auto-seleccionar el sensor actual si coincide
					if (wcscmp(vtID.bstrVal, g_cfg.sensorID) == 0) {
						SendMessage(hCombo, CB_SETCURSEL, idx, 0);
						found = true;
					}
				}

				VariantClear(&vtName);
				VariantClear(&vtID);
				pclsObj->Release();
			}
			pEnumerator->Release();

			if (!found && SendMessage(hCombo, CB_GETCOUNT, 0, 0) > 0) {
				SendMessage(hCombo, CB_SETCURSEL, 0, 0);
			}
		}
	}

	// Manejo de error si no hay sensores
	if (SendMessage(hCombo, CB_GETCOUNT, 0, 0) == 0) {
		SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)L"Error: No sensors found. Open LHM!");
		EnableWindow(GetDlgItem(hDlg, IDOK), FALSE);
	}
}

INT_PTR CALLBACK SettingsDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
	switch (message) {
	case WM_INITDIALOG: {
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

		PopulateSensorList(hDlg);

		SetDlgItemInt(hDlg, IDC_TEMP_LOW, g_cfg.tempLow, TRUE);
		SetDlgItemInt(hDlg, IDC_TEMP_HIGH, g_cfg.tempHigh, TRUE);
		SetDlgItemInt(hDlg, IDC_TEMP_ALERT, g_cfg.tempAlert, TRUE);
		return (INT_PTR)TRUE;
	}
	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK) {
			int idx = (int)SendMessage(GetDlgItem(hDlg, IDC_SENSOR_ID), CB_GETCURSEL, 0, 0);
			if (idx == CB_ERR) {
				MessageBoxW(hDlg, L"Please select a sensor from the list.", L"Error", MB_ICONWARNING);
				return TRUE;
			}

			wchar_t* selectedID = (wchar_t*)SendMessage(GetDlgItem(hDlg, IDC_SENSOR_ID), CB_GETITEMDATA, idx, 0);
			int tLow = GetDlgItemInt(hDlg, IDC_TEMP_LOW, NULL, TRUE);
			int tHigh = GetDlgItemInt(hDlg, IDC_TEMP_HIGH, NULL, TRUE);
			int tAlert = GetDlgItemInt(hDlg, IDC_TEMP_ALERT, NULL, TRUE);

			if (tLow < 0 || tAlert > 150 || tLow >= tHigh || tHigh >= tAlert) {
				MessageBoxW(hDlg, L"Invalid values or illogical order (0-150).\nEnsure: Green < Yellow < Red.", L"Error", MB_ICONWARNING);
				return TRUE;
			}

			if (selectedID && selectedID != (wchar_t*)CB_ERR) {
				wcscpy_s(g_cfg.sensorID, selectedID);
			}
			g_cfg.tempLow = tLow;
			g_cfg.tempHigh = tHigh;
			g_cfg.tempAlert = tAlert;

			SaveSettings();
			EndDialog(hDlg, IDOK);
			return TRUE;
		}
		if (LOWORD(wParam) == IDCANCEL) { EndDialog(hDlg, IDCANCEL); return TRUE; }
		break;

	case WM_DESTROY: {
		HWND hCombo = GetDlgItem(hDlg, IDC_SENSOR_ID);
		int count = (int)SendMessage(hCombo, CB_GETCOUNT, 0, 0);
		for (int i = 0; i < count; i++) {
			wchar_t* ptr = (wchar_t*)SendMessage(hCombo, CB_GETITEMDATA, i, 0);
			if (ptr && ptr != (wchar_t*)CB_ERR) free(ptr);
		}
		break;
	}
	}
	return FALSE;
}

static void Log(const char* text) {
	const char* filename = "debug_log.txt";
	std::vector<std::string> lines;
	{
		std::ifstream in(filename);
		std::string line;
		while (std::getline(in, line)) lines.push_back(line);
	}
	time_t now = time(0);
	char dt[26];
	ctime_s(dt, sizeof(dt), &now);
	dt[24] = '\0';
	char buffer[1024];
	snprintf(buffer, sizeof(buffer), "[%s] %s", dt, text);
	lines.push_back(buffer);
	if (lines.size() > 500) lines.erase(lines.begin(), lines.begin() + (lines.size() - 500));
	std::ofstream out(filename, std::ios::trunc);
	if (!out) return;
	for (const auto& l : lines) out << l << "\n";
}

static void FinalCleanup(HWND hWnd) {
	static bool cleaned = false;
	if (cleaned) return;
	cleaned = true;
	Log("[MysticFight] Cleaning resources...");

	// 1. Apagamos LEDs (mientras la DLL sigue cargada)
	if (lpMLAPI_SetLedStyle && g_deviceName) {
		lpMLAPI_SetLedStyle(g_deviceName, 0, g_styleSteady);
		for (int i = 0; i < g_totalLeds; i++) {
			if (lpMLAPI_SetLedColor) lpMLAPI_SetLedColor(g_deviceName, i, 0, 0, 0);
		}
	}

	// 2. Liberamos Icono y Hotkeys
	NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA) };
	nid.hWnd = hWnd; nid.uID = 1;
	Shell_NotifyIcon(NIM_DELETE, &nid);
	if (hWnd) UnregisterHotKey(hWnd, 1);

	// 3. Liberamos Strings de MSI
	if (g_deviceName) SysFreeString(g_deviceName);
	if (g_styleSteady) SysFreeString(g_styleSteady);


	// 4. Liberamos la DLL de MSI (Ya no la necesitamos)
	if (g_hLibrary) {
		FreeLibrary(g_hLibrary);
		g_hLibrary = NULL;
	}

	// 5. Liberamos WMI (Antes de apagar COM)
	if (g_pEnumTemperatura) { g_pEnumTemperatura->Release(); g_pEnumTemperatura = NULL; }
	if (g_pSvc) { g_pSvc->Release(); g_pSvc = NULL; }
	if (g_pLoc) { g_pLoc->Release(); g_pLoc = NULL; }

	// 6. Cerramos Mutex
	if (g_hMutex) { ReleaseMutex(g_hMutex); CloseHandle(g_hMutex); g_hMutex = NULL; }

	// 7. APAGAR COM (EL ÚLTIMO)
	CoUninitialize();

	Log("[MysticFight] BYE BYE (Cleanup finished)");
}

static LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
	switch (message) {
	case WM_TRAYICON:
		if (lParam == WM_RBUTTONUP) {
			POINT curPoint; GetCursorPos(&curPoint);
			HMENU hMenu = CreatePopupMenu();
			if (hMenu) {
				InsertMenu(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_CONFIG, L"Settings");
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
		if (LOWORD(wParam) == ID_TRAY_CONFIG) {
			if (DialogBoxW(GetModuleHandle(NULL), MAKEINTRESOURCE(IDD_SETTINGS), hWnd, SettingsDlgProc) == IDOK) {

				if (g_pEnumTemperatura) {
					g_pEnumTemperatura->Release();
					g_pEnumTemperatura = NULL;
				}
				Log("[WMI] Sensor actualizado, reiniciando consulta...");
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

static bool InitWMI() {
	if (g_pSvc) { g_pSvc->Release(); g_pSvc = NULL; }
	if (g_pLoc) { g_pLoc->Release(); g_pLoc = NULL; }
	if (FAILED(CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&g_pLoc))) return false;
	if (FAILED(g_pLoc->ConnectServer(_bstr_t(L"ROOT\\LibreHardwareMonitor"), NULL, NULL, 0, NULL, 0, 0, &g_pSvc))) return false;
	return (SUCCEEDED(CoSetProxyBlanket(g_pSvc, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, NULL, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE)));
}

static float GetCPUTempFast() {
	if (!g_pSvc) return 0.0f;

	// 1. Si el enumerador es NULL (primera vez o tras cambiar sensor), lo creamos
	if (g_pEnumTemperatura == NULL) {
		std::wstring queryStr = L"SELECT Value FROM Sensor WHERE Identifier = '" + std::wstring(g_cfg.sensorID) + L"'";
		bstr_t bstrQuery(queryStr.c_str());
		bstr_t bstrWQL("WQL");

		// Usamos WBEM_FLAG_RETURN_IMMEDIATELY para que no bloquee
		HRESULT hr = g_pSvc->ExecQuery(bstrWQL, bstrQuery,
			WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
			NULL, &g_pEnumTemperatura);

		if (FAILED(hr)) return 0.0f;
	}

	float temp = 0.0f;
	IWbemClassObject* pclsObj = NULL;
	ULONG uReturn = 0;

	// 2. En lugar de hacer una Query nueva, reseteamos el que ya tenemos
	// Esto es lo que ahorra un 90% de CPU
	g_pEnumTemperatura->Reset();

	if (g_pEnumTemperatura->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn) == S_OK && uReturn > 0) {
		VARIANT vtProp;
		if (SUCCEEDED(pclsObj->Get(L"Value", 0, &vtProp, 0, 0))) {
			if (vtProp.vt == VT_R4) temp = vtProp.fltVal;
			else if (vtProp.vt == VT_R8) temp = (float)vtProp.dblVal;
			VariantClear(&vtProp);
		}
		pclsObj->Release();
	}
	else {
		// Si falla el Next, liberamos para que se reintente la Query en el próximo ciclo
		g_pEnumTemperatura->Release();
		g_pEnumTemperatura = NULL;
	}
	return temp;
}

static void ShowNotification(HWND hWnd, HINSTANCE hInstance, const wchar_t* title, const wchar_t* info) {
	NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA), hWnd, 1, NIF_INFO };
	wcscpy_s(nid.szInfoTitle, title);
	wcscpy_s(nid.szInfo, info);
	nid.dwInfoFlags = NIIF_USER | NIIF_LARGE_ICON;
	Shell_NotifyIcon(NIM_MODIFY, &nid);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	SetPriorityClass(GetCurrentProcess(), HIGH_PRIORITY_CLASS);
	SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);
	CoInitializeEx(0, COINIT_MULTITHREADED);
	LoadSettings();

	g_hMutex = CreateMutex(NULL, TRUE, L"Global\\MysticFight_Unique_Mutex");
	if (g_hMutex == NULL || GetLastError() == ERROR_ALREADY_EXISTS) {
		if (g_hMutex) {
			MessageBox(NULL, L"Error: An instance of MysticFight is already running.", L"Error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
			CloseHandle(g_hMutex);
		}
		CoUninitialize();
		return 0;
	}

	if (wcslen(g_cfg.sensorID) == 0) {
		if (DialogBoxW(hInstance, MAKEINTRESOURCE(IDD_SETTINGS), NULL, SettingsDlgProc) == IDCANCEL) {
			if (g_hMutex) { ReleaseMutex(g_hMutex); CloseHandle(g_hMutex); }
			CoUninitialize();
			return 0;
		}
	}

	Log("[MysticFight] Starting application...");
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
		Log("[MysticLight] FATAL: Could not find 'MysticLight_SDK.dll'.");
		MessageBox(NULL, L"Could not find 'MysticLight_SDK.dll'.", L"Error", MB_OK | MB_ICONERROR);
		FinalCleanup(hWnd);
		return 1;
	}

	auto fnMLAPI_Initialize = (LPMLAPI_Initialize)GetProcAddress(g_hLibrary, "MLAPI_Initialize");
	auto fnMLAPI_GetDeviceInfo = (LPMLAPI_GetDeviceInfo)GetProcAddress(g_hLibrary, "MLAPI_GetDeviceInfo");
	lpMLAPI_SetLedColor = (LPMLAPI_SetLedColor)GetProcAddress(g_hLibrary, "MLAPI_SetLedColor");
	lpMLAPI_SetLedStyle = (LPMLAPI_SetLedStyle)GetProcAddress(g_hLibrary, "MLAPI_SetLedStyle");

	Log("[MysticLight] Attempting to initialize SDK...");
	if (!fnMLAPI_Initialize || fnMLAPI_Initialize() != 0) {
		bool initialized = false;
		for (int i = 1; i <= 10; i++) {
			char retryMsg[100];
			snprintf(retryMsg, sizeof(retryMsg), "[MysticLight] Attempt %d/10 failed. Retrying in 5s...", i);
			Log(retryMsg);

			if (fnMLAPI_Initialize && fnMLAPI_Initialize() == 0) {
				Log("[MysticLight] SDK Initialized successfully on retry.");
				initialized = true;
				break;
			}

			if (i == 10) {
				Log("[MysticLight] FATAL: SDK could not be initialized after 10 attempts.");
				MessageBox(NULL, L"MSI SDK Initialization failed. Ensure MSI Center/Dragon Center is installed and service is running.", L"Critical Error", MB_OK | MB_ICONERROR);
				FinalCleanup(hWnd);
				return 1;
			}
			Sleep(5000);
		}
	}
	else {
		Log("[MysticLight] SDK Initialized successfully at first attempt.");
	}

	// --- REEMPLAZO SEGURO ---
	SAFEARRAY* pDevType = nullptr;
	SAFEARRAY* pLedCount = nullptr;

	if (fnMLAPI_GetDeviceInfo && fnMLAPI_GetDeviceInfo(&pDevType, &pLedCount) == 0 && pDevType && pLedCount) {
		BSTR* pNames = nullptr;
		// IMPORTANTE: Verifica si el SDK devuelve BSTR o LONG para los counts. 
		// En MSI suele ser BSTR, así que mantenemos BSTR* para pCounts
		BSTR* pCounts = nullptr;

		// Bloqueamos los arrays para leerlos con seguridad
		if (SUCCEEDED(SafeArrayAccessData(pDevType, (void**)&pNames)) &&
			SUCCEEDED(SafeArrayAccessData(pLedCount, (void**)&pCounts))) {

			long lBound, uBound;
			SafeArrayGetLBound(pDevType, 1, &lBound);
			SafeArrayGetUBound(pDevType, 1, &uBound);
			long count = uBound - lBound + 1;

			if (count > 0) {
				// Liberamos si ya existía algo para evitar fugas
				if (g_deviceName) SysFreeString(g_deviceName);

				g_deviceName = SysAllocString(pNames[0]);
				g_totalLeds = _wtoi(pCounts[0]); // El SDK de MSI entrega el número como texto

				char devInfo[256];
				snprintf(devInfo, sizeof(devInfo), "[MysticLight] Device: %ls | LEDs: %d", g_deviceName, g_totalLeds);
				Log(devInfo);
			}

			// Obligatorio: Desbloquear siempre
			SafeArrayUnaccessData(pDevType);
			SafeArrayUnaccessData(pLedCount);
		}

		// Liberar los SAFEARRAY que creó el SDK
		SafeArrayDestroy(pDevType);
		SafeArrayDestroy(pLedCount);
	}

	g_styleSteady = SysAllocString(L"Steady");


	Log("[WMI] Connecting to LibreHardwareMonitor...");
	bool wmiConnected = false;
	for (int j = 1; j <= 10; j++) {
		if (InitWMI()) {
			Log("[WMI] Connected to namespace successfully.");
			wmiConnected = true;
			break;
		}
		char wmiRetry[100];
		snprintf(wmiRetry, sizeof(wmiRetry), "[WMI] Connection attempt %d/10 failed.", j);
		Log(wmiRetry);

		if (j == 10) {
			Log("[WMI] FATAL: Could not connect to WMI. Is LibreHardwareMonitor running?");
			MessageBox(NULL, L"Could not connect to LibreHardwareMonitor.\nPlease ensure it is open and 'WMI Server' is enabled.", L"WMI Error", MB_OK | MB_ICONWARNING);
			FinalCleanup(hWnd);
			return 1;
		}
		Sleep(5000);
	}

	RegisterHotKey(hWnd, 1, MOD_CONTROL | MOD_SHIFT | MOD_ALT, 0x4C);

	NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA), hWnd, 1, NIF_ICON | NIF_MESSAGE | NIF_TIP, WM_TRAYICON };
	nid.hIcon = (HICON)LoadImage(hInstance, MAKEINTRESOURCE(101), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
	swprintf_s(nid.szTip, L"MysticFight %s (by tonikelope)", APP_VERSION);
	Shell_NotifyIcon(NIM_ADD, &nid);

	ShowNotification(hWnd, hInstance, windowTitle, L"Let's dance baby");

	DWORD R = 0, G = 0, B = 0;
	DWORD lastR = 999, lastG = 999;
	
	MSG msg = { 0 };

	while (g_Running) {
		// 1. PROCESAR MENSAJES (Interfaz de usuario siempre fluida)
		while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
			if (msg.message == WM_QUIT) {
				g_Running = false;
				break;
			}
			if (msg.message == WM_HOTKEY) {
				g_LedsEnabled = !g_LedsEnabled;
				
				lastR = 999; lastG = 999;

				if (g_LedsEnabled) {
					MessageBeep(MB_OK);
					for (int i = 0; i < g_totalLeds; i++)
						if (lpMLAPI_SetLedColor) lpMLAPI_SetLedColor(g_deviceName, i, R, G, B);
				}
				else {
					MessageBeep(MB_ICONHAND);
					for (int i = 0; i < g_totalLeds; i++)
						if (lpMLAPI_SetLedColor) lpMLAPI_SetLedColor(g_deviceName, i, 0, 0, 0);
				}
			}
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}

		if (!g_Running) break;

		// 2. LÓGICA DE HARDWARE (Solo si los LEDs están activos)
		if (g_LedsEnabled) {
			float rawTemp = GetCPUTempFast();

			if (rawTemp <= 0) {
				InitWMI(); // Intento silencioso de reconexión
			}
			else {
				float temp = floorf(rawTemp * 4.0f + 0.5f) / 4.0f;

					if (temp <= (float)g_cfg.tempLow) {
						R = 0; G = 255; B = 0;
					}
					else if (temp <= (float)g_cfg.tempHigh) {
						// SEGURIDAD: Evitar división por cero
						float divisor = (float)g_cfg.tempHigh - (float)g_cfg.tempLow;
						float ratio = (divisor > 0) ? (temp - (float)g_cfg.tempLow) / divisor : 1.0f;
						R = (DWORD)(255 * ratio);
						G = 255; B = 0;
					}
					else {
						// SEGURIDAD: Evitar división por cero
						float divisor = (float)g_cfg.tempAlert - (float)g_cfg.tempHigh;
						float ratio = (divisor > 0) ? (temp - (float)g_cfg.tempHigh) / divisor : 1.0f;
						R = 255;
						G = (DWORD)(255 * (1.0f - ratio));
						B = 0;
					}
				

				// Solo actualizar el SDK si el color realmente cambió
				if (R != lastR || G != lastG) {
					for (int i = 0; i < g_totalLeds; i++) {
						if (lpMLAPI_SetLedColor) lpMLAPI_SetLedColor(g_deviceName, i, R, G, B);
					}
					lastR = R; lastG = G;
				}
			}
		}

		// 3. LA ESPERA INTELIGENTE
		// Espera 500ms O hasta que el usuario mueva el ratón/haga clic/use el teclado
		MsgWaitForMultipleObjects(0, NULL, FALSE, 500, QS_ALLINPUT);
	}

	FinalCleanup(hWnd);
	return 0;
}
