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
#include <atomic>
#include "resource.h"

#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Shell32.lib")

#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 1001
#define ID_TRAY_CONFIG 2001

const wchar_t* APP_VERSION = L"v0.7"; //MULTI-THREAD

std::atomic<bool> g_LedsEnabled{ true };
std::atomic<bool> g_Running{ true };


// --- GLOBAL CONTROLS FOR STEP 2 ---
CRITICAL_SECTION g_csConfig; // Synchronizes config access between UI and Thread
HANDLE g_hThread = NULL;     // Handle for the temperature worker thread
HANDLE g_hStopEvent = NULL; // Interruptor de emergencia

// Forward declaration so WinMain knows this function exists later
DWORD WINAPI TemperatureThread(LPVOID lpParam);

// CONFIG STRUCTURE
struct Config {
    wchar_t sensorID[256];
    int tempLow, tempHigh, tempAlert;
    bool lightningEffect;
};

Config g_cfg;
const wchar_t* INI_FILE = L".\\config.ini";

typedef int (*LPMLAPI_Initialize)();
typedef int (*LPMLAPI_GetDeviceInfo)(SAFEARRAY** pDevType, SAFEARRAY** pLedCount);
typedef int (*LPMLAPI_SetLedColor)(BSTR type, DWORD index, DWORD R, DWORD G, DWORD B);
typedef int (*LPMLAPI_SetLedStyle)(BSTR type, DWORD index, BSTR style);

IWbemServices* g_pSvc = NULL;
IWbemLocator* g_pLoc = NULL;
BSTR g_deviceName = NULL;
BSTR g_styleSteady = NULL;
BSTR g_styleLightning = NULL;
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
    g_cfg.lightningEffect = GetPrivateProfileIntW(L"Settings", L"Lightning", 1, INI_FILE) != 0;
}

void SaveSettings() {
    WritePrivateProfileStringW(L"Settings", L"SensorID", g_cfg.sensorID, INI_FILE);
    WritePrivateProfileStringW(L"Settings", L"TempLow", std::to_wstring(g_cfg.tempLow).c_str(), INI_FILE);
    WritePrivateProfileStringW(L"Settings", L"TempHigh", std::to_wstring(g_cfg.tempHigh).c_str(), INI_FILE);
    WritePrivateProfileStringW(L"Settings", L"TempAlert", std::to_wstring(g_cfg.tempAlert).c_str(), INI_FILE);
    WritePrivateProfileStringW(L"Settings", L"Lightning", g_cfg.lightningEffect ? L"1" : L"0", INI_FILE);
}

void PopulateSensorList(HWND hDlg) {
    HWND hCombo = GetDlgItem(hDlg, IDC_SENSOR_ID);

    // --- 1. PREVENT MEMORY LEAKS ---
    // Before resetting the ComboBox, we must free the strings allocated with _wcsdup 
    // stored in the ItemData of each entry.
    int count = (int)SendMessage(hCombo, CB_GETCOUNT, 0, 0);
    for (int i = 0; i < count; i++) {
        void* ptr = (void*)SendMessage(hCombo, CB_GETITEMDATA, i, 0);
        if (ptr && ptr != (void*)CB_ERR) {
            free(ptr); // Release the heap-allocated wide string
        }
    }
    // Now it's safe to clear the list
    SendMessage(hCombo, CB_RESETCONTENT, 0, 0);

    // --- 2. WMI CONNECTION CHECK ---
    if (!g_pSvc) {
        InitWMI();
    }

    bool found = false;
    if (g_pSvc) {
        IEnumWbemClassObject* pEnumerator = NULL;
        // Query LibreHardwareMonitor for temperature sensors only
        bstr_t query("SELECT Name, Identifier FROM Sensor WHERE SensorType = 'Temperature'");

        HRESULT hr = g_pSvc->ExecQuery(
            bstr_t("WQL"),
            query,
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            NULL,
            &pEnumerator
        );

        if (SUCCEEDED(hr)) {
            IWbemClassObject* pclsObj = NULL;
            ULONG uReturn = 0;

            // Iterate through the results
            while (pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn) == S_OK) {
                VARIANT vtName, vtID;
                pclsObj->Get(L"Name", 0, &vtName, 0, 0);
                pclsObj->Get(L"Identifier", 0, &vtID, 0, 0);

                // Add the user-friendly name (e.g., "CPU Core #1") to the dropdown
                int idx = (int)SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)vtName.bstrVal);

                // --- 3. ALLOCATE PERSISTENT DATA ---
                // We duplicate the unique Identifier string because the WMI object 
                // will be released shortly. This pointer is freed in WM_DESTROY or next Populate call.
                wchar_t* persistentID = _wcsdup(vtID.bstrVal);
                SendMessage(hCombo, CB_SETITEMDATA, idx, (LPARAM)persistentID);

                // Check if this sensor matches the one currently saved in config.ini
                if (wcscmp(vtID.bstrVal, g_cfg.sensorID) == 0) {
                    SendMessage(hCombo, CB_SETCURSEL, idx, 0);
                    found = true;
                }

                VariantClear(&vtName);
                VariantClear(&vtID);
                pclsObj->Release();
            }
            pEnumerator->Release();
        }
    }

    // --- 4. UI STATE HANDLING ---
    if (SendMessage(hCombo, CB_GETCOUNT, 0, 0) == 0) {
        // No sensors found: LHM might be closed or WMI server disabled
        SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)L"!!! ERROR: OPEN LHM (and reopen settings) !!!");
        SendMessage(hCombo, CB_SETCURSEL, 0, 0);
        EnableWindow(GetDlgItem(hDlg, IDOK), FALSE); // Block saving if no sensor is selected
    }
    else {
        // Sensors available: If the previously saved ID wasn't found, select the first one
        if (!found) {
            SendMessage(hCombo, CB_SETCURSEL, 0, 0);
        }
        EnableWindow(GetDlgItem(hDlg, IDOK), TRUE);
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
        CheckDlgButton(hDlg, IDC_CHECK_ALERTA, g_cfg.lightningEffect ? BST_CHECKED : BST_UNCHECKED);
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

            EnterCriticalSection(&g_csConfig); // Bloqueamos para que el hilo no lea mientras escribimos
            if (selectedID) wcscpy_s(g_cfg.sensorID, selectedID);
            g_cfg.tempLow = tLow;
            g_cfg.tempHigh = tHigh;
            g_cfg.tempAlert = tAlert;
            g_cfg.lightningEffect = (IsDlgButtonChecked(hDlg, IDC_CHECK_ALERTA) == BST_CHECKED);
            LeaveCriticalSection(&g_csConfig); // Desbloqueamos

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
            if (ptr) free(ptr);
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
    NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA) };
    nid.hWnd = hWnd; nid.uID = 1;
    Shell_NotifyIcon(NIM_DELETE, &nid);
    if (hWnd) UnregisterHotKey(hWnd, 1);
    if (g_deviceName) SysFreeString(g_deviceName);
    if (g_styleSteady) SysFreeString(g_styleSteady);
    if (g_styleLightning) SysFreeString(g_styleLightning);
    if (g_pSvc) g_pSvc->Release();
    if (g_pLoc) g_pLoc->Release();
    CoUninitialize();
    if (g_hLibrary) FreeLibrary(g_hLibrary);
    if (g_hMutex) { ReleaseMutex(g_hMutex); CloseHandle(g_hMutex); }
    if (g_hStopEvent) {
        CloseHandle(g_hStopEvent);
        g_hStopEvent = NULL;
    }
    Log("[MysticFight] Cleanup finished");
}

static LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    
    case WM_HOTKEY:
        if (wParam == 1) {
            
            bool newState = !g_LedsEnabled.load();
            g_LedsEnabled.store(newState);

            if (newState) {
                MessageBeep(MB_OK);
                SetEvent(g_hStopEvent);
            }
            else {
                MessageBeep(MB_ICONHAND);
                // Apagado físico inmediato
                for (int i = 0; i < g_totalLeds; i++) lpMLAPI_SetLedColor(g_deviceName, i, 0, 0, 0);
            }
        }
        return 0; 

    case WM_TRAYICON:
        if (lParam == WM_RBUTTONUP) {
            POINT curPoint; GetCursorPos(&curPoint);
            HMENU hMenu = CreatePopupMenu();
            if (hMenu) {
                InsertMenu(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_CONFIG, L"Settings");
                InsertMenu(hMenu, -1, MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
                InsertMenu(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_EXIT, L"Exit MysticFight");

                SetForegroundWindow(hWnd); // Necesario para que el menú se cierre al pinchar fuera
                TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, curPoint.x, curPoint.y, 0, hWnd, NULL);
                DestroyMenu(hMenu);
            }
        }
        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == ID_TRAY_EXIT) {
            Log("[MysticFight] Exit requested from Menu.");
            g_Running = false; // Esto detiene el bucle del hilo de temperatura
            PostQuitMessage(0); // Esto rompe el bucle GetMessage en WinMain
        }
        if (LOWORD(wParam) == ID_TRAY_CONFIG) {
            DialogBoxW(GetModuleHandle(NULL), MAKEINTRESOURCE(IDD_SETTINGS), hWnd, SettingsDlgProc);
        }
        break;

    case WM_CLOSE:
    case WM_DESTROY:
        g_Running = false;
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

static float GetCPUTempFast(const wchar_t* sensorID) { // <--- Recibe el ID local
    if (!g_pSvc || !sensorID) return 0.0f;
    static bstr_t bstrWQL("WQL");

    // Usamos el sensorID pasado, no g_cfg.sensorID directamente
    std::wstring queryStr = L"SELECT Value FROM Sensor WHERE Identifier = '" + std::wstring(sensorID) + L"'";

    bstr_t bstrQuery(queryStr.c_str());
    float temp = 0.0f;
    IEnumWbemClassObject* pEnumerator = NULL;

    if (SUCCEEDED(g_pSvc->ExecQuery(bstrWQL, bstrQuery, WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator))) {
        IWbemClassObject* pclsObj = NULL;
        ULONG uReturn = 0;
        if (pEnumerator->Next(100, 1, &pclsObj, &uReturn) == S_OK && uReturn > 0) {
            VARIANT vtProp;
            pclsObj->Get(L"Value", 0, &vtProp, 0, 0);
            temp = (vtProp.vt == VT_R4) ? vtProp.fltVal : 0.0f; // Verificación extra de tipo
            VariantClear(&vtProp);
            pclsObj->Release();
        }
        pEnumerator->Release();
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

DWORD WINAPI TemperatureThread(LPVOID lpParam) {
    CoInitializeEx(NULL, COINIT_MULTITHREADED);

    SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);

    DWORD R = 0, G = 0, B = 0;
    DWORD lastR = 999, lastG = 999;
    bool modoAlertaActivo = false;
    Config localCfg;

    while (g_Running.load()) {
        if (g_LedsEnabled.load()) {
            EnterCriticalSection(&g_csConfig);
            localCfg = g_cfg;
            LeaveCriticalSection(&g_csConfig);

            float rawTemp = GetCPUTempFast(localCfg.sensorID);
            if (rawTemp <= 0) {
                InitWMI();
            }
            else {
                float temp = floorf(rawTemp * 4.0f + 0.5f) / 4.0f;

                if (temp >= (float)localCfg.tempAlert) {
                    R = 255; G = 0; B = 0;
                    if (!modoAlertaActivo && localCfg.lightningEffect) {
                        lpMLAPI_SetLedStyle(g_deviceName, 0, g_styleLightning);
                        modoAlertaActivo = true;
                    }
                }
                else {
                    if (modoAlertaActivo) {
                        lpMLAPI_SetLedStyle(g_deviceName, 0, g_styleSteady);
                        modoAlertaActivo = false;
                    }
                    // ... (Tu lógica de colores ratio Low/High/Alert se mantiene igual)
                    if (temp <= (float)localCfg.tempLow) { R = 0; G = 255; B = 0; }
                    else if (temp <= (float)localCfg.tempHigh) {
                        float ratio = (temp - (float)localCfg.tempLow) / ((float)localCfg.tempHigh - (float)localCfg.tempLow);
                        R = (DWORD)(255 * ratio); G = 255; B = 0;
                    }
                    else {
                        float ratio = (temp - (float)localCfg.tempHigh) / ((float)localCfg.tempAlert - (float)localCfg.tempHigh);
                        R = 255; G = (DWORD)(255 * (1.0f - ratio)); B = 0;
                    }
                }

                // Si cambia el color O si acabamos de activar g_LedsEnabled (lastR será 999)
                if (R != lastR || G != lastG) {
                    for (int i = 0; i < g_totalLeds; i++) {
                        lpMLAPI_SetLedColor(g_deviceName, i, R, G, B);
                    }
                    lastR = R; lastG = G;
                }
            }
        }
        else {
            // Si está desactivado, reseteamos las memorias para que al activar pinte a la primera
            lastR = 999;
            lastG = 999;
        }

        if (WaitForSingleObject(g_hStopEvent, 500) == WAIT_OBJECT_0) {

            ResetEvent(g_hStopEvent);

            if (!g_Running.load()) {
                break;
            }
        }
    }

    Log("[MysticFight] Temperature monitoring thread BYE BYE");

    CoUninitialize();

    return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    SetPriorityClass(GetCurrentProcess(), HIGH_PRIORITY_CLASS);

    CoInitializeEx(0, COINIT_MULTITHREADED);

    InitializeCriticalSection(&g_csConfig);

    LoadSettings();

    g_hMutex = CreateMutex(NULL, TRUE, L"Global\\MysticFight_Unique_Mutex");
    if (g_hMutex == NULL || GetLastError() == ERROR_ALREADY_EXISTS) {
        if (g_hMutex) {
            MessageBox(NULL, L"Error: An instance of MysticFight is already running.", L"Error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
            CloseHandle(g_hMutex);
        }
        CoUninitialize();
        DeleteCriticalSection(&g_csConfig); // Cleanup CS
        return 0;
    }

    if (wcslen(g_cfg.sensorID) == 0) {
        if (DialogBoxW(hInstance, MAKEINTRESOURCE(IDD_SETTINGS), NULL, SettingsDlgProc) == IDCANCEL) {
            if (g_hMutex) CloseHandle(g_hMutex);
            CoUninitialize();
            DeleteCriticalSection(&g_csConfig);
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
        DeleteCriticalSection(&g_csConfig);
        return 1;
    }

    auto lpMLAPI_Initialize = (LPMLAPI_Initialize)GetProcAddress(g_hLibrary, "MLAPI_Initialize");
    auto lpMLAPI_GetDeviceInfo = (LPMLAPI_GetDeviceInfo)GetProcAddress(g_hLibrary, "MLAPI_GetDeviceInfo");
    lpMLAPI_SetLedColor = (LPMLAPI_SetLedColor)GetProcAddress(g_hLibrary, "MLAPI_SetLedColor");
    lpMLAPI_SetLedStyle = (LPMLAPI_SetLedStyle)GetProcAddress(g_hLibrary, "MLAPI_SetLedStyle");

    Log("[MysticLight] Attempting to initialize SDK...");
    if (lpMLAPI_Initialize() != 0) {
        bool initialized = false;
        for (int i = 1; i <= 10; i++) {
            char retryMsg[100];
            snprintf(retryMsg, sizeof(retryMsg), "[MysticLight] Attempt %d/10 failed. Retrying in 5s...", i);
            Log(retryMsg);
            if (lpMLAPI_Initialize() == 0) {
                Log("[MysticLight] SDK Initialized successfully on retry.");
                initialized = true;
                break;
            }
            if (i == 10) {
                Log("[MysticLight] FATAL: SDK could not be initialized after 10 attempts.");
                MessageBox(NULL, L"MSI SDK Initialization failed. Ensure MSI Center/Dragon Center is installed and service is running.", L"Critical Error", MB_OK | MB_ICONERROR);
                FinalCleanup(hWnd);
                DeleteCriticalSection(&g_csConfig);
                return 1;
            }
            Sleep(5000);
        }
    }
    else {
        Log("[MysticLight] SDK Initialized successfully at first attempt.");
    }

    SAFEARRAY* pDevType = nullptr, * pLedCount = nullptr;
    if (lpMLAPI_GetDeviceInfo(&pDevType, &pLedCount) == 0 && pDevType) {
        BSTR* pData = (BSTR*)(pDevType->pvData);
        BSTR* pCounts = (BSTR*)(pLedCount->pvData);
        g_deviceName = SysAllocString(pData[0]);
        g_totalLeds = _wtoi(pCounts[0]);
        SafeArrayDestroy(pDevType); SafeArrayDestroy(pLedCount);

        char devInfo[150];
        snprintf(devInfo, sizeof(devInfo), "[MysticLight] Device detected: %ls | LEDs: %d", g_deviceName, g_totalLeds);
        Log(devInfo);
    }

    g_styleSteady = SysAllocString(L"Steady");
    g_styleLightning = SysAllocString(L"Lightning");

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
            DeleteCriticalSection(&g_csConfig);
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

    // --- STEP 2: LAUNCH WORKER THREAD ---
    g_hStopEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
    g_hThread = CreateThread(NULL, 0, TemperatureThread, NULL, 0, NULL);

    // --- MAIN UI LOOP ---
    MSG msg = { 0 };
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // --- SHUTDOWN ---
    g_Running = false;
    
    if (g_hStopEvent) SetEvent(g_hStopEvent);
    
    if (g_hThread) {
        DWORD waitResult = WaitForSingleObject(g_hThread, 2000);

        if (waitResult == WAIT_TIMEOUT) {
            Log("[Thread] Temp thread is stuck! Forcefully terminating...");
            TerminateThread(g_hThread, 0);
        }

        CloseHandle(g_hThread);

        g_hThread = NULL;
    }

    DeleteCriticalSection(&g_csConfig);

    FinalCleanup(hWnd);

    Log("[MysticFight] Main thread BYE BYE");

    return 0;
}
