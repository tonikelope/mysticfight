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
#include <winhttp.h>
#include <algorithm>
#include <thread>
#include <atomic>
#include "resource.h"

#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "winmm.lib")

// =============================================================
// GLOBAL CONSTANTS AND DEFINITIONS
// =============================================================

// UI & Tray Identifiers
#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 1001
#define ID_TRAY_CONFIG 2001
#define ID_TRAY_LOG 3001
#define ID_TRAY_ABOUT 4001

// Application Metadata
const wchar_t* APP_VERSION = L"v2.45";
const wchar_t* LOG_FILENAME = L"debug.log";
const wchar_t* INI_FILE = L".\\config.ini";
const wchar_t* TASK_NAME = L"MysticFight";

// Sentinel Values for State Machine
const DWORD RGB_LED_REFRESH = 999;  // Signals a mandatory style refresh (e.g., startup or reset)
const DWORD RGB_LEDS_OFF = 1000;     // Signals that the hardware is currently in the "Off" state

// Timing Configuration (Milliseconds)
const ULONGLONG MAIN_LOOP_DELAY_MS = 40;        // Animation speed (25 FPS).
const ULONGLONG SENSOR_POLL_INTERVAL_MS = 500;  // How often we check the Sensor (Temperature)
const float SMOOTHING_FACTOR = 0.15f;           // Interpolation speed (0.01 = Slow, 1.0 = Instant).

const ULONGLONG LHM_RETRY_DELAY_MS = 5000;      // Cooldown before retrying WMI connection
const ULONGLONG RESET_KILL_TASK_WAIT_MS = 2000; // Watchdog: Time to wait after killing processes
const ULONGLONG RESET_RESTART_TASK_DELAY_MS = 5000; // Watchdog: Time to wait after restarting service
const ULONGLONG WINHTTP_TIMEOUT_MS = 60000;

// Buffer Sizes
const int HEX_COLOR_LEN = 7;
const int LOG_BUFFER_SIZE = 512;
const int SENSOR_ID_LEN = 256;

// =============================================================
// DATA STRUCTURES
// =============================================================

struct Config {
    wchar_t sensorID[SENSOR_ID_LEN];
    wchar_t targetDevice[256];      // Internal MSI device type (e.g., L"MSI_MB")
    int targetLedIndex;             // Selected LED area index
    int tempLow, tempMed, tempHigh;
    COLORREF colorLow, colorMed, colorHigh;
    wchar_t webServerUrl[256];
};

Config g_cfg;

// =============================================================
// DATA SOURCE STATE MACHINE & CONSTANTS
// =============================================================

// LHM HTTP Config
const wchar_t* LHM_PATH = L"/data.json";

// State Machine for Data Source
enum class DataSource {
    Searching, // Initial state or after failure: try WMI, then HTTP
    WMI,       // Locked state: Use only WMI
    HTTP       // Locked state: Use only HTTP
};

// Global state variable
DataSource g_activeSource = DataSource::Searching;

// Appends log entries to the debug file.
static void Log(const char* text) {
    time_t now = time(0);
    char dt[26];
    ctime_s(dt, sizeof(dt), &now);
    dt[24] = '\0';

    std::ofstream out(LOG_FILENAME, std::ios::app);
    if (out) {
        out << "[" << dt << "] " << text << "\n";
    }
}

// =============================================================
// HTTP & JSON PARSING UTILS (FALLBACK SYSTEM)
// =============================================================

// Performs a GET request to LHM and returns raw JSON
static std::string FetchLHMJson() {
    std::string responseData;
    HINTERNET hSession = NULL, hConnect = NULL, hRequest = NULL;
    URL_COMPONENTS urlComp = { sizeof(URL_COMPONENTS) };

    urlComp.dwHostNameLength = (DWORD)-1;
    urlComp.dwUrlPathLength = (DWORD)-1;
    urlComp.dwExtraInfoLength = (DWORD)-1;

    if (!WinHttpCrackUrl(g_cfg.webServerUrl, (DWORD)wcslen(g_cfg.webServerUrl), 0, &urlComp)) {
        Log("[MysticFight] Error: Invalid Web Server URL format.");
        return "";
    }

    hSession = WinHttpOpen(L"MysticFight/2.3", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) return "";

    WinHttpSetTimeouts(hSession, WINHTTP_TIMEOUT_MS, WINHTTP_TIMEOUT_MS, WINHTTP_TIMEOUT_MS, WINHTTP_TIMEOUT_MS);

    std::wstring host(urlComp.lpszHostName, urlComp.dwHostNameLength);
    hConnect = WinHttpConnect(hSession, host.c_str(), urlComp.nPort, 0);

    if (hConnect) {
        std::wstring path = std::wstring(urlComp.lpszUrlPath, urlComp.dwUrlPathLength) + L"/data.json";
        hRequest = WinHttpOpenRequest(hConnect, L"GET", path.c_str(), NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, 0);

        if (hRequest && WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
            WinHttpReceiveResponse(hRequest, NULL)) {

            DWORD dwSize = 0, dwDownloaded = 0;
            // Bucle optimizado con vector para seguridad de memoria
            while (WinHttpQueryDataAvailable(hRequest, &dwSize) && dwSize > 0) {
                std::vector<char> buffer(dwSize);
                if (WinHttpReadData(hRequest, (LPVOID)buffer.data(), dwSize, &dwDownloaded)) {
                    responseData.append(buffer.data(), dwDownloaded);
                }
            }
        }
    }

    if (hRequest) WinHttpCloseHandle(hRequest);
    if (hConnect) WinHttpCloseHandle(hConnect);
    if (hSession) WinHttpCloseHandle(hSession);
    return responseData;
}

// "Surgical" parser to extract value from specific SensorID
static float ParseLHMJsonForTemp(const std::string& json, const wchar_t* sensorIDW) {
    if (json.empty() || !sensorIDW) return -1.0f;

    char sensorID[256];
    size_t converted;
    wcstombs_s(&converted, sensorID, sizeof(sensorID), sensorIDW, _TRUNCATE);

    // 1. Find Sensor ID key
    std::string searchKey = "\"SensorId\":\"" + std::string(sensorID) + "\"";
    size_t idPos = json.find(searchKey);
    if (idPos == std::string::npos) return -1.0f;

    // 2. Isolate the JSON object {...} surrounding this ID
    size_t objStart = json.rfind('{', idPos);
    size_t objEnd = json.find('}', idPos);
    if (objStart == std::string::npos || objEnd == std::string::npos) return -1.0f;

    std::string objBlock = json.substr(objStart, objEnd - objStart);

    // 3. Find "Value" within this block
    size_t valPos = objBlock.find("\"Value\":");
    if (valPos == std::string::npos) return -1.0f;

    valPos += 9; // Skip "Value":"
    size_t valEnd = objBlock.find('"', valPos);
    if (valEnd == std::string::npos) return -1.0f;

    std::string valStr = objBlock.substr(valPos, valEnd - valPos); // e.g., "49,4 °C"

    // 4. Sanitize (European comma to dot)
    std::replace(valStr.begin(), valStr.end(), ',', '.');

    try {
        return std::stof(valStr);
    }
    catch (...) {
        return -1.0f;
    }
}

// =============================================================
// MSI SDK TYPE DEFINITIONS
// =============================================================
typedef int (*LPMLAPI_Initialize)();
typedef int (*LPMLAPI_GetDeviceInfo)(SAFEARRAY** pDevType, SAFEARRAY** pLedCount);
typedef int (*LPMLAPI_GetDeviceNameEx)(BSTR type, DWORD index, BSTR* pDevName);
typedef int (*LPMLAPI_GetLedInfo)(BSTR, DWORD, BSTR*, SAFEARRAY**);
typedef int (*LPMLAPI_SetLedColor)(BSTR type, DWORD index, DWORD R, DWORD G, DWORD B);
typedef int (*LPMLAPI_SetLedStyle)(BSTR type, DWORD index, BSTR style);
typedef int (*LPMLAPI_SetLedSpeed)(BSTR type, DWORD index, DWORD level);
typedef int (*LPMLAPI_Release)();

// =============================================================
// COM SMART POINTERS
// =============================================================
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

// =============================================================
// GLOBAL STATE VARIABLES
// =============================================================
_bstr_t g_cachedSensorPath = L"";
_bstr_t g_target_device = L"";
bool g_pathCached = false;
bool g_Running = true;
bool g_LedsEnabled = true;

std::atomic<float> g_asyncTemp(-1.0f);
HANDLE g_hSensorEvent = NULL;

// Watchdog & Timing State
ULONGLONG g_ResetTimer = 0;
bool g_Resetting_sdk = false;
int g_ResetStage = 0;
ULONGLONG g_lastDataSourceSearchRetry = 0;

// WMI Interfaces
IWbemServicesPtr g_pSvc = nullptr;
IWbemLocatorPtr g_pLoc = nullptr;

// LED State Tracking
DWORD lastR = RGB_LED_REFRESH;
DWORD lastG = RGB_LED_REFRESH;
DWORD lastB = RGB_LED_REFRESH;

// Hardware Handles
BSTR g_deviceName = NULL;
HMODULE g_hLibrary = NULL;
int g_totalLeds = 0;
HANDLE g_hMutex = NULL;

// SDK Function Pointers
LPMLAPI_Initialize lpMLAPI_Initialize = nullptr;
LPMLAPI_GetDeviceInfo lpMLAPI_GetDeviceInfo = nullptr;
LPMLAPI_SetLedColor lpMLAPI_SetLedColor = nullptr;
LPMLAPI_SetLedStyle lpMLAPI_SetLedStyle = nullptr;
LPMLAPI_SetLedSpeed lpMLAPI_SetLedSpeed = nullptr;
LPMLAPI_Release lpMLAPI_Release = nullptr;
LPMLAPI_GetDeviceNameEx lpMLAPI_GetDeviceNameEx = nullptr;
LPMLAPI_GetLedInfo lpMLAPI_GetLedInfo = nullptr;


// =============================================================
// UTILITY FUNCTIONS
// =============================================================

// Executes a process using Explorer to avoid inheriting Admin privileges.
static void RunShellNonAdmin(const wchar_t* path) {
    ShellExecuteW(NULL, L"open", L"explorer.exe", path, NULL, SW_SHOWNORMAL);
}

// Validates Hex color string format (#RRGGBB).
static bool IsValidHex(const wchar_t* hex) {
    if (!hex || wcslen(hex) != HEX_COLOR_LEN || hex[0] != L'#') return false;
    for (int i = 1; i < HEX_COLOR_LEN; i++) {
        if (!iswxdigit(hex[i])) return false;
    }
    return true;
}

// Converts Hex string to COLORREF.
static COLORREF HexToColor(const wchar_t* hex) {
    if (hex[0] == L'#') hex++;
    unsigned int r = 0, g = 0, b = 0;
    if (swscanf_s(hex, L"%02x%02x%02x", &r, &g, &b) == 3) {
        return RGB(r, g, b);
    }
    return RGB(0, 0, 0);
}

// Converts COLORREF to Hex string.
static void ColorToHex(COLORREF color, wchar_t* out, size_t size) {
    swprintf_s(out, size, L"#%02X%02X%02X", GetRValue(color), GetGValue(color), GetBValue(color));
}

// Allocates string on the process heap (Memory Safety).
static wchar_t* HeapDupString(const wchar_t* src) {
    if (!src) return nullptr;
    size_t size = (wcslen(src) + 1) * sizeof(wchar_t);
    wchar_t* dest = (wchar_t*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, size);
    if (dest) wcscpy_s(dest, wcslen(src) + 1, src);
    return dest;
}

// Cleans up heap memory associated with ComboBox items.
static void ClearComboHeapData(HWND hCombo) {
    int count = (int)SendMessage(hCombo, CB_GETCOUNT, 0, 0);
    for (int i = 0; i < count; i++) {
        wchar_t* ptr = (wchar_t*)SendMessage(hCombo, CB_GETITEMDATA, i, 0);
        if (ptr && ptr != (wchar_t*)CB_ERR) HeapFree(GetProcessHeap(), 0, ptr);
    }
    SendMessage(hCombo, CB_RESETCONTENT, 0, 0);
}

// Initializes WMI connection to LibreHardwareMonitor.
static bool InitWMI() {
    g_pSvc = nullptr;
    g_pLoc = nullptr;
    g_pathCached = false;

    HRESULT hr = g_pLoc.CreateInstance(CLSID_WbemLocator, nullptr, CLSCTX_INPROC_SERVER);
    if (FAILED(hr)) return false;

    hr = g_pLoc->ConnectServer(
        _bstr_t(L"ROOT\\LibreHardwareMonitor"),
        nullptr, nullptr, nullptr, 0, nullptr, nullptr,
        &g_pSvc
    );
    if (FAILED(hr)) return false;

    hr = CoSetProxyBlanket(
        g_pSvc, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, nullptr,
        RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE
    );

    return SUCCEEDED(hr);
}



// Truncates log file if it exceeds size limit to prevent bloat.
static void TrimLogFile() {
    WIN32_FILE_ATTRIBUTE_DATA fad;
    if (!GetFileAttributesExW(LOG_FILENAME, GetFileExInfoStandard, &fad)) return;

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
            for (size_t i = lines.size() - 500; i < lines.size(); ++i) {
                out << lines[i] << "\n";
            }
        }
    }
}

// Controls Windows Task Scheduler tasks (Start/Stop).
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
            IRunningTaskPtr pRunningTask;
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

// Verifies if the scheduled task exists and points to the current executable.
static bool ValidStartupTaskExists() {
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
        if (FAILED(pRootFolder->GetTask(_bstr_t(TASK_NAME), &pRegisteredTask))) return false;

        ITaskDefinitionPtr pDefinition;
        if (FAILED(pRegisteredTask->get_Definition(&pDefinition))) return false;

        IActionCollectionPtr pActions;
        if (FAILED(pDefinition->get_Actions(&pActions))) return false;

        IActionPtr pAction;
        if (FAILED(pActions->get_Item(1, &pAction))) return false;

        IExecActionPtr pExecAction = pAction;
        if (pExecAction == nullptr) return false;

        BSTR bstrPath = NULL;
        if (SUCCEEDED(pExecAction->get_Path(&bstrPath))) {
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

// Registers application in Task Scheduler for high-privilege startup.
static void SetStartupTask(bool run) {
    
    wchar_t szPath[MAX_PATH], szDir[MAX_PATH];
    
    if (GetModuleFileNameW(NULL, szPath, MAX_PATH) == 0) return;
    
    if (run && ValidStartupTaskExists()) return;
    
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

        hr = pRootFolder->DeleteTask(_bstr_t(TASK_NAME), 0);

        if(SUCCEEDED(hr)) Log("[MysticFight] Startup task removed");
        
        if (!run) return;

        ITaskDefinitionPtr pTask;
        pService->NewTask(0, &pTask);

        IPrincipalPtr pPrincipal;
        pTask->get_Principal(&pPrincipal);
        if (pPrincipal != nullptr) {
            pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);
            pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
        }

        ITaskSettingsPtr pSettings;
        pTask->get_Settings(&pSettings);
        if (pSettings != nullptr) {
            pSettings->put_StartWhenAvailable(VARIANT_TRUE);
            pSettings->put_DisallowStartIfOnBatteries(VARIANT_FALSE);
            pSettings->put_StopIfGoingOnBatteries(VARIANT_FALSE);
            pSettings->put_ExecutionTimeLimit(_bstr_t(L"PT0S"));
            pSettings->put_AllowHardTerminate(VARIANT_TRUE);
        }

        ITriggerCollectionPtr pTriggers;
        pTask->get_Triggers(&pTriggers);
        ITriggerPtr pTrigger;
        pTriggers->Create(TASK_TRIGGER_LOGON, &pTrigger);

        ILogonTriggerPtr pLogonTrigger = pTrigger;
        if (pLogonTrigger != nullptr) {
            pLogonTrigger->put_Delay(_bstr_t(L"PT30S"));
        }

        IActionCollectionPtr pActions;
        pTask->get_Actions(&pActions);
        IActionPtr pAction;
        pActions->Create(TASK_ACTION_EXEC, &pAction);

        IExecActionPtr pExecAction = pAction;
        if (pExecAction != nullptr) {
            pExecAction->put_Path(_bstr_t(szPath));
            pExecAction->put_WorkingDirectory(_bstr_t(szDir));
        }

        IRegisteredTaskPtr pRegisteredTask;
        pRootFolder->RegisterTaskDefinition(
            _bstr_t(TASK_NAME),
            pTask,
            TASK_CREATE_OR_UPDATE,
            _variant_t(), _variant_t(),
            TASK_LOGON_INTERACTIVE_TOKEN,
            _variant_t(L""),
            &pRegisteredTask
        );

        Log("[MysticFight] New Startup task created");
    }
    catch (const _com_error& e) {
        char buffer[256];
        snprintf(buffer, sizeof(buffer), "[MysticFight] TaskScheduler ERROR: HRESULT 0x%08X", e.Error());
        Log(buffer);
    }
}

// Safely extracts integers from SafeArray variants.
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
        return 0; // Unsupported Type
    }
}



// Writes configuration settings to INI file.
void SaveSettings() {
    WritePrivateProfileStringW(L"Settings", L"SensorID", g_cfg.sensorID, INI_FILE);
    WritePrivateProfileStringW(L"Settings", L"TargetDevice", g_cfg.targetDevice, INI_FILE);
    WritePrivateProfileStringW(L"Settings", L"TargetLedIndex", std::to_wstring(g_cfg.targetLedIndex).c_str(), INI_FILE);
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

    WritePrivateProfileStringW(L"Settings", L"WebServerUrl", g_cfg.webServerUrl, INI_FILE);
}

// Reads configuration from INI file with full validation.
void LoadSettings() {
    bool needsReset = false;

    // 1. Load temperature thresholds
    int tL = GetPrivateProfileIntW(L"Settings", L"TempLow", 50, INI_FILE);
    int tM = GetPrivateProfileIntW(L"Settings", L"TempMed", 70, INI_FILE);
    int tH = GetPrivateProfileIntW(L"Settings", L"TempHigh", 90, INI_FILE);

    // 2. Load hardware target settings
    // Default to "MSI_MB" if missing
    GetPrivateProfileStringW(L"Settings", L"TargetDevice", L"MSI_MB", g_cfg.targetDevice, 256, INI_FILE);
    // Default to 0 (Select All) if missing
    g_cfg.targetLedIndex = GetPrivateProfileIntW(L"Settings", L"TargetLedIndex", 0, INI_FILE);

    // 3. Load Colors
    wchar_t hL[10], hM[10], hH[10];
    GetPrivateProfileStringW(L"Settings", L"ColorLow", L"#00FF00", hL, 10, INI_FILE);
    GetPrivateProfileStringW(L"Settings", L"ColorMed", L"#FFFF00", hM, 10, INI_FILE);
    GetPrivateProfileStringW(L"Settings", L"ColorHigh", L"#FF0000", hH, 10, INI_FILE);

    GetPrivateProfileStringW(L"Settings", L"WebServerUrl", L"http://localhost:8085", g_cfg.webServerUrl, 256, INI_FILE);

    // 4. VALIDATION LOGIC
    // We check for:
    // - Invalid Temperatures (out of order or out of bounds)
    // - Invalid Hex Colors
    // - NEW: Invalid Device Name (Empty string)
    // - NEW: Invalid LED Index (Negative or impossibly high)
    if (tL < 0 || tH > 110 || tL >= tM || tM >= tH ||
        !IsValidHex(hL) || !IsValidHex(hM) || !IsValidHex(hH) ||
        wcslen(g_cfg.targetDevice) == 0 ||  // Check if Device Name is empty
        g_cfg.targetLedIndex < 0 ||         // Check if Index is negative
        g_cfg.targetLedIndex > 255)         // Check if Index is too high (safety cap)
    {
        needsReset = true;
    }

    if (needsReset) {
        // 5. RESTORE DEFAULTS
        g_cfg.tempLow = 50; g_cfg.tempMed = 70; g_cfg.tempHigh = 90;
        g_cfg.colorLow = RGB(0, 255, 0); g_cfg.colorMed = RGB(255, 255, 0); g_cfg.colorHigh = RGB(255, 0, 0);

        // Reset Hardware to safe defaults
        wcscpy_s(g_cfg.targetDevice, L"MSI_MB");
        g_cfg.targetLedIndex = 0;

        SaveSettings(); // Overwrite the corrupted INI file

        Log("[MysticFight] Config file corrupted or invalid. Factory defaults restored.");
        MessageBoxW(NULL, L"Configuration was corrupted. Factory defaults restored.", L"MysticFight", MB_OK | MB_ICONINFORMATION);
    }
    else {
        // 6. APPLY LOADED VALUES
        g_cfg.tempLow = tL; g_cfg.tempMed = tM; g_cfg.tempHigh = tH;
        g_cfg.colorLow = HexToColor(hL); g_cfg.colorMed = HexToColor(hM); g_cfg.colorHigh = HexToColor(hH);
        // Note: targetDevice and targetLedIndex are already in g_cfg from the read step
    }

    // Always load SensorID last (it's allowed to be empty on first run)
    GetPrivateProfileStringW(L"Settings", L"SensorID", L"", g_cfg.sensorID, SENSOR_ID_LEN, INI_FILE);
}


void PopulateAreaList(HWND hDlg, const wchar_t* deviceType) {
    HWND hComboArea = GetDlgItem(hDlg, IDC_COMBO_AREA);
    SendMessage(hComboArea, CB_RESETCONTENT, 0, 0);

    if (!deviceType || !lpMLAPI_GetLedInfo || !lpMLAPI_GetDeviceInfo) {
        EnableWindow(hComboArea, FALSE);
        return;
    }

    // --- OBTENER EL CONTEO REAL DE LEDS PARA ESTE DISPOSITIVO ---
    SAFEARRAY* pDevType = nullptr;
    SAFEARRAY* pLedCount = nullptr;
    int currentDeviceLedCount = 0;

    if (lpMLAPI_GetDeviceInfo(&pDevType, &pLedCount) == 0 && pDevType && pLedCount) {
        BSTR* pTypes = nullptr;
        void* pCountsRaw = nullptr;
        VARTYPE vtCount;
        SafeArrayGetVartype(pLedCount, &vtCount);

        if (SUCCEEDED(SafeArrayAccessData(pDevType, (void**)&pTypes)) &&
            SUCCEEDED(SafeArrayAccessData(pLedCount, &pCountsRaw))) {

            long lBound, uBound;
            SafeArrayGetLBound(pDevType, 1, &lBound);
            SafeArrayGetUBound(pDevType, 1, &uBound);
            long totalDevices = uBound - lBound + 1;

            // Buscamos el índice del dispositivo actual para sacar su LedCount
            for (long j = 0; j < totalDevices; j++) {
                if (wcscmp(pTypes[j], deviceType) == 0) {
                    currentDeviceLedCount = GetIntFromSafeArray(pCountsRaw, vtCount, j);
                    break;
                }
            }
            SafeArrayUnaccessData(pDevType);
            SafeArrayUnaccessData(pLedCount);
        }
        SafeArrayDestroy(pDevType);
        SafeArrayDestroy(pLedCount);


        // Si no detectamos LEDs, usamos una búsqueda de seguridad o salimos
        if (currentDeviceLedCount <= 0) currentDeviceLedCount = 0;

        // --- LLENAR EL COMBOBOX CON EL CONTEO REAL ---
        for (DWORD i = 0; i < (DWORD)currentDeviceLedCount; i++) {
            BSTR ledName = nullptr;
            SAFEARRAY* pStyles = nullptr;
            if (lpMLAPI_GetLedInfo((BSTR)deviceType, i, &ledName, &pStyles) == 0) {
                int idx = (int)SendMessageW(hComboArea, CB_ADDSTRING, 0, (LPARAM)(ledName ? ledName : L"Unknown Area"));
                SendMessage(hComboArea, CB_SETITEMDATA, idx, (LPARAM)i);
                if (i == (DWORD)g_cfg.targetLedIndex) SendMessage(hComboArea, CB_SETCURSEL, idx, 0);

                if (pStyles) SafeArrayDestroy(pStyles);
                if (ledName) SysFreeString(ledName);
            }
        }

        // --- LÓGICA DE ACTIVACIÓN/DESACTIVACIÓN (UX) ---
        int finalCount = (int)SendMessage(hComboArea, CB_GETCOUNT, 0, 0);

        if (finalCount <= 1) {
            if (finalCount == 1) SendMessage(hComboArea, CB_SETCURSEL, 0, 0);
            EnableWindow(hComboArea, FALSE);
        }
        else {
            if (SendMessage(hComboArea, CB_GETCURSEL, 0, 0) == CB_ERR)
                SendMessage(hComboArea, CB_SETCURSEL, 0, 0);
            EnableWindow(hComboArea, TRUE);
        }
    }
    else {
        EnableWindow(hComboArea, FALSE);
        Log("[MysticFight] [Settings - PopulateAreaList] SDK Error or No MSI Devices found.");
    }
}

// Populates the list of MSI devices detected by the SDK.
void PopulateDeviceList(HWND hDlg) {
    HWND hComboDev = GetDlgItem(hDlg, IDC_COMBO_DEVICE);
    ClearComboHeapData(hComboDev);
    SAFEARRAY* pDevType = nullptr;
    SAFEARRAY* pLedCount = nullptr;

    if (lpMLAPI_GetDeviceInfo && lpMLAPI_GetDeviceInfo(&pDevType, &pLedCount) == 0) {
        BSTR* pTypes = nullptr;
        SafeArrayAccessData(pDevType, (void**)&pTypes);
        long lBound, uBound;
        SafeArrayGetLBound(pDevType, 1, &lBound);
        SafeArrayGetUBound(pDevType, 1, &uBound);
        long count = uBound - lBound + 1;

        for (long i = 0; i < count; i++) {
            BSTR typeInternal = pTypes[i];
            BSTR friendlyName = NULL;
            if (lpMLAPI_GetDeviceNameEx) lpMLAPI_GetDeviceNameEx(typeInternal, i, &friendlyName);
            int idx = (int)SendMessageW(hComboDev, CB_ADDSTRING, 0, (LPARAM)(friendlyName ? friendlyName : typeInternal));
            SendMessage(hComboDev, CB_SETITEMDATA, idx, (LPARAM)HeapDupString(typeInternal));
            if (wcscmp(typeInternal, g_cfg.targetDevice) == 0) SendMessage(hComboDev, CB_SETCURSEL, idx, 0);
            if (friendlyName) SysFreeString(friendlyName);
        }
        SafeArrayUnaccessData(pDevType);
        SafeArrayDestroy(pDevType);
        SafeArrayDestroy(pLedCount);
        
        if (SendMessage(hComboDev, CB_GETCURSEL, 0, 0) == CB_ERR && SendMessage(hComboDev, CB_GETCOUNT, 0, 0) > 0)
            SendMessage(hComboDev, CB_SETCURSEL, 0, 0);

        // --- LÓGICA DE ACTIVACIÓN/DESACTIVACIÓN (UX) ---
        int finalCount = (int)SendMessage(hComboDev, CB_GETCOUNT, 0, 0);
        if (finalCount <= 1) {
            if (finalCount == 1) SendMessage(hComboDev, CB_SETCURSEL, 0, 0);
            EnableWindow(hComboDev, FALSE);
        }
        else {
            if (SendMessage(hComboDev, CB_GETCURSEL, 0, 0) == CB_ERR)
                SendMessage(hComboDev, CB_SETCURSEL, 0, 0);
            EnableWindow(hComboDev, TRUE);
        }
    }
    else {
        EnableWindow(hComboDev, FALSE);
        Log("[MysticFight] [Settings - PopulateDeviceList] SDK Error or No MSI Devices found.");
    }
}

// Helper to extract string value from a JSON block (e.g., "Text":"CPU Core")
static std::wstring ExtractJsonString(const std::string& block, const std::string& key) {
    std::string searchKey = "\"" + key + "\":\"";
    size_t start = block.find(searchKey);
    if (start == std::string::npos) return L"";

    start += searchKey.length();
    size_t end = block.find("\"", start);
    if (end == std::string::npos) return L"";

    std::string val = block.substr(start, end - start);

    // Convert UTF8 to Wide String (Windows)
    int len = MultiByteToWideChar(CP_UTF8, 0, val.c_str(), -1, NULL, 0);
    if (len > 0) {
        std::vector<wchar_t> buf(len);
        MultiByteToWideChar(CP_UTF8, 0, val.c_str(), -1, &buf[0], len);
        return std::wstring(&buf[0]);
    }
    return L"";
}

void PopulateSensorList(HWND hDlg) {
    HWND hCombo = GetDlgItem(hDlg, IDC_SENSOR_ID);
    ClearComboHeapData(hCombo);

    // --- STEP A: TRY WMI (Priority) ---
    if (g_activeSource == DataSource::WMI) {
        // Lazy connection: Only connect if not already connected
        if (!g_pSvc) InitWMI();

        if (g_pSvc) {
            IEnumWbemClassObjectPtr pEnumerator = NULL;
            bstr_t query("SELECT Name, Identifier FROM Sensor WHERE SensorType = 'Temperature'");

            // Execute Query
            HRESULT hr = g_pSvc->ExecQuery(
                bstr_t("WQL"),
                query,
                WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
                NULL,
                &pEnumerator
            );

            if (SUCCEEDED(hr)) {

                IWbemClassObjectPtr pclsObj = NULL;
                ULONG uReturn = 0;
                bool found = false;

                while (pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn) == S_OK) {
                    _variant_t vtName, vtID;
                    if (SUCCEEDED(pclsObj->Get(L"Name", 0, &vtName, 0, 0)) &&
                        SUCCEEDED(pclsObj->Get(L"Identifier", 0, &vtID, 0, 0))) {

                        if (vtName.vt == VT_BSTR && vtID.vt == VT_BSTR) {
                            int idx = (int)SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)vtName.bstrVal);
                            if (idx != CB_ERR) {
                                SendMessage(hCombo, CB_SETITEMDATA, idx, (LPARAM)HeapDupString(vtID.bstrVal));

                                // Auto-select current config
                                if (wcscmp(vtID.bstrVal, g_cfg.sensorID) == 0) {
                                    SendMessage(hCombo, CB_SETCURSEL, idx, 0);
                                    found = true;
                                }
                            }
                        }
                    }
                }
                
                // Select first item if nothing matched
                if (!found && SendMessage(hCombo, CB_GETCOUNT, 0, 0) > 0)
                    SendMessage(hCombo, CB_SETCURSEL, 0, 0);
            }
        }
    }
    else {

        //HTTP
        std::string json = FetchLHMJson();
        if (json.empty()) return;

        std::string typeKey = "\"Type\":\"Temperature\"";
        size_t pos = 0;
        bool foundAny = false;

        // Search for all occurrences of "Type":"Temperature"
        while ((pos = json.find(typeKey, pos)) != std::string::npos) {
            // Find the limits of the current JSON object {...}
            size_t blockEnd = json.find('}', pos);
            size_t blockStart = json.rfind('{', pos);

            if (blockEnd != std::string::npos && blockStart != std::string::npos) {
                std::string block = json.substr(blockStart, blockEnd - blockStart + 1);

                std::wstring name = ExtractJsonString(block, "Text");
                std::wstring id = ExtractJsonString(block, "SensorId");

                if (!name.empty() && !id.empty()) {
                    int idx = (int)SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)name.c_str());
                    if (idx != CB_ERR) {
                        SendMessage(hCombo, CB_SETITEMDATA, idx, (LPARAM)HeapDupString(id.c_str()));
                        // Select if matches current config
                        if (wcscmp(id.c_str(), g_cfg.sensorID) == 0) {
                            SendMessage(hCombo, CB_SETCURSEL, idx, 0);
                            foundAny = true;
                        }
                    }
                }
            }

            pos = blockEnd; // Advance
        }

        if (!foundAny && SendMessage(hCombo, CB_GETCOUNT, 0, 0) > 0)
            SendMessage(hCombo, CB_SETCURSEL, 0, 0);
    }

    // --- LÓGICA DE ACTIVACIÓN/DESACTIVACIÓN (UX) ---
    int finalCount = (int)SendMessage(hCombo, CB_GETCOUNT, 0, 0);
    
    if (finalCount <= 1) {
        if (finalCount == 1) SendMessage(hCombo, CB_SETCURSEL, 0, 0);
        EnableWindow(hCombo, FALSE);
    }
    else {
        if (SendMessage(hCombo, CB_GETCURSEL, 0, 0) == CB_ERR)
            SendMessage(hCombo, CB_SETCURSEL, 0, 0);
        EnableWindow(hCombo, TRUE);
    }
}

static void CacheSensorPath() {
    // Reset state: assume failure until proven otherwise
    g_pathCached = false;

    if (!g_pSvc || wcslen(g_cfg.sensorID) == 0) return;

    std::wstring wqlQuery = L"SELECT * FROM Sensor WHERE Identifier = '" + std::wstring(g_cfg.sensorID) + L"'";

    char debugBuf[LOG_BUFFER_SIZE];
    snprintf(debugBuf, sizeof(debugBuf), "[MysticFight] Searching WMI path for: %ls", g_cfg.sensorID);
    Log(debugBuf);

    IEnumWbemClassObjectPtr pEnum = nullptr;

    HRESULT hr = g_pSvc->ExecQuery(
        _bstr_t(L"WQL"),
        _bstr_t(wqlQuery.c_str()),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnum
    );

    if (SUCCEEDED(hr) && pEnum) {
        IWbemClassObjectPtr pObj = nullptr;
        ULONG uRet = 0;

        if (pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet) == S_OK && uRet > 0) {
            _variant_t vtPath;

            hr = pObj->Get(L"__RELPATH", 0, &vtPath, 0, 0);

            if (FAILED(hr) || vtPath.vt != VT_BSTR) {
                pObj->Get(L"__PATH", 0, &vtPath, 0, 0);
            }

            if (SUCCEEDED(hr) && vtPath.vt == VT_BSTR) {
                g_cachedSensorPath = vtPath.bstrVal;
                g_pathCached = true;

                char pathLog[512];
                // Cast to wchar_t* for safety
                snprintf(pathLog, sizeof(pathLog), "[MysticFight] LHM WMI PATH CACHED -> %ls", (wchar_t*)g_cachedSensorPath);
                Log(pathLog);
            }
            else {
                Log("[MysticFight] Error: Object found but without a valid path.");
            }
        }
        else {
            Log("[MysticFight] Error: Sensor not found on LHM (Is it running?)");
        }
    }
    else {
        Log("[MysticFight] Error: WMI Query execution failed.");
    }
}

// Calculates color brightness to ensure text contrast in UI.
static bool IsColorDark(COLORREF col) {
    double brightness = (GetRValue(col) * 299 + GetGValue(col) * 587 + GetBValue(col) * 114) / 1000.0;
    return brightness < 128.0;
}

// =============================================================
// WINDOW PROCEDURE: SETTINGS DIALOG
// =============================================================
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
                if (IsColorDark(c)) SetTextColor(hdc, RGB(255, 255, 255));
                else SetTextColor(hdc, RGB(0, 0, 0));
                return (INT_PTR)hSelectedBrush;
            }
        }
        return (INT_PTR)GetStockObject(WHITE_BRUSH);
    }

    case WM_INITDIALOG: {
        RECT rcDlg;
        GetWindowRect(hDlg, &rcDlg);

        int dwWidth = rcDlg.right - rcDlg.left;
        int dwHeight = rcDlg.bottom - rcDlg.top;

        int x = (GetSystemMetrics(SM_CXSCREEN) - dwWidth) / 2;
        int y = (GetSystemMetrics(SM_CYSCREEN) - dwHeight) / 2;

        SetWindowPos(hDlg, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE);
        
        PopulateSensorList(hDlg);
        PopulateDeviceList(hDlg);

        SetDlgItemTextW(hDlg, IDC_EDIT_SERVER, g_cfg.webServerUrl);

        // Populate initial Area list based on current Device selection
        int devIdx = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_DEVICE), CB_GETCURSEL, 0, 0);
        if (devIdx != CB_ERR) {
            wchar_t* devType = (wchar_t*)SendMessage(GetDlgItem(hDlg, IDC_COMBO_DEVICE), CB_GETITEMDATA, devIdx, 0);
            if (devType) PopulateAreaList(hDlg, devType);
        }

        SetDlgItemInt(hDlg, IDC_TEMP_LOW, g_cfg.tempLow, TRUE);
        SetDlgItemInt(hDlg, IDC_TEMP_MED, g_cfg.tempMed, TRUE);
        SetDlgItemInt(hDlg, IDC_TEMP_HIGH, g_cfg.tempHigh, TRUE);

        wchar_t hexBuf[10];
        ColorToHex(g_cfg.colorLow, hexBuf, 10); SetDlgItemTextW(hDlg, IDC_HEX_LOW, hexBuf);
        ColorToHex(g_cfg.colorMed, hexBuf, 10); SetDlgItemTextW(hDlg, IDC_HEX_MED, hexBuf);
        ColorToHex(g_cfg.colorHigh, hexBuf, 10); SetDlgItemTextW(hDlg, IDC_HEX_HIGH, hexBuf);

        hBrushLow = CreateSolidBrush(g_cfg.colorLow);
        hBrushMed = CreateSolidBrush(g_cfg.colorMed);
        hBrushHigh = CreateSolidBrush(g_cfg.colorHigh);

        if (ValidStartupTaskExists()) CheckDlgButton(hDlg, IDC_CHK_STARTUP, BST_CHECKED);
        return (INT_PTR)TRUE;
    }

    case WM_COMMAND: {
        int id = LOWORD(wParam);
        int code = HIWORD(wParam);

        if (id == IDC_COMBO_DEVICE && code == CBN_SELCHANGE) {
            int idx = (int)SendMessage((HWND)lParam, CB_GETCURSEL, 0, 0);
            if (idx != CB_ERR) {
                wchar_t* devType = (wchar_t*)SendMessage((HWND)lParam, CB_GETITEMDATA, idx, 0);
                if (devType) PopulateAreaList(hDlg, devType);
            }
            return TRUE;
        }

        if (code == EN_CHANGE) {
            wchar_t buf[10];
            GetDlgItemTextW(hDlg, id, buf, 10);
            if (IsValidHex(buf)) {
                COLORREF newCol = HexToColor(buf);
                HBRUSH newBrush = CreateSolidBrush(newCol);
                if (id == IDC_HEX_LOW) { if (hBrushLow) DeleteObject(hBrushLow); hBrushLow = newBrush; }
                else if (id == IDC_HEX_MED) { if (hBrushMed) DeleteObject(hBrushMed); hBrushMed = newBrush; }
                else if (id == IDC_HEX_HIGH) { if (hBrushHigh) DeleteObject(hBrushHigh); hBrushHigh = newBrush; }
                InvalidateRect(GetDlgItem(hDlg, id), NULL, TRUE);
            }
            return TRUE;
        }

        switch (id) {
        case IDOK: {
            // Threshold validation
            int tL = GetDlgItemInt(hDlg, IDC_TEMP_LOW, NULL, TRUE);
            int tM = GetDlgItemInt(hDlg, IDC_TEMP_MED, NULL, TRUE);
            int tH = GetDlgItemInt(hDlg, IDC_TEMP_HIGH, NULL, TRUE);

            if (tL < 0 || tH > 110 || tL >= tM || tM >= tH) {
                MessageBoxW(hDlg, L"Invalid temperature range.", L"Error", MB_ICONWARNING);
                return TRUE;
            }

            GetDlgItemTextW(hDlg, IDC_EDIT_SERVER, g_cfg.webServerUrl, 256);

            // Save hardware selection
            int idxDev = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_DEVICE), CB_GETCURSEL, 0, 0);
            if (idxDev != CB_ERR) {
                wchar_t* devName = (wchar_t*)SendMessage(GetDlgItem(hDlg, IDC_COMBO_DEVICE), CB_GETITEMDATA, idxDev, 0);
                if (devName) wcscpy_s(g_cfg.targetDevice, devName);
            }
            int idxArea = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_AREA), CB_GETCURSEL, 0, 0);
            if (idxArea != CB_ERR) g_cfg.targetLedIndex = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_AREA), CB_GETITEMDATA, idxArea, 0);

            // Save sensor selection
            int idxSens = (int)SendMessage(GetDlgItem(hDlg, IDC_SENSOR_ID), CB_GETCURSEL, 0, 0);
            if (idxSens != CB_ERR) {
                wchar_t* sID = (wchar_t*)SendMessage(GetDlgItem(hDlg, IDC_SENSOR_ID), CB_GETITEMDATA, idxSens, 0);
                if (sID) wcscpy_s(g_cfg.sensorID, sID);
            }

            g_cfg.tempLow = tL; g_cfg.tempMed = tM; g_cfg.tempHigh = tH;
            wchar_t hL[10], hM[10], hH[10];
            GetDlgItemTextW(hDlg, IDC_HEX_LOW, hL, 10); g_cfg.colorLow = HexToColor(hL);
            GetDlgItemTextW(hDlg, IDC_HEX_MED, hM, 10); g_cfg.colorMed = HexToColor(hM);
            GetDlgItemTextW(hDlg, IDC_HEX_HIGH, hH, 10); g_cfg.colorHigh = HexToColor(hH);

            SetStartupTask(IsDlgButtonChecked(hDlg, IDC_CHK_STARTUP) == BST_CHECKED);
            
            SaveSettings();
            
            g_target_device = g_cfg.targetDevice;
            g_activeSource = DataSource::Searching;

            lastR = RGB_LED_REFRESH; // Force hardware update
            
            EndDialog(hDlg, IDOK);
            return TRUE;
        }
        case IDCANCEL: EndDialog(hDlg, IDCANCEL); return TRUE;
        case IDC_BTN_RESET:
            SetDlgItemInt(hDlg, IDC_TEMP_LOW, 50, TRUE);
            SetDlgItemInt(hDlg, IDC_TEMP_MED, 70, TRUE);
            SetDlgItemInt(hDlg, IDC_TEMP_HIGH, 90, TRUE);
            SetDlgItemTextW(hDlg, IDC_HEX_LOW, L"#00FF00");
            SetDlgItemTextW(hDlg, IDC_HEX_MED, L"#FFFF00");
            SetDlgItemTextW(hDlg, IDC_HEX_HIGH, L"#FF0000");
            return TRUE;
        }
        break;
    }
    case WM_DESTROY:
        if (hBrushLow) DeleteObject(hBrushLow);
        if (hBrushMed) DeleteObject(hBrushMed);
        if (hBrushHigh) DeleteObject(hBrushHigh);
        ClearComboHeapData(GetDlgItem(hDlg, IDC_SENSOR_ID));
        ClearComboHeapData(GetDlgItem(hDlg, IDC_COMBO_DEVICE));
        break;
    }
    return (INT_PTR)FALSE;
}

// =============================================================
// WINDOW PROCEDURE: ABOUT DIALOG
// =============================================================
INT_PTR CALLBACK AboutDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    static HFONT hFontLink = NULL;

    switch (message) {
    case WM_INITDIALOG: {
        RECT rcDlg;
        GetWindowRect(hDlg, &rcDlg);

        int dwWidth = rcDlg.right - rcDlg.left;
        int dwHeight = rcDlg.bottom - rcDlg.top;

        int x = (GetSystemMetrics(SM_CXSCREEN) - dwWidth) / 2;
        int y = (GetSystemMetrics(SM_CYSCREEN) - dwHeight) / 2;

        SetWindowPos(hDlg, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE);
        HICON hIcon = (HICON)LoadImage(GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_ICON1), IMAGE_ICON, 48, 48, LR_SHARED);
        SendMessage(GetDlgItem(hDlg, IDC_ABOUT_ICON), STM_SETIMAGE, IMAGE_ICON, (LPARAM)hIcon);

        std::wstring versionStr = L"MysticFight " + std::wstring(APP_VERSION);
        SetDlgItemTextW(hDlg, IDC_ABOUT_VERSION, versionStr.c_str());

        HFONT hFont = (HFONT)SendMessage(GetDlgItem(hDlg, IDC_GITHUB_LINK), WM_GETFONT, 0, 0);
        LOGFONT lf;
        GetObject(hFont, sizeof(LOGFONT), &lf);
        lf.lfUnderline = TRUE;
        hFontLink = CreateFontIndirect(&lf);
        SendMessage(GetDlgItem(hDlg, IDC_GITHUB_LINK), WM_SETFONT, (WPARAM)hFontLink, TRUE);
        return (INT_PTR)TRUE;
    }

    case WM_CTLCOLORSTATIC: {
        if ((HWND)lParam == GetDlgItem(hDlg, IDC_GITHUB_LINK)) {
            SetTextColor((HDC)wParam, RGB(0, 102, 204));
            SetBkMode((HDC)wParam, TRANSPARENT);
            return (INT_PTR)GetStockObject(NULL_BRUSH);
        }
        break;
    }

    case WM_SETCURSOR: {
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
    case WM_DESTROY:
        if (hFontLink) {
            DeleteObject(hFontLink);
            hFontLink = NULL;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// =============================================================
// RESOURCE CLEANUP
// =============================================================
static void FinalCleanup(HWND hWnd) {
    static bool cleaned = false;
    if (cleaned) return;
    cleaned = true;

    Log("[MysticFight] Starting cleaning...");

    // 1. Hardware Shutdown
    if (g_hLibrary) {
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

    // 2. Global BSTR Cleanup
    if (g_deviceName) {
        SysFreeString(g_deviceName);
        g_deviceName = NULL;
    }

    // 3. UI Cleanup
    if (hWnd) {
        NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA), hWnd, 1 };
        Shell_NotifyIcon(NIM_DELETE, &nid);
        UnregisterHotKey(hWnd, 1);
    }

    // 4. Smart Pointers (WMI)
    g_pSvc = nullptr;
    g_pLoc = nullptr;

    // 5. Synchronization
    if (g_hMutex) {
        ReleaseMutex(g_hMutex);
        CloseHandle(g_hMutex);
        g_hMutex = NULL;
    }

    // 6. COM Uninitialization
    CoUninitialize();

    Log("[MysticFight] Cleaning finished.");
}

// =============================================================
// WINDOW PROCEDURE: MAIN
// =============================================================
static LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_HOTKEY:
        if (wParam == 1) {
            g_LedsEnabled = !g_LedsEnabled;
            
            // Play sound from Resources
            if (g_LedsEnabled) {
                PlaySound(MAKEINTRESOURCE(IDR_WAV_LIGHTS_ON), GetModuleHandle(NULL),
                    SND_RESOURCE | SND_ASYNC);
            }
            else {
                PlaySound(MAKEINTRESOURCE(IDR_WAV_LIGHTS_OFF), GetModuleHandle(NULL),
                    SND_RESOURCE | SND_ASYNC);
            }
            
            lastR = RGB_LED_REFRESH;
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

// Retrieves CPU Temperature using the active strategy (WMI or HTTP)
static float GetCPUTempFast() {
    ULONGLONG currentTime = GetTickCount64();

    float temp = -1.0f;

    // CASE 1: LOCKED ON WMI
    if (g_activeSource == DataSource::WMI) {
        bool success = false;
        
        if (g_pSvc && g_pathCached) {
            try {
                IWbemClassObjectPtr pObj = nullptr;
                if (SUCCEEDED(g_pSvc->GetObject(g_cachedSensorPath, 0, NULL, &pObj, NULL)) && pObj) {
                    _variant_t vtVal;
                    if (SUCCEEDED(pObj->Get(L"Value", 0, &vtVal, 0, 0))) {
                        if (vtVal.vt == VT_R4) temp = vtVal.fltVal;
                        else if (vtVal.vt == VT_R8) temp = (float)vtVal.dblVal;
                        success = true;
                    }
                }
            }
            catch (...) { success = false; }
        }

        if (!success) {
            Log("[MysticFight] WMI reading failed. Resetting to SEARCH mode.");
            g_activeSource = DataSource::Searching;
            g_pathCached = false;
            return -1.0f;
        }
        return temp;
    }

    // CASE 2: LOCKED ON HTTP
    else if (g_activeSource == DataSource::HTTP) {
        std::string json = FetchLHMJson();
        
        if (!json.empty()) {
            temp = ParseLHMJsonForTemp(json, g_cfg.sensorID);
        }

        if (temp < 0.0f) {
            Log("[MysticFight] HTTP reading failed. Resetting to SEARCH mode.");
            g_activeSource = DataSource::Searching; // Unlock to try alternatives next frame
            return -1.0f;
        }
        return temp;
    }

    // CASE 3: SEARCHING (First run or after failure)
    else {
        
        if (currentTime - g_lastDataSourceSearchRetry <= LHM_RETRY_DELAY_MS) {
            return -1.0f;
        }

        g_lastDataSourceSearchRetry = currentTime;
        
        Log("[MysticFight] Attempting fresh data source discovery...");

        // A) Try WMI First
        if (InitWMI()) {
            
            CacheSensorPath();

            // El fallo estaba aquí: g_pathCached debe ser TRUE para bloquearse en WMI
            if (g_pathCached) {
                g_activeSource = DataSource::WMI;
                Log("[MysticFight] Data Source detected: WMI.");
                return GetCPUTempFast();
            }
            else {
                Log("[MysticFight] WMI connected but target sensor not found. Trying HTTP...");
            }
        }
        else {
            Log("[MysticFight] WMI Service not available. Trying HTTP...");
        }

        // B) Try HTTP Second (if WMI failed)
        std::string json = FetchLHMJson();
        
        if (!json.empty()) {
            temp = ParseLHMJsonForTemp(json, g_cfg.sensorID);
            if (temp >= 0.0f) {
                // SUCCESS: Lock onto HTTP
                g_activeSource = DataSource::HTTP;
                Log("[MysticFight] Data Source detected: HTTP.");
                g_pSvc = nullptr;
                g_pLoc = nullptr;
                return temp;
            }
            else {
                Log("[MysticFight] HTTP connected but target sensor not found.");
            }
        }
        else {
            Log("[MysticFight] HTTP Service not available. Is Web Server enabled on LHM?");
        }
    }

    return -1.0f;
}

static void SensorThread() {
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

    while (g_Running) {
        float temp = GetCPUTempFast();
        g_asyncTemp.store(temp);

        // Espera de forma eficiente. Se despierta si pasan 500ms 
        // O si señalizamos el evento g_hSensorEvent al cerrar.
        WaitForSingleObject(g_hSensorEvent, (DWORD)SENSOR_POLL_INTERVAL_MS);
    }

    CoUninitialize();
}

// Dumps all available temperature sensors to the log file
// Adapts automatically to the active source (WMI or HTTP)
static void LogAllLHMTemperatureSensors() {
    Log("[MysticFight] --- DUMPING ALL LHM TEMPERATURE SENSORS ---");

    // 1. WMI DUMP STRATEGY
    // We try this if we are locked to WMI or if we are still searching/determining
    if (g_activeSource == DataSource::WMI || g_activeSource == DataSource::Searching) {
        if (g_pSvc) {
            IEnumWbemClassObjectPtr pEnum = nullptr;
            HRESULT hr = g_pSvc->ExecQuery(
                bstr_t("WQL"),
                bstr_t("SELECT Identifier, Name, Value FROM Sensor WHERE SensorType = 'Temperature'"),
                WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
                NULL,
                &pEnum
            );

            if (SUCCEEDED(hr) && pEnum) {
                IWbemClassObjectPtr pObj = nullptr;
                ULONG uRet = 0;
                bool foundAny = false;

                while (pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet) == S_OK) {
                    _variant_t vtID, vtName, vtVal;
                    pObj->Get(L"Identifier", 0, &vtID, 0, 0);
                    pObj->Get(L"Name", 0, &vtName, 0, 0);
                    pObj->Get(L"Value", 0, &vtVal, 0, 0);

                    float val = 0.0f;
                    if (vtVal.vt == VT_R4) val = vtVal.fltVal;
                    else if (vtVal.vt == VT_R8) val = (float)vtVal.dblVal;

                    char buf[LOG_BUFFER_SIZE];
                    // Corrected format: Starts with [MysticFight]
                    snprintf(buf, sizeof(buf), "[MysticFight] [WMI] ID: %ls | Name: %ls | Value: %.1f",
                        (vtID.vt == VT_BSTR ? vtID.bstrVal : L"N/A"),
                        (vtName.vt == VT_BSTR ? vtName.bstrVal : L"N/A"),
                        val);
                    Log(buf);
                    foundAny = true;
                }
                if (foundAny) {
                    Log("[MysticFight] --- END OF WMI DUMP ---");
                    return; // If WMI worked, we are done
                }
            }
        }
    }

    // 2. HTTP DUMP STRATEGY
    // We try this if WMI returned nothing OR if we are locked to HTTP
    if (g_activeSource == DataSource::HTTP || g_activeSource == DataSource::Searching) {
        std::string json = FetchLHMJson();
        if (json.empty()) {
            Log("[MysticFight] [Dump] HTTP JSON is empty or unreachable.");
            return;
        }

        std::string typeKey = "\"Type\":\"Temperature\"";
        size_t pos = 0;
        bool foundAny = false;

        while ((pos = json.find(typeKey, pos)) != std::string::npos) {
            // Find object boundaries
            size_t blockStart = json.rfind('{', pos);
            size_t blockEnd = json.find('}', pos);

            if (blockStart != std::string::npos && blockEnd != std::string::npos) {
                std::string block = json.substr(blockStart, blockEnd - blockStart + 1);

                std::wstring name = ExtractJsonString(block, "Text");
                std::wstring id = ExtractJsonString(block, "SensorId");
                std::wstring valStr = ExtractJsonString(block, "Value");

                if (!id.empty()) {
                    char buf[LOG_BUFFER_SIZE];
                    // Corrected format: Starts with [MysticFight]
                    snprintf(buf, sizeof(buf), "[MysticFight] [HTTP] ID: %ls | Name: %ls | Value: %ls",
                        id.c_str(), name.c_str(), valStr.c_str());
                    Log(buf);
                    foundAny = true;
                }
            }
            pos = blockEnd;
        }

        if (foundAny) Log("[MysticFight] --- END OF HTTP DUMP ---");
        else Log("[MysticFight] [Dump] No temperature sensors found in HTTP stream.");
    }
}

static void ShowNotification(HWND hWnd, HINSTANCE hInstance, const wchar_t* title, const wchar_t* info) {
    NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA), hWnd, 1, NIF_INFO };
    wcscpy_s(nid.szInfoTitle, title);
    wcscpy_s(nid.szInfo, info);
    nid.dwInfoFlags = NIIF_USER | NIIF_LARGE_ICON;
    Shell_NotifyIcon(NIM_MODIFY, &nid);
}



// Extracts embedded MSI DLL resource to disk.
static bool ExtractMSIDLL() {
    wchar_t szPath[MAX_PATH];
    GetModuleFileNameW(NULL, szPath, MAX_PATH);
    wchar_t* lastSlash = wcsrchr(szPath, L'\\');
    if (lastSlash) *(lastSlash + 1) = L'\0';

    std::wstring dllPath = std::wstring(szPath) + L"MysticLight_SDK.dll";

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

static void MSIHwardwareDetection() {
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

                char devInfo[LOG_BUFFER_SIZE];
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
                        char ledLine[LOG_BUFFER_SIZE];
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
}


// Automatically selects the first available temperature sensor if none is configured.
static bool AutoSelectFirstSensor() {
    if (!g_pSvc) return false;

    IEnumWbemClassObjectPtr pEnum = nullptr;
    // Query for any sensor of type Temperature
    HRESULT hr = g_pSvc->ExecQuery(
        _bstr_t(L"WQL"),
        _bstr_t(L"SELECT Identifier FROM Sensor WHERE SensorType = 'Temperature'"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnum
    );

    if (SUCCEEDED(hr) && pEnum) {
        IWbemClassObjectPtr pObj = nullptr;
        ULONG uRet = 0;

        // We only need the very first result
        if (pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet) == S_OK && uRet > 0) {
            _variant_t vtID;
            hr = pObj->Get(L"Identifier", 0, &vtID, 0, 0);

            if (SUCCEEDED(hr) && vtID.vt == VT_BSTR) {
                // 1. Update Global Config
                wcscpy_s(g_cfg.sensorID, vtID.bstrVal);

                // 2. Persist to INI immediately
                WritePrivateProfileStringW(L"Settings", L"SensorID", g_cfg.sensorID, INI_FILE);

                char logBuf[LOG_BUFFER_SIZE];
                snprintf(logBuf, sizeof(logBuf), "[MysticFight] Auto-selected default sensor: %ls", g_cfg.sensorID);
                Log(logBuf);

                return true;
            }
        }
    }

    Log("[MysticFight] Auto-select failed: No temperature sensors found in LHM.");
    return false;
}

// =============================================================
// MAIN ENTRY POINT
// =============================================================
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {

    // --- A. PRIORITIES & COM INITIALIZATION ---
    SetPriorityClass(GetCurrentProcess(), HIGH_PRIORITY_CLASS);
    SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);

    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (FAILED(hr)) {
        Log("[MysticFight] Critical: COM Initialization failed.");
        return 1;
    }

    // --- B. SINGLE INSTANCE CHECK ---
    g_hMutex = CreateMutex(NULL, TRUE, L"Global\\MysticFight_Unique_Mutex");
    if (g_hMutex == NULL || GetLastError() == ERROR_ALREADY_EXISTS) {
        if (g_hMutex) {
            MessageBoxW(NULL, L"MysticFight is already running.", L"Information", MB_OK | MB_ICONINFORMATION);
            CloseHandle(g_hMutex);
        }
        CoUninitialize();
        return 0;
    }

    // --- C. COMPONENT PREPARATION (DLL) ---
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

    // --- D. LOGGING & CONFIGURATION ---
    TrimLogFile();

    char versionMsg[128];
    snprintf(versionMsg, sizeof(versionMsg), "[MysticFight] MysticFight %ls started", APP_VERSION);
    Log(versionMsg);

    LoadSettings();
    g_target_device = g_cfg.targetDevice;

    wchar_t hL[10], hM[10], hH[10];
    ColorToHex(g_cfg.colorLow, hL, 10);
    ColorToHex(g_cfg.colorMed, hM, 10);
    ColorToHex(g_cfg.colorHigh, hH, 10);

    char startupCfg[LOG_BUFFER_SIZE];
    snprintf(startupCfg, sizeof(startupCfg),
        "[MysticFight] Config Loaded - Device: %ls | LED Area index: %d | Sensor: %ls | Low: %dºC (%ls) | Med: %dºC (%ls) | High: %dºC (%ls)",
        g_cfg.targetDevice,
        g_cfg.targetLedIndex,
        g_cfg.sensorID,
        g_cfg.tempLow, hL,
        g_cfg.tempMed, hM,
        g_cfg.tempHigh, hH
    );
    Log(startupCfg);

    // --- E. WINDOW REGISTRATION ---
    wchar_t windowTitle[100];
    swprintf_s(windowTitle, L"MysticFight %s (by tonikelope)", APP_VERSION);

    WNDCLASS wc = { 0 };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"MysticFight_Class";
    wc.hIcon = (HICON)LoadImage(hInstance, MAKEINTRESOURCE(101), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);

    RegisterClass(&wc);

    HWND hWnd = CreateWindowEx(0, wc.lpszClassName, windowTitle, 0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);

    // Register Tray Icon
    NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA), hWnd, 1, NIF_ICON | NIF_MESSAGE | NIF_TIP, WM_TRAYICON };
    nid.hIcon = (HICON)LoadImage(hInstance, MAKEINTRESOURCE(101), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
    swprintf_s(nid.szTip, L"MysticFight %s (by tonikelope)", APP_VERSION);
    Shell_NotifyIcon(NIM_ADD, &nid);

    RegisterHotKey(hWnd, 1, MOD_CONTROL | MOD_SHIFT | MOD_ALT, 0x4C);
    ShowNotification(hWnd, hInstance, windowTitle, L"Let's dance baby");

    // --- F. BIND DLL FUNCTIONS ---
    lpMLAPI_Initialize = (LPMLAPI_Initialize)GetProcAddress(g_hLibrary, "MLAPI_Initialize");
    lpMLAPI_GetDeviceInfo = (LPMLAPI_GetDeviceInfo)GetProcAddress(g_hLibrary, "MLAPI_GetDeviceInfo");
    lpMLAPI_SetLedColor = (LPMLAPI_SetLedColor)GetProcAddress(g_hLibrary, "MLAPI_SetLedColor");
    lpMLAPI_SetLedStyle = (LPMLAPI_SetLedStyle)GetProcAddress(g_hLibrary, "MLAPI_SetLedStyle");
    lpMLAPI_SetLedSpeed = (LPMLAPI_SetLedSpeed)GetProcAddress(g_hLibrary, "MLAPI_SetLedSpeed");
    lpMLAPI_Release = (LPMLAPI_Release)GetProcAddress(g_hLibrary, "MLAPI_Release");
    lpMLAPI_GetDeviceNameEx = (LPMLAPI_GetDeviceNameEx)GetProcAddress(g_hLibrary, "MLAPI_GetDeviceNameEx");
    lpMLAPI_GetLedInfo = (LPMLAPI_GetLedInfo)GetProcAddress(g_hLibrary, "MLAPI_GetLedInfo");

    // --- G. SDK INITIALIZATION ---
    Log("[MysticFight] Attempting to initialize SDK...");
    bool sdkReady = false;

    if (!lpMLAPI_Initialize || lpMLAPI_Initialize() != 0) {
        Log("[MysticFight] SDK not available at startup. Trying to start on background...");
        g_Resetting_sdk = true;
        g_ResetStage = 0;
        g_ResetTimer = 0;
    }
    else {
        Log("[MysticFight] SDK Initialized successfully at startup");
        sdkReady = true;
    }

    // --- H. HARDWARE DETECTION BLOCK ---
    if (sdkReady) {
        MSIHwardwareDetection();
    }

    // --- I. MAIN LOOP PREPARATION ---
    _bstr_t bstrOff(L"Off");
    _bstr_t bstrBreath(L"Breath");
    _bstr_t bstrSteady(L"Steady");

    // "Current" Floating point values for smooth precision (Animation State)
    float currR = 0.0f, currG = 0.0f, currB = 0.0f;

    // "Target" Integer values (Where we want to go based on Temperature)
    DWORD targetR = 0, targetG = 0, targetB = 0;

    // Timer state for Sensor reading
    ULONGLONG lastSensorReadTime = 0;

    // State flags
    MSG msg = { 0 };
    bool lhmAlive = true;  // Tracks if data source is healthy
    bool firstRun = true;  // Used to snap color instantly on startup (no fade in from black)

    g_hSensorEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
    std::thread sThread(SensorThread);
    sThread.detach();

    // =============================================================
    // MAIN APPLICATION LOOP
    // =============================================================
    while (g_Running) {

        // 1. PROCESS WINDOW MESSAGES (Standard Windows boilerplate)
        while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
            if (msg.message == WM_QUIT) { g_Running = false; break; }
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        if (!g_Running) break;

        ULONGLONG currentTime = GetTickCount64();

        // 2. WATCHDOG: SDK RESET STATE MACHINE (Keep your existing logic mostly generic)
        if (g_Resetting_sdk) {
            if (currentTime >= g_ResetTimer) {
                switch (g_ResetStage) {
                case 0: // Kill Process
                    Log("[MysticFight] Reset Stage 0: Cleaning MSI...");
                    ControlScheduledTask(L"MSI Task Host - LEDKeeper2_Host", false);
                    system("taskkill /F /IM LEDKeeper2.exe /T > nul 2>&1");
                    g_ResetTimer = currentTime + RESET_KILL_TASK_WAIT_MS;
                    g_ResetStage = 1;
                    break;
                case 1: // Restart Service
                    Log("[MysticFight] Reset Stage 1: Restarting MSI...");
                    ControlScheduledTask(L"MSI Task Host - LEDKeeper2_Host", true);
                    g_ResetTimer = currentTime + RESET_RESTART_TASK_DELAY_MS;
                    g_ResetStage = 2;
                    break;
                case 2: // Re-initialize SDK
                    Log("[MysticFight] Reset Stage 2: Re-initializing SDK...");
                    if (lpMLAPI_Initialize && lpMLAPI_Initialize() == 0) {
                        MSIHwardwareDetection();
                        // Restore Steady style immediately after reset
                        if (lpMLAPI_SetLedStyle) lpMLAPI_SetLedStyle(g_deviceName, 0, bstrSteady);
                    }
                    g_ResetTimer = currentTime + RESET_RESTART_TASK_DELAY_MS;
                    g_Resetting_sdk = false;

                    // IMPORTANT: Force hardware refresh after reset
                    lastR = RGB_LED_REFRESH; // Usamos la global
                    firstRun = true;
                    break;
                }
            }
        }

        // 3. MAIN LOGIC (Only if LEDs are enabled)
        else if (g_LedsEnabled) {

            // =========================================================
            // A. SENSOR POLLING & TARGET CALCULATION
            // Runs every SENSOR_POLL_INTERVAL_MS (e.g., 500ms)
            // =========================================================
            if (currentTime - lastSensorReadTime >= SENSOR_POLL_INTERVAL_MS || firstRun) {

                float rawTemp = g_asyncTemp.load();
                
                lastSensorReadTime = currentTime;

                if (rawTemp >= 0.0f) {
                    // RECOVERY: If we were disconnected, restore state
                    if (!lhmAlive) {
                        lhmAlive = true;
                        Log("[MysticFight] Connection established/recovered.");
                        lastR = RGB_LED_REFRESH;
                    }

                    // Round 0.5ºC (50.0, 50.5, 51.0...)
                    float temp = floorf(rawTemp * 2.0f + 0.5f) / 2.0f;

                    // --- COLOR MATH (RMS Calculation for Target) ---
                    float ratio = 0.0f;
                    COLORREF c1 = 0, c2 = 0;

                    if (temp <= (float)g_cfg.tempLow) {
                        targetR = GetRValue(g_cfg.colorLow);
                        targetG = GetGValue(g_cfg.colorLow);
                        targetB = GetBValue(g_cfg.colorLow);
                    }
                    else if (temp < (float)g_cfg.tempMed) {
                        ratio = (temp - (float)g_cfg.tempLow) / ((float)g_cfg.tempMed - (float)g_cfg.tempLow);

                        // FIX: Clamp ratio to prevent SQRT domain error (CRASH FIX)
                        if (ratio < 0.0f) ratio = 0.0f;
                        if (ratio > 1.0f) ratio = 1.0f;

                        c1 = g_cfg.colorLow; c2 = g_cfg.colorMed;

                        double r1 = (double)GetRValue(c1); double g1 = (double)GetGValue(c1); double b1 = (double)GetBValue(c1);
                        double r2 = (double)GetRValue(c2); double g2 = (double)GetGValue(c2); double b2 = (double)GetBValue(c2);

                        targetR = (DWORD)sqrt(r1 * r1 * (1.0 - ratio) + r2 * r2 * ratio);
                        targetG = (DWORD)sqrt(g1 * g1 * (1.0 - ratio) + g2 * g2 * ratio);
                        targetB = (DWORD)sqrt(b1 * b1 * (1.0 - ratio) + b2 * b2 * ratio);
                    }
                    else if (temp < (float)g_cfg.tempHigh) {
                        ratio = (temp - (float)g_cfg.tempMed) / ((float)g_cfg.tempHigh - (float)g_cfg.tempMed);

                        // FIX: Clamp ratio to prevent SQRT domain error (CRASH FIX)
                        if (ratio < 0.0f) ratio = 0.0f;
                        if (ratio > 1.0f) ratio = 1.0f;

                        c1 = g_cfg.colorMed; c2 = g_cfg.colorHigh;

                        double r1 = (double)GetRValue(c1); double g1 = (double)GetGValue(c1); double b1 = (double)GetBValue(c1);
                        double r2 = (double)GetRValue(c2); double g2 = (double)GetGValue(c2); double b2 = (double)GetBValue(c2);

                        targetR = (DWORD)sqrt(r1 * r1 * (1.0 - ratio) + r2 * r2 * ratio);
                        targetG = (DWORD)sqrt(g1 * g1 * (1.0 - ratio) + g2 * g2 * ratio);
                        targetB = (DWORD)sqrt(b1 * b1 * (1.0 - ratio) + b2 * b2 * ratio);
                    }
                    else {
                        targetR = GetRValue(g_cfg.colorHigh);
                        targetG = GetGValue(g_cfg.colorHigh);
                        targetB = GetBValue(g_cfg.colorHigh);
                    }

                    // On startup, snap directly to target (skip fade-in)
                    if (firstRun) {
                        currR = (float)targetR;
                        currG = (float)targetG;
                        currB = (float)targetB;
                        firstRun = false;
                    }
                }
            }

            // =========================================================
            // B. INTERPOLATION & RENDER
            // Runs every MAIN_LOOP_DELAY_MS (e.g., 50ms)
            // =========================================================
            if (lhmAlive) {
                // 1. Math: Move 'Current' towards 'Target'
                
                currR += ((float)targetR - currR) * SMOOTHING_FACTOR;
                currG += ((float)targetG - currG) * SMOOTHING_FACTOR;
                currB += ((float)targetB - currB) * SMOOTHING_FACTOR;
                
                // 2. Convert Float to Int for Hardware
                DWORD sendR = (DWORD)currR;
                DWORD sendG = (DWORD)currG;
                DWORD sendB = (DWORD)currB;

                // 3. Send to MSI SDK (Only if value changed visibly)
                // FIX: Usamos las variables globales lastR/G/B para coordinar con el Dialog
                if (sendR != lastR || sendG != lastG || sendB != lastB) {

                    int status = 0;

                    // Safety: Ensure we are in 'Steady' mode if we just came from 'Off'
                    if (lastR == RGB_LEDS_OFF || lastR == RGB_LED_REFRESH) {
                        if (lpMLAPI_SetLedStyle) status = lpMLAPI_SetLedStyle(g_target_device, g_cfg.targetLedIndex, bstrSteady);
                    }

                    if (status == 0 && lpMLAPI_SetLedColor) {
                        status = lpMLAPI_SetLedColor(g_target_device, g_cfg.targetLedIndex, sendR, sendG, sendB);
                    }

                    if (status != 0) {
                        // If hardware fails, trigger the Watchdog
                        Log("[MysticFight] SDK SetColor failed. Triggering Reset...");
                        g_Resetting_sdk = true; g_ResetStage = 0; g_ResetTimer = 0;
                    }
                    else {
                        // Success: Update tracking (Globals)
                        lastR = sendR; lastG = sendG; lastB = sendB;
                    }
                }
            }

        } // End g_LedsEnabled

        // 4. OFF MODE LOGIC
        else {
            if (lastR != RGB_LEDS_OFF) {
                int status = 0;
                if (lpMLAPI_SetLedStyle) status = lpMLAPI_SetLedStyle(g_deviceName, 0, bstrOff);

                if (status != 0) {
                    g_Resetting_sdk = true; g_ResetStage = 0; g_ResetTimer = 0;
                }
                else {
                    lastR = RGB_LEDS_OFF;
                }
            }
        }

        // 5. WAIT FOR NEXT FRAME
        // This controls the animation speed (e.g., 50ms)
        MsgWaitForMultipleObjects(0, NULL, FALSE, (DWORD)MAIN_LOOP_DELAY_MS, QS_ALLINPUT);
    }

    SetEvent(g_hSensorEvent);

    // --- 4. CLEANUP ---
    FinalCleanup(hWnd);
    Log("[MysticFight] BYE BYE");

    return 0;
}