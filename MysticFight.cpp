/**
 * ============================================================================
 * MysticFight - RGB Control & Hardware Monitor Integration
 * Author: Gemini (& tonikelope, at least i want to believe it).
 *
 * Description:
 * This application synchronizes MSI hardware RGB lighting with system temperatures
 * retrieved via LibreHardwareMonitor (LHM) JSON API.
 * ============================================================================
 */

#include "resource.h" // Ensure this exists in your project

 // -----------------------------------------------------------------------------
 // STANDARD LIBRARY INCLUDES
 // -----------------------------------------------------------------------------
#include <algorithm>
#include <atomic>
#include <ctime>
#include <fstream>
#include <math.h>
#include <mutex>
#include <string>
#include <thread>
#include <vector>
#include <map>

// -----------------------------------------------------------------------------
// JSON LIBRARY (MODERN PARSING)
// -----------------------------------------------------------------------------
#include <nlohmann/json.hpp>
using json = nlohmann::json;

// -----------------------------------------------------------------------------
// WINDOWS & COM INCLUDES
// -----------------------------------------------------------------------------
#include <windows.h>
#include <comdef.h>
#include <comip.h>
#include <commdlg.h>
#include <commctrl.h>
#include <shellapi.h>
#include <shlobj.h>
#include <taskschd.h>
#include <tlhelp32.h>
#include <winhttp.h>

// Linker instructions for Modern UI Visual Styles
#pragma comment(linker,"\"/manifestdependency:type='win32' \
name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

// -----------------------------------------------------------------------------
// LIBRARY LINKING
// -----------------------------------------------------------------------------
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "winmm.lib")
#pragma comment(lib, "comctl32.lib")


// ============================================================================
// CONSTANTS & DEFINITIONS
// ============================================================================

// MSI SDK Error Constants
#define MLAPI_OK                 0
#define MLAPI_ERROR             -1
#define MLAPI_TIMEOUT           -2
#define MLAPI_NO_IMPLEMENTED    -3
#define MLAPI_NOT_INITIALIZED   -4
#define MLAPI_INVALID_ARGUMENT  -101
#define MLAPI_DEVICE_NOT_FOUND  -102
#define MLAPI_NOT_SUPPORTED     -103

// UI & Tray Identifiers
#define WM_TRAYICON             (WM_USER + 1)
#define ID_TRAY_EXIT            1001
#define ID_TRAY_CONFIG          2001
#define ID_TRAY_LOG             3001
#define ID_TRAY_ABOUT           4001
#define ID_TRAY_PROFILE_BASE    5000

// New: Profile Switching Menu IDs (5000 - 5004)
#define ID_TRAY_PROFILE_START   5000 

// Application Metadata
const wchar_t* APP_VERSION = L"v2.76";
const wchar_t* LOG_FILENAME = L"debug.log";
const wchar_t* INI_FILE = L".\\config.ini";
const wchar_t* TASK_NAME = L"MysticFight";

// Sentinel Values for LED State
const DWORD RGB_LED_REFRESH = 999;  // Force hardware update
const DWORD RGB_LEDS_OFF = 1000;    // Hardware status is OFF

// Timing & Limits
const ULONGLONG MAIN_LOOP_OFF_DELAY_MS = 1000;      // Polling rate when LEDs are OFF
const ULONGLONG SENSOR_LOOP_OFF_DELAY_MS = 5000;    // Background polling rate
const ULONGLONG LHM_BOOT_TIMEOUT = 5000;            // SDK*LHM INIT TIMEOUT
const DWORD     MAX_HTTP_RESPONSE_SIZE = 4194304;   // 4MB Safety Limit
const ULONGLONG LHM_RETRY_DELAY_MS = 5000;          // Cooldown for searching LHM
const ULONGLONG KILL_LEDKEEPER_TASK_WAIT_MS = 2000; // Watchdog termination wait
const ULONGLONG RESET_LEDKEEPER_TASK_DELAY_MS = 5000; // Cooldown for task restart
const ULONGLONG RESET_SDK_RETRY_DELAY_MS = 1000;
const ULONGLONG HTTP_FAST_TIMEOUT = 1000;           // UI-blocking timeout
const ULONGLONG HTTP_NORMAL_TIMEOUT = 60000;        // Background network timeout
const int       HEX_COLOR_LEN = 7;                  // Format: #RRGGBB
const int       LOG_BUFFER_SIZE = 512;
const int       SENSOR_ID_LEN = 256;

// Default Factory Settings
const int   DEF_SENSOR_UPDATE_MS = 500;
const int   DEF_LED_REFRESH_FPS = 25;
const float DEF_SMOOTHING_FACTOR = 0.150f;

const int   DEF_TEMP_LOW = 50;
const int   DEF_TEMP_MED = 70;
const int   DEF_TEMP_HIGH = 90;

const wchar_t* DEF_COLOR_LOW = L"#00FF00"; // Green
const wchar_t* DEF_COLOR_MED = L"#FFFF00"; // Yellow
const wchar_t* DEF_COLOR_HIGH = L"#FF0000"; // Red

const wchar_t* DEF_SERVER_URL = L"http://localhost:8085";
const wchar_t* DEF_DEVICE_TYPE = L"MSI_MB";


// ============================================================================
// DATA STRUCTURES & ENUMS
// ============================================================================

// Helper structure to handle SAFEARRAY locking automatically (RAII)
/**
 * SafeArrayScopedLock - A RAII wrapper for SAFEARRAY data access.
 * Ensures SafeArrayUnaccessData is called when the object goes out of scope.
 */
template <typename T>
struct SafeArrayScopedLock {
    SAFEARRAY* psa;
    T* data;
    HRESULT hr;

    SafeArrayScopedLock(SAFEARRAY* psa) : psa(psa), data(nullptr), hr(E_FAIL) {
        if (psa) {
            hr = SafeArrayAccessData(psa, (void**)&data);
        }
    }

    ~SafeArrayScopedLock() {
        if (psa && SUCCEEDED(hr)) {
            SafeArrayUnaccessData(psa);
        }
    }

    bool isLocked() const { return SUCCEEDED(hr) && data != nullptr; }
};

enum class DataSource {
    Searching,
    HTTP
};

struct Config {
    wchar_t sensorID[SENSOR_ID_LEN];
    wchar_t targetDevice[256];
    int targetLedIndex;
    int tempLow, tempMed, tempHigh;
    COLORREF colorLow, colorMed, colorHigh;
    wchar_t webServerUrl[256];
    wchar_t label[64];
    int sensorUpdateMS;
    int ledRefreshFPS;
    float smoothingFactor;
};

struct AppHotkeys {
    WORD toggleLEDs;
    WORD profile1;
    WORD profile2;
    WORD profile3;
    WORD profile4;
    WORD profile5;
};

struct GlobalConfig {
    Config profiles[5];
    int activeProfileIndex;
    AppHotkeys hotkeys;
};


// ============================================================================
// MSI SDK TYPE DEFINITIONS
// ============================================================================
typedef int (*LPMLAPI_Initialize)();
typedef int (*LPMLAPI_GetDeviceInfo)(SAFEARRAY** pDevType, SAFEARRAY** pLedCount);
typedef int (*LPMLAPI_GetDeviceNameEx)(BSTR type, DWORD index, BSTR* pDevName);
typedef int (*LPMLAPI_GetLedInfo)(BSTR, DWORD, BSTR*, SAFEARRAY**);
typedef int (*LPMLAPI_SetLedColor)(BSTR type, DWORD index, DWORD R, DWORD G, DWORD B);
typedef int (*LPMLAPI_SetLedStyle)(BSTR type, DWORD index, BSTR style);
typedef int (*LPMLAPI_SetLedSpeed)(BSTR type, DWORD index, DWORD level);
typedef int (*LPMLAPI_Release)();


// ============================================================================
// COM SMART POINTERS (Task Scheduler Wrapper)
// ============================================================================
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


// ============================================================================
// GLOBAL STATE VARIABLES
// ============================================================================

GlobalConfig g_Global;
std::mutex g_cfgMutex;
std::mutex g_httpMutex;

// Atomic Logic Flags
std::atomic<bool> g_ResetHttp{ false };
std::atomic<DataSource> g_activeSource{ DataSource::Searching };
std::atomic<bool> g_Running{ true };
std::atomic<bool> g_windows_shutdown{ false };
std::atomic<bool> g_LedsEnabled{ true };
std::atomic<float> g_asyncTemp{ -1.0f };
std::atomic<int> g_httpConnectionDrops{ 0 };
std::atomic<bool> g_SettingsOpen{ false };
std::atomic<bool> g_AboutOpen{ false };
std::atomic<bool> g_pendingStyleChange{ true };
std::atomic<int> g_ConfigVersion{ 0 };

// LED Real-time State Tracking
std::atomic<DWORD> lastR{ RGB_LED_REFRESH }, lastG{ RGB_LED_REFRESH }, lastB{ RGB_LED_REFRESH };
std::atomic<float> currR{ 0.0f }, currG{ 0.0f }, currB{ 0.0f };
std::atomic<float> lastTemp{ -1.0f };

// Thread Synchronization Handles
HANDLE g_hSensorEvent = NULL;
HANDLE g_hSourceResolvedEvent = NULL;
HANDLE g_hMutex = NULL;
HINTERNET g_hSession = NULL;
HINTERNET g_hConnect = NULL;

// Hardware SDK State
BSTR g_deviceName = NULL;
HMODULE g_hLibrary = NULL;
int g_totalLeds = 0;

ULONGLONG g_lastDataSourceSearchRetry = 0;

// SDK AUTO RESET
std::atomic<bool> g_Resetting_sdk{ false };
int g_ResetStage = 0;
ULONGLONG g_ResetTimer = 0;
int g_sdkFailCount = 0;
const int MAX_SDK_FAILURES = 5;
const int SDK_MAX_RESETS = 3;
int g_resetCounter = 0;

// SDK Function Pointers
LPMLAPI_Initialize      lpMLAPI_Initialize = nullptr;
LPMLAPI_GetDeviceInfo   lpMLAPI_GetDeviceInfo = nullptr;
LPMLAPI_SetLedColor     lpMLAPI_SetLedColor = nullptr;
LPMLAPI_SetLedStyle     lpMLAPI_SetLedStyle = nullptr;
LPMLAPI_SetLedSpeed     lpMLAPI_SetLedSpeed = nullptr;
LPMLAPI_Release         lpMLAPI_Release = nullptr;
LPMLAPI_GetDeviceNameEx lpMLAPI_GetDeviceNameEx = nullptr;
LPMLAPI_GetLedInfo      lpMLAPI_GetLedInfo = nullptr;

// Helper Accessor Macro for Thread Safety
inline Config& GetCfgSafe() { return g_Global.profiles[g_Global.activeProfileIndex]; }
#define g_cfg GetCfgSafe()


// ============================================================================
// UTILITY FUNCTIONS (Logging, Strings, & Privileges)
// ============================================================================

/**
 * Updates the tray icon tooltip to show real-time status
 */
void UpdateStatus(HWND hWnd, const wchar_t* status) {
    NOTIFYICONDATAW nid = { sizeof(nid), hWnd, 1, NIF_TIP };

    if (status != NULL) {
        swprintf_s(nid.szTip, L"MysticFight %ls (by tonikelope) - %ls", APP_VERSION, status);
    }
    else {
        swprintf_s(nid.szTip, L"MysticFight %ls (by tonikelope)", APP_VERSION);
    }

    Shell_NotifyIconW(NIM_MODIFY, &nid);
}

/**
 * Standard thread-safe logger to file
 */
static void Log(const char* text) {
    static std::mutex logMutex;
    std::lock_guard<std::mutex> lock(logMutex);

    time_t now = time(0);
    char dt[26];
    ctime_s(dt, sizeof(dt), &now);
    dt[24] = '\0';

    std::ofstream out(LOG_FILENAME, std::ios::app);
    if (out) {
        out << "[" << dt << "] " << text << "\n";
    }
}

/**
 * Log MSI SDK specific errors with human-readable descriptions
 */
static void LogSDKError(const char* functionName, int errorCode) {
    std::string errorDesc;
    switch (errorCode) {
    case MLAPI_ERROR:             errorDesc = "Generic error"; break;
    case MLAPI_TIMEOUT:           errorDesc = "Request timeout"; break;
    case MLAPI_NO_IMPLEMENTED:    errorDesc = "MSI application not found or version not supported"; break;
    case MLAPI_NOT_INITIALIZED:   errorDesc = "SDK not initialized"; break;
    case MLAPI_INVALID_ARGUMENT:  errorDesc = "Invalid parameter value"; break;
    case MLAPI_DEVICE_NOT_FOUND:  errorDesc = "Device not found"; break;
    case MLAPI_NOT_SUPPORTED:     errorDesc = "Feature not supported"; break;
    default:                      errorDesc = "Unknown SDK error code: " + std::to_string(errorCode); break;
    }

    char buffer[LOG_BUFFER_SIZE];
    snprintf(buffer, LOG_BUFFER_SIZE, "[MysticFight] SDK Error in %s: %s", functionName, errorDesc.c_str());
    Log(buffer);
}

/**
 * Prevents log file from growing indefinitely
 */
static void TrimLogFile() {
    WIN32_FILE_ATTRIBUTE_DATA fad;
    if (!GetFileAttributesExW(LOG_FILENAME, GetFileExInfoStandard, &fad)) return;

    LARGE_INTEGER size;
    size.LowPart = fad.nFileSizeLow;
    size.HighPart = fad.nFileSizeHigh;
    if (size.QuadPart < 1024 * 1024) return; // 1MB Limit

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

/**
 * Executes a process without administrative inheritance
 */
static void RunShellNonAdmin(const wchar_t* path) {
    ShellExecuteW(NULL, L"open", L"explorer.exe", path, NULL, SW_SHOWNORMAL);
}

/**
 * Requests SeDebugPrivilege to allow process termination
 */
static bool EnableDebugPrivilege() {
    HANDLE hToken;
    LUID luid;
    TOKEN_PRIVILEGES tp;

    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken)) return false;
    if (!LookupPrivilegeValue(NULL, SE_DEBUG_NAME, &luid)) { CloseHandle(hToken); return false; }

    tp.PrivilegeCount = 1;
    tp.Privileges[0].Luid = luid;
    tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;

    bool result = AdjustTokenPrivileges(hToken, FALSE, &tp, sizeof(TOKEN_PRIVILEGES), NULL, NULL);
    CloseHandle(hToken);
    return result && (GetLastError() == ERROR_SUCCESS);
}

/**
 * Modern Windows Notification (Toast style via Tray)
 */
static void ShowNotification(HWND hWnd, const wchar_t* title, const wchar_t* info, DWORD iconType = NIIF_USER) {
    NOTIFYICONDATAW nid = { sizeof(NOTIFYICONDATAW), hWnd, 1, NIF_INFO };
    wcsncpy_s(nid.szInfoTitle, title, _TRUNCATE);
    wcsncpy_s(nid.szInfo, info, _TRUNCATE);
    nid.dwInfoFlags = iconType | NIIF_LARGE_ICON;
    Shell_NotifyIconW(NIM_MODIFY, &nid);
}

/**
 * Terminates an external process by filename
 */
static void KillProcessByName(const wchar_t* filename) {
    EnableDebugPrivilege();

    HANDLE hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnap == INVALID_HANDLE_VALUE) return;

    PROCESSENTRY32W pe;
    pe.dwSize = sizeof(PROCESSENTRY32W);

    if (Process32FirstW(hSnap, &pe)) {
        do {
            if (_wcsicmp(pe.szExeFile, filename) == 0) {
                HANDLE hProc = OpenProcess(PROCESS_TERMINATE, FALSE, pe.th32ProcessID);
                if (hProc) {
                    if (TerminateProcess(hProc, 1)) {
                        char buf[128];
                        snprintf(buf, sizeof(buf), "[Watchdog] Successfully terminated: %ls", filename);
                        Log(buf);
                    }
                    else {
                        Log("[Watchdog] Found process but TerminateProcess failed.");
                    }
                    CloseHandle(hProc);
                }
                else {
                    DWORD err = GetLastError();
                    char buf[128];
                    snprintf(buf, sizeof(buf), "[Watchdog] OpenProcess failed for %ls (Error: %lu)", filename, err);
                    Log(buf);
                }
            }
        } while (Process32NextW(hSnap, &pe));
    }
    CloseHandle(hSnap);
}

/**
 * String duplication on Process Heap
 */
static wchar_t* HeapDupString(const wchar_t* src) {
    if (!src) return nullptr;
    size_t size = (wcslen(src) + 1) * sizeof(wchar_t);
    wchar_t* dest = (wchar_t*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, size);
    if (dest) wcscpy_s(dest, wcslen(src) + 1, src);
    return dest;
}

/**
 * Clears ComboBox items and frees associated heap data
 */
static void ClearComboHeapData(HWND hCombo) {
    int count = (int)SendMessage(hCombo, CB_GETCOUNT, 0, 0);
    for (int i = 0; i < count; i++) {
        wchar_t* ptr = (wchar_t*)SendMessage(hCombo, CB_GETITEMDATA, i, 0);
        if (ptr && ptr != (wchar_t*)CB_ERR) {
            HeapFree(GetProcessHeap(), 0, ptr);
        }
    }
    SendMessage(hCombo, CB_RESETCONTENT, 0, 0);
}


// ============================================================================
// COLOR CONVERSION HELPERS
// ============================================================================

static bool IsValidHex(const wchar_t* hex) {
    if (!hex || wcslen(hex) != HEX_COLOR_LEN || hex[0] != L'#') return false;
    for (int i = 1; i < HEX_COLOR_LEN; i++) {
        if (!iswxdigit(hex[i])) return false;
    }
    return true;
}

static COLORREF HexToColor(const wchar_t* hex) {
    if (hex[0] == L'#') hex++;
    unsigned int r = 0, g = 0, b = 0;
    if (swscanf_s(hex, L"%02x%02x%02x", &r, &g, &b) == 3) {
        return RGB(r, g, b);
    }
    return RGB(0, 0, 0);
}

static void ColorToHex(COLORREF color, wchar_t* out, size_t size) {
    swprintf_s(out, size, L"#%02X%02X%02X", GetRValue(color), GetGValue(color), GetBValue(color));
}

static bool IsColorDark(COLORREF col) {
    double brightness = (GetRValue(col) * 299 + GetGValue(col) * 587 + GetBValue(col) * 114) / 1000.0;
    return brightness < 128.0;
}

static int GetIntFromSafeArray(void* pRawData, VARTYPE vt, int index) {
    if (!pRawData) return 0;
    switch (vt) {
    case VT_BSTR: return _wtoi(((BSTR*)pRawData)[index]);
    case VT_I4: case VT_INT: return (int)((long*)pRawData)[index];
    case VT_UI4: case VT_UINT: return (int)((unsigned int*)pRawData)[index];
    case VT_I2: return (int)((short*)pRawData)[index];
    default: return 0;
    }
}


// ============================================================================
// HOTKEY MANAGEMENT
// ============================================================================

static UINT GetModifiersFromHotkey(WORD hotkey) {
    BYTE modifiers = HIBYTE(hotkey);
    UINT fsModifiers = 0;
    if (modifiers & HOTKEYF_SHIFT) fsModifiers |= MOD_SHIFT;
    if (modifiers & HOTKEYF_CONTROL) fsModifiers |= MOD_CONTROL;
    if (modifiers & HOTKEYF_ALT) fsModifiers |= MOD_ALT;
    return fsModifiers;
}

static WORD GetHotkeyForUI(UINT mod, UINT vk) {
    BYTE uiMod = 0;
    if (mod & MOD_SHIFT) uiMod |= HOTKEYF_SHIFT;
    if (mod & MOD_CONTROL) uiMod |= HOTKEYF_CONTROL;
    if (mod & MOD_ALT) uiMod |= HOTKEYF_ALT;
    return MAKEWORD(vk, uiMod);
}

static void SafeRegisterHotkey(HWND hWnd, int id, WORD hotkeyConfig) {
    if (!hWnd) return;
    UINT vk = LOBYTE(hotkeyConfig);
    if (vk == 0) return;
    UINT mod = GetModifiersFromHotkey(hotkeyConfig);
    RegisterHotKey(hWnd, id, mod | MOD_NOREPEAT, vk);
}

static void RegisterAppHotkeys(HWND hWnd) {
    if (!hWnd) return;
    UnregisterHotKey(hWnd, 1);
    for (int i = 101; i <= 105; i++) UnregisterHotKey(hWnd, i);

    SafeRegisterHotkey(hWnd, 1, g_Global.hotkeys.toggleLEDs);
    SafeRegisterHotkey(hWnd, 101, g_Global.hotkeys.profile1);
    SafeRegisterHotkey(hWnd, 102, g_Global.hotkeys.profile2);
    SafeRegisterHotkey(hWnd, 103, g_Global.hotkeys.profile3);
    SafeRegisterHotkey(hWnd, 104, g_Global.hotkeys.profile4);
    SafeRegisterHotkey(hWnd, 105, g_Global.hotkeys.profile5);
}


// ============================================================================
// NETWORK & JSON PARSING (LibreHardwareMonitor Integration)
// ============================================================================

/**
 * Thread-safe HTTP fetcher for LHM JSON data
 */
static std::string FetchLHMJson(const wchar_t* serverUrl, int timeout) {
    std::lock_guard<std::mutex> lock(g_httpMutex);

    if (!g_Running) return "";

    if (g_ResetHttp) {
        if (g_hConnect) { WinHttpCloseHandle(g_hConnect); g_hConnect = NULL; }
        if (g_hSession) { WinHttpCloseHandle(g_hSession); g_hSession = NULL; }
        g_ResetHttp = false;
    }

    std::string responseData;
    HINTERNET hRequest = NULL;
    bool connectionFailed = false;

    if (!g_hSession) {
        g_hSession = WinHttpOpen(L"MysticFight (Persistent)", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (!g_hSession) return "";
    }
    WinHttpSetTimeouts(g_hSession, timeout, timeout, timeout, timeout);

    if (!g_hConnect) {
        URL_COMPONENTS urlComp = { sizeof(URL_COMPONENTS) };
        urlComp.dwHostNameLength = (DWORD)-1;
        urlComp.dwUrlPathLength = (DWORD)-1;

        if (WinHttpCrackUrl(serverUrl, (DWORD)wcslen(serverUrl), 0, &urlComp)) {
            std::wstring host(urlComp.lpszHostName, urlComp.dwHostNameLength);
            g_hConnect = WinHttpConnect(g_hSession, host.c_str(), urlComp.nPort, 0);
        }
        if (!g_hConnect) return "";
    }

    hRequest = WinHttpOpenRequest(g_hConnect, L"GET", L"/data.json", NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, 0);
    if (!hRequest) {
        WinHttpCloseHandle(g_hConnect); g_hConnect = NULL;
        return "";
    }

    if (WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
        WinHttpReceiveResponse(hRequest, NULL)) {

        DWORD dwSize = 0, dwDownloaded = 0;
        do {
            dwSize = 0;
            if (!WinHttpQueryDataAvailable(hRequest, &dwSize)) { connectionFailed = true; break; }
            if (dwSize == 0) break;

            if (responseData.size() + dwSize > MAX_HTTP_RESPONSE_SIZE) {
                Log("[MysticFight] HTTP Response exceeded size limit. Aborting.");
                connectionFailed = true;
                break;
            }

            char chunk[4096];
            DWORD toRead = (dwSize > sizeof(chunk)) ? sizeof(chunk) : dwSize;
            if (WinHttpReadData(hRequest, (LPVOID)chunk, toRead, &dwDownloaded)) {
                responseData.append(chunk, dwDownloaded);
            }
            else {
                connectionFailed = true; break;
            }
        } while (dwSize > 0);
    }
    else {
        connectionFailed = true;
    }

    WinHttpCloseHandle(hRequest);

    if (connectionFailed) {
        Log("[MysticFight] HTTP Connection dropped. Resetting handle...");
        if (g_hConnect) { WinHttpCloseHandle(g_hConnect); g_hConnect = NULL; }
        return "";
    }

    return responseData;
}

/**
 * Modern JSON Logic: Clean LHM Value Strings
 */
static float ParseLHMValue(const json& jVal) {
    try {
        if (jVal.is_number()) return jVal.get<float>();

        if (jVal.is_string()) {
            std::string valStr = jVal.get<std::string>();
            std::replace(valStr.begin(), valStr.end(), ',', '.'); // Coma a punto

            size_t spacePos = valStr.find(' ');
            if (spacePos != std::string::npos) valStr = valStr.substr(0, spacePos);

            return std::stof(valStr);
        }
    }
    catch (...) { return -1.0f; }
    return -1.0f;
}

/**
 * Modern JSON Logic: Recursive Search
 */
static float FindSensorValueRecursive(const json& jNode, const std::string& targetId) {
    if (jNode.is_object()) {
        // CORRECCIÓN: Usamos "SensorId" que es único y estable
        if (jNode.contains("SensorId")) {
            std::string currentSensorId = jNode["SensorId"].get<std::string>();

            if (currentSensorId == targetId) {
                if (jNode.contains("Value") && !jNode["Value"].is_null()) {
                    return ParseLHMValue(jNode["Value"]);
                }
            }
        }

        // Si no es el buscado, mirar en hijos
        if (jNode.contains("Children") && jNode["Children"].is_array()) {
            for (const auto& child : jNode["Children"]) {
                float val = FindSensorValueRecursive(child, targetId);
                if (val != -1.0f) return val;
            }
        }
    }
    else if (jNode.is_array()) {
        for (const auto& element : jNode) {
            float val = FindSensorValueRecursive(element, targetId);
            if (val != -1.0f) return val;
        }
    }
    return -1.0f;
}



/**
 * Main Sensor Parser
 */
static float ParseLHMJsonForTemp(const std::string& jsonStr, const wchar_t* sensorIDW) {
    if (jsonStr.empty() || !sensorIDW) return -1.0f;

    // Convert Wchar ID to Standard String
    char sensorID[256];
    size_t converted;
    wcstombs_s(&converted, sensorID, sizeof(sensorID), sensorIDW, _TRUNCATE);

    try {
        // 1. Parse JSON string to DOM object
        auto j = json::parse(jsonStr);

        // 2. Search value navigating the tree
        return FindSensorValueRecursive(j, std::string(sensorID));

    }
    catch (const json::parse_error& e) {
        // Log error if JSON is corrupt
        char buf[256];
        snprintf(buf, sizeof(buf), "[JSON Error] %s", e.what());
        Log(buf);
        return -1.0f;
    }
}


// ============================================================================
// CONFIGURATION PERSISTENCE (INI)
// ============================================================================

static void SaveSettings() {
    std::lock_guard<std::mutex> lock(g_cfgMutex);

    WritePrivateProfileStringW(L"Global", L"ActiveProfile", std::to_wstring(g_Global.activeProfileIndex).c_str(), INI_FILE);

    for (int i = 0; i < 5; i++) {
        std::wstring section = L"Settings";
        if (i > 0) section += L"_" + std::to_wstring(i);

        Config& p = g_Global.profiles[i];

        WritePrivateProfileStringW(section.c_str(), L"Label", p.label, INI_FILE);
        WritePrivateProfileStringW(section.c_str(), L"SensorUpdateMS", std::to_wstring(p.sensorUpdateMS).c_str(), INI_FILE);
        WritePrivateProfileStringW(section.c_str(), L"LedRefreshFPS", std::to_wstring(p.ledRefreshFPS).c_str(), INI_FILE);

        wchar_t sfBuf[32]; swprintf_s(sfBuf, L"%.3f", p.smoothingFactor);
        WritePrivateProfileStringW(section.c_str(), L"SmoothingFactor", sfBuf, INI_FILE);

        WritePrivateProfileStringW(section.c_str(), L"SensorID", p.sensorID, INI_FILE);
        WritePrivateProfileStringW(section.c_str(), L"TargetDevice", p.targetDevice, INI_FILE);
        WritePrivateProfileStringW(section.c_str(), L"TargetLedIndex", std::to_wstring(p.targetLedIndex).c_str(), INI_FILE);

        WritePrivateProfileStringW(section.c_str(), L"TempLow", std::to_wstring(p.tempLow).c_str(), INI_FILE);
        WritePrivateProfileStringW(section.c_str(), L"TempMed", std::to_wstring(p.tempMed).c_str(), INI_FILE);
        WritePrivateProfileStringW(section.c_str(), L"TempHigh", std::to_wstring(p.tempHigh).c_str(), INI_FILE);

        wchar_t hL[10], hM[10], hH[10];
        ColorToHex(p.colorLow, hL, 10);
        ColorToHex(p.colorMed, hM, 10);
        ColorToHex(p.colorHigh, hH, 10);

        WritePrivateProfileStringW(section.c_str(), L"ColorLow", hL, INI_FILE);
        WritePrivateProfileStringW(section.c_str(), L"ColorMed", hM, INI_FILE);
        WritePrivateProfileStringW(section.c_str(), L"ColorHigh", hH, INI_FILE);

        WritePrivateProfileStringW(section.c_str(), L"WebServerUrl", p.webServerUrl, INI_FILE);
    }

    WritePrivateProfileStringW(L"Hotkeys", L"Toggle", std::to_wstring(g_Global.hotkeys.toggleLEDs).c_str(), INI_FILE);
    WritePrivateProfileStringW(L"Hotkeys", L"Profile1", std::to_wstring(g_Global.hotkeys.profile1).c_str(), INI_FILE);
    WritePrivateProfileStringW(L"Hotkeys", L"Profile2", std::to_wstring(g_Global.hotkeys.profile2).c_str(), INI_FILE);
    WritePrivateProfileStringW(L"Hotkeys", L"Profile3", std::to_wstring(g_Global.hotkeys.profile3).c_str(), INI_FILE);
    WritePrivateProfileStringW(L"Hotkeys", L"Profile4", std::to_wstring(g_Global.hotkeys.profile4).c_str(), INI_FILE);
    WritePrivateProfileStringW(L"Hotkeys", L"Profile5", std::to_wstring(g_Global.hotkeys.profile5).c_str(), INI_FILE);
}

/**
 * Attempts to find the first temperature sensor if none is configured
 */
static bool AutoSelectFirstSensor() {
    wchar_t localURL[256];
    {
        std::lock_guard<std::mutex> lock(g_cfgMutex);
        memcpy(localURL, g_cfg.webServerUrl, sizeof(g_cfg.webServerUrl));
    }

    std::string jsonStr = FetchLHMJson(localURL, HTTP_FAST_TIMEOUT);
    if (jsonStr.empty()) return false;

    // Simplified Auto-Select logic using string search for fallback compatibility
    std::string typeKey = "\"Type\":\"Temperature\"";
    size_t pos = jsonStr.find(typeKey);

    if (pos != std::string::npos) {
        size_t blockStart = jsonStr.rfind('{', pos);
        size_t blockEnd = jsonStr.find('}', pos);

        if (blockStart != std::string::npos && blockEnd != std::string::npos) {
            std::string block = jsonStr.substr(blockStart, blockEnd - blockStart + 1);

            // Re-implement basic string extraction just for this fallback
            std::string idKey = "\"SensorId\":\"";
            size_t idStart = block.find(idKey);
            if (idStart != std::string::npos) {
                idStart += idKey.length();
                size_t idEnd = block.find("\"", idStart);
                if (idEnd != std::string::npos) {
                    std::string idVal = block.substr(idStart, idEnd - idStart);

                    int len = MultiByteToWideChar(CP_UTF8, 0, idVal.c_str(), -1, NULL, 0);
                    std::vector<wchar_t> wId(len);
                    MultiByteToWideChar(CP_UTF8, 0, idVal.c_str(), -1, &wId[0], len);

                    std::lock_guard<std::mutex> lock(g_cfgMutex);
                    for (int i = 0; i < 5; i++) {
                        if (wcslen(g_Global.profiles[i].sensorID) == 0) {
                            wcscpy_s(g_Global.profiles[i].sensorID, &wId[0]);
                        }
                    }
                    Log("[MysticFight] Auto-assigned sensor to empty profiles.");
                    return true;
                }
            }
        }
    }
    return false;
}

static void LoadSettings() {
    g_Global.activeProfileIndex = GetPrivateProfileIntW(L"Global", L"ActiveProfile", 0, INI_FILE);
    if (g_Global.activeProfileIndex < 0 || g_Global.activeProfileIndex > 4) g_Global.activeProfileIndex = 0;

    for (int i = 0; i < 5; i++) {
        std::wstring section = L"Settings";
        if (i > 0) section += L"_" + std::to_wstring(i);
        Config& p = g_Global.profiles[i];

        p.tempLow = GetPrivateProfileIntW(section.c_str(), L"TempLow", DEF_TEMP_LOW, INI_FILE);
        p.tempMed = GetPrivateProfileIntW(section.c_str(), L"TempMed", DEF_TEMP_MED, INI_FILE);
        p.tempHigh = GetPrivateProfileIntW(section.c_str(), L"TempHigh", DEF_TEMP_HIGH, INI_FILE);

        GetPrivateProfileStringW(section.c_str(), L"TargetDevice", DEF_DEVICE_TYPE, p.targetDevice, 256, INI_FILE);
        p.targetLedIndex = GetPrivateProfileIntW(section.c_str(), L"TargetLedIndex", 0, INI_FILE);

        wchar_t hL[10], hM[10], hH[10];
        GetPrivateProfileStringW(section.c_str(), L"ColorLow", DEF_COLOR_LOW, hL, 10, INI_FILE);
        GetPrivateProfileStringW(section.c_str(), L"ColorMed", DEF_COLOR_MED, hM, 10, INI_FILE);
        GetPrivateProfileStringW(section.c_str(), L"ColorHigh", DEF_COLOR_HIGH, hH, 10, INI_FILE);
        p.colorLow = HexToColor(hL); p.colorMed = HexToColor(hM); p.colorHigh = HexToColor(hH);

        GetPrivateProfileStringW(section.c_str(), L"SensorID", L"", p.sensorID, SENSOR_ID_LEN, INI_FILE);
        GetPrivateProfileStringW(section.c_str(), L"WebServerUrl", DEF_SERVER_URL, p.webServerUrl, 256, INI_FILE);

        wchar_t defLabel[64]; swprintf_s(defLabel, L"Profile %d", i + 1);
        GetPrivateProfileStringW(section.c_str(), L"Label", defLabel, p.label, 64, INI_FILE);

        p.sensorUpdateMS = GetPrivateProfileIntW(section.c_str(), L"SensorUpdateMS", DEF_SENSOR_UPDATE_MS, INI_FILE);
        p.ledRefreshFPS = GetPrivateProfileIntW(section.c_str(), L"LedRefreshFPS", DEF_LED_REFRESH_FPS, INI_FILE);

        wchar_t sfBuf[32], sfDef[32]; swprintf_s(sfDef, L"%.3f", DEF_SMOOTHING_FACTOR);
        GetPrivateProfileStringW(section.c_str(), L"SmoothingFactor", sfDef, sfBuf, 32, INI_FILE);
        p.smoothingFactor = (float)_wtof(sfBuf);
    }

    if (wcslen(g_cfg.sensorID) == 0) { AutoSelectFirstSensor(); SaveSettings(); }

    // Hotkey default values (Ctrl+Alt+Shift + Keys)
    BYTE defMod = HOTKEYF_CONTROL | HOTKEYF_SHIFT | HOTKEYF_ALT;
    g_Global.hotkeys.toggleLEDs = (WORD)GetPrivateProfileIntW(L"Hotkeys", L"Toggle", MAKEWORD(0x4C, defMod), INI_FILE);
    g_Global.hotkeys.profile1 = (WORD)GetPrivateProfileIntW(L"Hotkeys", L"Profile1", MAKEWORD('1', defMod), INI_FILE);
    g_Global.hotkeys.profile2 = (WORD)GetPrivateProfileIntW(L"Hotkeys", L"Profile2", MAKEWORD('2', defMod), INI_FILE);
    g_Global.hotkeys.profile3 = (WORD)GetPrivateProfileIntW(L"Hotkeys", L"Profile3", MAKEWORD('3', defMod), INI_FILE);
    g_Global.hotkeys.profile4 = (WORD)GetPrivateProfileIntW(L"Hotkeys", L"Profile4", MAKEWORD('4', defMod), INI_FILE);
    g_Global.hotkeys.profile5 = (WORD)GetPrivateProfileIntW(L"Hotkeys", L"Profile5", MAKEWORD('5', defMod), INI_FILE);
}


// ============================================================================
// TASK SCHEDULER HELPERS (Startup Management)
// ============================================================================

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

static bool ValidStartupTaskExists() {
    wchar_t szCurrentPath[MAX_PATH];
    if (GetModuleFileNameW(NULL, szCurrentPath, MAX_PATH) == 0) return false;

    try {
        ITaskServicePtr pService;
        if (FAILED(pService.CreateInstance(__uuidof(TaskScheduler), NULL, CLSCTX_INPROC_SERVER))) return false;

        pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
        ITaskFolderPtr pRootFolder;
        pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);

        IRegisteredTaskPtr pRegisteredTask;
        if (FAILED(pRootFolder->GetTask(_bstr_t(TASK_NAME), &pRegisteredTask))) return false;

        ITaskDefinitionPtr pDefinition;
        pRegisteredTask->get_Definition(&pDefinition);
        IActionCollectionPtr pActions;
        pDefinition->get_Actions(&pActions);

        IActionPtr pAction;
        if (FAILED(pActions->get_Item(1, &pAction))) return false;

        IExecActionPtr pExecAction = pAction;
        if (pExecAction == nullptr) return false;

        BSTR bstrPath = NULL;
        if (SUCCEEDED(pExecAction->get_Path(&bstrPath))) {
            _bstr_t pathWrapper(bstrPath, false);
            return (_wcsicmp(szCurrentPath, (const wchar_t*)pathWrapper) == 0);
        }
    }
    catch (...) {
        return false;
    }
    return false;
}

/**
 * Creates or removes a Task Scheduler entry with Highest Privileges
 */
static void SetStartupTask(bool run) {
    wchar_t szPath[MAX_PATH], szDir[MAX_PATH];
    if (GetModuleFileNameW(NULL, szPath, MAX_PATH) == 0) return;
    if (run && ValidStartupTaskExists()) return;

    wcscpy_s(szDir, szPath);
    wchar_t* lastSlash = wcsrchr(szDir, L'\\');
    if (lastSlash) *lastSlash = L'\0';

    try {
        ITaskServicePtr pService;
        if (FAILED(pService.CreateInstance(__uuidof(TaskScheduler), NULL, CLSCTX_INPROC_SERVER))) return;

        pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
        ITaskFolderPtr pRootFolder;
        pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);

        pRootFolder->DeleteTask(_bstr_t(TASK_NAME), 0);
        if (!run) return;

        ITaskDefinitionPtr pTask;
        pService->NewTask(0, &pTask);

        IPrincipalPtr pPrincipal;
        if (SUCCEEDED(pTask->get_Principal(&pPrincipal)) && pPrincipal != nullptr) {
            pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);
            pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
        }

        ITaskSettingsPtr pSettings;
        if (SUCCEEDED(pTask->get_Settings(&pSettings)) && pSettings != nullptr) {
            pSettings->put_StartWhenAvailable(VARIANT_TRUE);
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
            _bstr_t(TASK_NAME), pTask, TASK_CREATE_OR_UPDATE, _variant_t(),
            _variant_t(), TASK_LOGON_INTERACTIVE_TOKEN, _variant_t(L""),
            &pRegisteredTask);

        Log("[MysticFight] Startup task created.");
    }
    catch (...) {
        Log("[MysticFight] Failed to create startup task.");
    }
}


// ============================================================================
// UI DATA POPULATION HELPERS (Modernized)
// ============================================================================

// Recursive JSON Collector for UI
static void CollectSensorsRecursive(const json& jNode, std::vector<std::pair<std::string, std::string>>& sensors) {
    if (jNode.is_object()) {
        bool isTemp = false;

        // Detectar si es temperatura por icono o tipo
        if (jNode.contains("ImageURL") && jNode["ImageURL"].get<std::string>() == "images_icon/temperature.png") isTemp = true;
        if (jNode.contains("Type") && jNode["Type"] == "Temperature") isTemp = true;

        // Si es temperatura, guardamos Nombre y SensorId (Ruta)
        if (isTemp && jNode.contains("Text") && jNode.contains("SensorId")) {
            std::string name = jNode["Text"].get<std::string>();
            std::string path = jNode["SensorId"].get<std::string>();
            sensors.push_back({ name, path });
        }

        // Recursión hijos
        if (jNode.contains("Children") && jNode["Children"].is_array()) {
            for (const auto& child : jNode["Children"]) CollectSensorsRecursive(child, sensors);
        }
    }
    else if (jNode.is_array()) {
        for (const auto& element : jNode) CollectSensorsRecursive(element, sensors);
    }
}

static void PopulateSensorList(HWND hDlg, const wchar_t* webServerUrl, const wchar_t* currentTargetID) {
    HWND hCombo = GetDlgItem(hDlg, IDC_SENSOR_ID);
    ClearComboHeapData(hCombo);

    std::string jsonStr = FetchLHMJson(webServerUrl, (int)HTTP_FAST_TIMEOUT);
    if (jsonStr.empty()) {
        SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)L"LHM Not Found (Port 8085)");
        EnableWindow(hCombo, FALSE);
        return;
    }

    try {
        auto j = json::parse(jsonStr);
        std::vector<std::pair<std::string, std::string>> sensors;
        CollectSensorsRecursive(j, sensors);

        bool idSelected = false;

        for (const auto& sensor : sensors) {
            // Convert strings to wide chars for Windows Controls
            int lenName = MultiByteToWideChar(CP_UTF8, 0, sensor.first.c_str(), -1, NULL, 0);
            int lenId = MultiByteToWideChar(CP_UTF8, 0, sensor.second.c_str(), -1, NULL, 0);

            std::vector<wchar_t> wName(lenName), wId(lenId);
            MultiByteToWideChar(CP_UTF8, 0, sensor.first.c_str(), -1, &wName[0], lenName);
            MultiByteToWideChar(CP_UTF8, 0, sensor.second.c_str(), -1, &wId[0], lenId);

            wchar_t* pSafeId = HeapDupString(&wId[0]);
            int idx = (int)SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)&wName[0]);
            if (idx != CB_ERR) {
                SendMessage(hCombo, CB_SETITEMDATA, idx, (LPARAM)pSafeId);
                if (currentTargetID && wcscmp(&wId[0], currentTargetID) == 0) {
                    SendMessage(hCombo, CB_SETCURSEL, idx, 0);
                    idSelected = true;
                }
            }
            else {
                HeapFree(GetProcessHeap(), 0, pSafeId);
            }
        }

        if (!idSelected && SendMessage(hCombo, CB_GETCOUNT, 0, 0) > 0) {
            SendMessage(hCombo, CB_SETCURSEL, 0, 0);
        }

        EnableWindow(hCombo, SendMessage(hCombo, CB_GETCOUNT, 0, 0) > 1);
    }
    catch (const std::exception& e) {
        Log("[UI Error] Error parsing JSON for list");
    }
}

// Forward declarations for UI synchronization
static void PopulateAreaList(HWND hDlg, const wchar_t* deviceType, int targetLedIndex);

/**
 * Transfers Profile Data into the Settings Dialog controls
 */
static void LoadProfileToUI(HWND hDlg, int profileIndex) {
    if (profileIndex < 0 || profileIndex >= 5) return;
    Config& p = g_Global.profiles[profileIndex];

    SetDlgItemInt(hDlg, IDC_TEMP_LOW, p.tempLow, TRUE);
    SetDlgItemInt(hDlg, IDC_TEMP_MED, p.tempMed, TRUE);
    SetDlgItemInt(hDlg, IDC_TEMP_HIGH, p.tempHigh, TRUE);
    SetDlgItemTextW(hDlg, IDC_EDIT_LABEL, p.label);
    SetDlgItemTextW(hDlg, IDC_EDIT_SERVER, p.webServerUrl);

    wchar_t hexBuf[10];
    ColorToHex(p.colorLow, hexBuf, 10);  SetDlgItemTextW(hDlg, IDC_HEX_LOW, hexBuf);
    ColorToHex(p.colorMed, hexBuf, 10);  SetDlgItemTextW(hDlg, IDC_HEX_MED, hexBuf);
    ColorToHex(p.colorHigh, hexBuf, 10); SetDlgItemTextW(hDlg, IDC_HEX_HIGH, hexBuf);

    HWND hComboDev = GetDlgItem(hDlg, IDC_COMBO_DEVICE);
    int devCount = (int)SendMessage(hComboDev, CB_GETCOUNT, 0, 0);
    for (int i = 0; i < devCount; i++) {
        wchar_t* devName = (wchar_t*)SendMessage(hComboDev, CB_GETITEMDATA, i, 0);
        if (devName && wcscmp(devName, p.targetDevice) == 0) {
            SendMessage(hComboDev, CB_SETCURSEL, i, 0);
            break;
        }
    }

    PopulateAreaList(hDlg, p.targetDevice, p.targetLedIndex);
    PopulateSensorList(hDlg, p.webServerUrl, p.sensorID);

    auto SetComboByVal = [&](int id, int val, std::vector<int> mapping) {
        HWND hC = GetDlgItem(hDlg, id);
        for (size_t i = 0; i < mapping.size(); ++i) {
            if (mapping[i] == val) { SendMessage(hC, CB_SETCURSEL, (WPARAM)i, 0); return; }
        }
        SendMessage(hC, CB_SETCURSEL, 0, 0);
        };

    SetComboByVal(IDC_COMBO_SENSOR_UPDATE, p.sensorUpdateMS, { 250, 500, 1000 });
    SetComboByVal(IDC_COMBO_LED_FPS, p.ledRefreshFPS, { 25, 20, 15 });

    HWND hSmooth = GetDlgItem(hDlg, IDC_COMBO_SMOOTHING);
    if (p.smoothingFactor > 0.9f) SendMessage(hSmooth, CB_SETCURSEL, 0, 0);
    else if (p.smoothingFactor < 0.1f) SendMessage(hSmooth, CB_SETCURSEL, 1, 0);
    else if (p.smoothingFactor > 0.2f) SendMessage(hSmooth, CB_SETCURSEL, 3, 0);
    else SendMessage(hSmooth, CB_SETCURSEL, 2, 0);

    CheckDlgButton(hDlg, IDC_CHK_ACTIVE_PROFILE, (profileIndex == g_Global.activeProfileIndex) ? BST_CHECKED : BST_UNCHECKED);
    InvalidateRect(hDlg, NULL, TRUE);
}

/**
 * Transfers Dialog data back to global config profiles
 */
static void SaveUIToProfile(HWND hDlg, int profileIndex) {
    Config& p = g_Global.profiles[profileIndex];

    p.tempLow = GetDlgItemInt(hDlg, IDC_TEMP_LOW, NULL, TRUE);
    p.tempMed = GetDlgItemInt(hDlg, IDC_TEMP_MED, NULL, TRUE);
    p.tempHigh = GetDlgItemInt(hDlg, IDC_TEMP_HIGH, NULL, TRUE);

    wchar_t buf[256];
    GetDlgItemTextW(hDlg, IDC_HEX_LOW, buf, 10);
    p.colorLow = HexToColor(buf);
    GetDlgItemTextW(hDlg, IDC_HEX_MED, buf, 10);
    p.colorMed = HexToColor(buf);
    GetDlgItemTextW(hDlg, IDC_HEX_HIGH, buf, 10);
    p.colorHigh = HexToColor(buf);

    GetDlgItemTextW(hDlg, IDC_EDIT_SERVER, p.webServerUrl, 256);

    int idxDev = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_DEVICE), CB_GETCURSEL, 0, 0);
    if (idxDev != CB_ERR) {
        wchar_t* devName = (wchar_t*)SendMessage(GetDlgItem(hDlg, IDC_COMBO_DEVICE), CB_GETITEMDATA, idxDev, 0);
        if (devName) wcscpy_s(p.targetDevice, devName);
    }

    int idxArea = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_AREA), CB_GETCURSEL, 0, 0);
    if (idxArea != CB_ERR) {
        p.targetLedIndex = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_AREA), CB_GETITEMDATA, idxArea, 0);
    }

    int idxSens = (int)SendMessage(GetDlgItem(hDlg, IDC_SENSOR_ID), CB_GETCURSEL, 0, 0);
    if (idxSens != CB_ERR) {
        wchar_t* sID = (wchar_t*)SendMessage(GetDlgItem(hDlg, IDC_SENSOR_ID), CB_GETITEMDATA, idxSens, 0);
        if (sID) wcscpy_s(p.sensorID, sID);
    }

    GetDlgItemTextW(hDlg, IDC_EDIT_LABEL, p.label, 64);

    int idxSensor = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_SENSOR_UPDATE), CB_GETCURSEL, 0, 0);
    if (idxSensor == 0) p.sensorUpdateMS = 250;
    else if (idxSensor == 2) p.sensorUpdateMS = 1000;
    else p.sensorUpdateMS = 500;

    int idxFPS = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_LED_FPS), CB_GETCURSEL, 0, 0);
    if (idxFPS == 1) p.ledRefreshFPS = 20;
    else if (idxFPS == 2) p.ledRefreshFPS = 15;
    else p.ledRefreshFPS = 25;

    int idxSmooth = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_SMOOTHING), CB_GETCURSEL, 0, 0);
    if (idxSmooth == 0) p.smoothingFactor = 1.0f;
    else if (idxSmooth == 1) p.smoothingFactor = 0.05f;
    else if (idxSmooth == 3) p.smoothingFactor = 0.4f;
    else p.smoothingFactor = 0.15f;
}


// ============================================================================
// MSI SDK HARDWARE DETECTION & MANAGEMENT
// ============================================================================

static void PopulateAreaList(HWND hDlg, const wchar_t* deviceType, int targetLedIndex) {
    HWND hComboArea = GetDlgItem(hDlg, IDC_COMBO_AREA);
    SendMessage(hComboArea, CB_RESETCONTENT, 0, 0);

    if (!deviceType || !lpMLAPI_GetLedInfo || !lpMLAPI_GetDeviceInfo) {
        EnableWindow(hComboArea, FALSE);
        return;
    }

    SAFEARRAY* pDevType = nullptr;
    SAFEARRAY* pLedCount = nullptr;
    int currentDeviceLedCount = 0;

    int hrInfo = lpMLAPI_GetDeviceInfo(&pDevType, &pLedCount);
    if (hrInfo == MLAPI_OK && pDevType && pLedCount) {

        {
            SafeArrayScopedLock<BSTR> typeLock(pDevType);
            SafeArrayScopedLock<void> countLock(pLedCount);

            if (typeLock.isLocked() && countLock.isLocked()) {
                VARTYPE vtCount;
                SafeArrayGetVartype(pLedCount, &vtCount);

                long lBound, uBound;
                SafeArrayGetLBound(pDevType, 1, &lBound);
                SafeArrayGetUBound(pDevType, 1, &uBound);
                long totalDevices = uBound - lBound + 1;

                for (long j = 0; j < totalDevices; j++) {
                    if (wcscmp(typeLock.data[j], deviceType) == 0) {
                        currentDeviceLedCount = GetIntFromSafeArray(countLock.data, vtCount, (int)j);
                        break;
                    }
                }
            }
        } // Locks released here

        SafeArrayDestroy(pDevType);
        SafeArrayDestroy(pLedCount);
    }
    else if (hrInfo != MLAPI_OK) {
        LogSDKError("MLAPI_GetDeviceInfo (PopulateAreaList)", hrInfo);
    }

    if (currentDeviceLedCount > 0) {
        _bstr_t bstrDevice(deviceType);

        for (DWORD i = 0; i < (DWORD)currentDeviceLedCount; i++) {
            BSTR ledName = nullptr;
            SAFEARRAY* pStyles = nullptr;

            int hrLed = lpMLAPI_GetLedInfo(bstrDevice, i, &ledName, &pStyles);
            if (hrLed == MLAPI_OK) {
                int idx = (int)SendMessageW(hComboArea, CB_ADDSTRING, 0, (LPARAM)(ledName ? ledName : L"Unknown Area"));
                SendMessage(hComboArea, CB_SETITEMDATA, idx, (LPARAM)i);

                if (i == (DWORD)targetLedIndex) {
                    SendMessage(hComboArea, CB_SETCURSEL, idx, 0);
                }

                if (pStyles) SafeArrayDestroy(pStyles);
                if (ledName) SysFreeString(ledName);
            }
            else {
                LogSDKError("MLAPI_GetLedInfo", hrLed);
            }
        }
    }

    if (SendMessage(hComboArea, CB_GETCURSEL, 0, 0) == CB_ERR && SendMessage(hComboArea, CB_GETCOUNT, 0, 0) > 0) {
        SendMessage(hComboArea, CB_SETCURSEL, 0, 0);
    }

    int finalCount = (int)SendMessage(hComboArea, CB_GETCOUNT, 0, 0);
    EnableWindow(hComboArea, finalCount > 1);
}

static void PopulateDeviceList(HWND hDlg) {
    HWND hComboDev = GetDlgItem(hDlg, IDC_COMBO_DEVICE);
    ClearComboHeapData(hComboDev);

    SAFEARRAY* pDevType = nullptr, * pLedCount = nullptr;

    int hrInfo = MLAPI_ERROR;
    if (lpMLAPI_GetDeviceInfo) {
        hrInfo = lpMLAPI_GetDeviceInfo(&pDevType, &pLedCount);
    }

    if (hrInfo == MLAPI_OK) {
        {
            SafeArrayScopedLock<BSTR> typeLock(pDevType);

            if (typeLock.isLocked()) {
                long lBound, uBound;
                SafeArrayGetLBound(pDevType, 1, &lBound);
                SafeArrayGetUBound(pDevType, 1, &uBound);
                long count = uBound - lBound + 1;

                for (long i = 0; i < count; i++) {
                    _bstr_t bstrType(typeLock.data[i]);
                    BSTR friendlyBSTR = nullptr;

                    if (lpMLAPI_GetDeviceNameEx) {
                        int hrName = lpMLAPI_GetDeviceNameEx(bstrType, i, &friendlyBSTR);
                        if (hrName != MLAPI_OK) LogSDKError("MLAPI_GetDeviceNameEx", hrName);
                    }

                    _bstr_t bstrFriendly(friendlyBSTR, false);

                    wchar_t* pSafeStr = HeapDupString(bstrType);
                    int idx = (int)SendMessageW(hComboDev, CB_ADDSTRING, 0, (LPARAM)(wchar_t*)bstrFriendly);

                    if (idx != CB_ERR) {
                        SendMessage(hComboDev, CB_SETITEMDATA, idx, (LPARAM)pSafeStr);
                        if (wcscmp(bstrType, g_cfg.targetDevice) == 0)
                            SendMessage(hComboDev, CB_SETCURSEL, idx, 0);
                    }
                    else {
                        HeapFree(GetProcessHeap(), 0, pSafeStr);
                    }
                }
            }
        } // Locks released

        SafeArrayDestroy(pDevType);
        SafeArrayDestroy(pLedCount);

        int finalCount = (int)SendMessage(hComboDev, CB_GETCOUNT, 0, 0);
        EnableWindow(hComboDev, finalCount > 1);
    }
    else {
        if (hrInfo != MLAPI_ERROR) LogSDKError("MLAPI_GetDeviceInfo (PopulateDeviceList)", hrInfo);
        EnableWindow(hComboDev, FALSE);
    }
}


// ============================================================================
// CORE MONITORING LOGIC
// ============================================================================

/**
 * Debug utility to dump all available LHM sensors to the log file
 */
static void LogAllLHMTemperatureSensors() {
    Log("[MysticFight] --- DUMPING ALL LHM TEMPERATURE SENSORS ---");

    wchar_t localURL[256];
    {
        std::lock_guard<std::mutex> lock(g_cfgMutex);
        wcscpy_s(localURL, g_cfg.webServerUrl);
    }

    // 1. Descargamos el JSON crudo
    std::string jsonStr = FetchLHMJson(localURL, (int)HTTP_FAST_TIMEOUT);

    if (jsonStr.empty()) {
        Log("[MysticFight] Error: LHM returned EMPTY response.");
        return;
    }

    // 2. Intentamos parsear con control de errores detallado
    try {
        auto j = json::parse(jsonStr);
        std::vector<std::pair<std::string, std::string>> sensors;
        CollectSensorsRecursive(j, sensors);

        if (sensors.empty()) {
            Log("[MysticFight] JSON parsed, but NO temperature sensors found.");
            // Imprimimos el inicio del JSON para ver qué estructura tiene
            std::string debugSnippet = jsonStr.substr(0, 200);
            Log(("[JSON Snippet]: " + debugSnippet).c_str());
        }
        else {
            for (const auto& s : sensors) {
                // Try to fetch value for logging
                float val = FindSensorValueRecursive(j, s.second);

                int lenName = MultiByteToWideChar(CP_UTF8, 0, s.first.c_str(), -1, NULL, 0);
                int lenId = MultiByteToWideChar(CP_UTF8, 0, s.second.c_str(), -1, NULL, 0);
                std::vector<wchar_t> wName(lenName + 1), wId(lenId + 1); // +1 por seguridad
                MultiByteToWideChar(CP_UTF8, 0, s.first.c_str(), -1, &wName[0], lenName);
                MultiByteToWideChar(CP_UTF8, 0, s.second.c_str(), -1, &wId[0], lenId);

                wchar_t wBuf[LOG_BUFFER_SIZE];
                swprintf_s(wBuf, L"[HTTP] ID: %ls | Name: %ls | Value: %.1f", &wId[0], &wName[0], val);

                // Conversión segura para el Log (char*)
                _bstr_t logConverter(wBuf);
                Log(logConverter);
            }
        }
    }
    catch (const json::parse_error& e) {
        // ERROR DE SINTAXIS JSON
        char buf[512];
        snprintf(buf, sizeof(buf), "[JSON PARSE ERROR] Byte: %llu | Msg: %s", e.byte, e.what());
        Log(buf);

        // Muestra qué demonios recibió
        std::string debugSnippet = jsonStr.substr(0, 150);
        Log(("[RECEIVED DATA]: " + debugSnippet).c_str());
    }
    catch (const std::exception& e) {
        // OTRO ERROR (posiblemente de tipo)
        char buf[256];
        snprintf(buf, sizeof(buf), "[STD EXCEPTION] %s", e.what());
        Log(buf);
    }
    catch (...) {
        Log("[UNKNOWN CRITICAL ERROR] Something crashed inside the JSON parser.");
    }

    Log("[MysticFight] --- END OF HTTP DUMP ---");
}

/**
 * Primary sensor polling function with state management
 */
static float GetCPUTempFast() {
    ULONGLONG currentTime = GetTickCount64();

    if (g_activeSource == DataSource::Searching) {
        if (currentTime - g_lastDataSourceSearchRetry <= LHM_RETRY_DELAY_MS) return -1.0f;
        g_lastDataSourceSearchRetry = currentTime;
    }

    wchar_t localSensorID[SENSOR_ID_LEN];
    wchar_t localURL[256];
    {
        std::lock_guard<std::mutex> lock(g_cfgMutex);
        wcscpy_s(localSensorID, g_cfg.sensorID);
        wcscpy_s(localURL, g_cfg.webServerUrl);
    }

    std::string jsonStr = FetchLHMJson(localURL, HTTP_NORMAL_TIMEOUT);

    if (jsonStr.empty()) {
        if (g_activeSource == DataSource::HTTP) {
            g_httpConnectionDrops++;
            if (g_httpConnectionDrops >= 2) {
                Log("[MysticFight] Connection lost. Triggering SEARCH mode.");
                g_activeSource = DataSource::Searching;
                g_pendingStyleChange = true; // Trigger white fallback
                g_httpConnectionDrops = 0;
                SetEvent(g_hSensorEvent);
            }
        }
        return -1.0f;
    }

    if (g_activeSource == DataSource::Searching) {
        Log("[MysticFight] LHM HTTP Data Source detected");
        g_activeSource = DataSource::HTTP;
        g_pendingStyleChange = true;
        SetEvent(g_hSourceResolvedEvent);
        LogAllLHMTemperatureSensors();
    }

    g_httpConnectionDrops = 0;
    return ParseLHMJsonForTemp(jsonStr, localSensorID);
}

/**
 * Secondary thread dedicated to network/sensor polling
 */
static void SensorThread() {
    while (g_Running) {
        g_asyncTemp = GetCPUTempFast();

        DWORD waitMs = (DWORD)SENSOR_LOOP_OFF_DELAY_MS;

        if (g_LedsEnabled) {
            std::lock_guard<std::mutex> lock(g_cfgMutex);
            waitMs = (DWORD)g_cfg.sensorUpdateMS;
        }

        WaitForSingleObject(g_hSensorEvent, waitMs);
    }
}

static void forceLEDRefresh() {
    lastTemp = -1.0f;
    currR = 0.0f; currG = 0.0f; currB = 0.0f;
    lastR = RGB_LED_REFRESH; lastG = RGB_LED_REFRESH; lastB = RGB_LED_REFRESH;
}

/**
 * Extracts MSI SDK DLL from resources if not present locally
 */
static bool ExtractMSIDLL() {
    wchar_t szPath[MAX_PATH];
    GetModuleFileNameW(NULL, szPath, MAX_PATH);
    wchar_t* lastSlash = wcsrchr(szPath, L'\\');
    if (lastSlash) *(lastSlash + 1) = L'\0';

    std::wstring dllPath = std::wstring(szPath) + L"MysticLight_SDK.dll";
    if (GetFileAttributesW(dllPath.c_str()) != INVALID_FILE_ATTRIBUTES) return true;

    HMODULE hModule = GetModuleHandle(NULL);
    HRSRC hRes = FindResourceW(hModule, (LPCWSTR)MAKEINTRESOURCE(IDR_MSI_DLL), L"BINARY");
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

/**
 * MSI Hardware Detection: Scans for compatible devices using RAII for memory safety.
 * Optimized to handle SAFEARRAY life cycles automatically.
 */
static void MSIHwardwareDetection() {
    SAFEARRAY* pDevType = nullptr;
    SAFEARRAY* pLedCount = nullptr;

    // Execute API call
    int hrInfo = MLAPI_ERROR;
    if (lpMLAPI_GetDeviceInfo) {
        hrInfo = lpMLAPI_GetDeviceInfo(&pDevType, &pLedCount);
    }

    if (hrInfo == MLAPI_OK && pDevType && pLedCount) {
        {
            // Use Scoped Locks to ensure memory is released regardless of function exit point
            SafeArrayScopedLock<BSTR> typeLock(pDevType);
            SafeArrayScopedLock<void> countLock(pLedCount);

            if (typeLock.isLocked() && countLock.isLocked()) {
                long lBound = 0, uBound = 0;
                SafeArrayGetLBound(pDevType, 1, &lBound);
                SafeArrayGetUBound(pDevType, 1, &uBound);
                long count = uBound - lBound + 1;

                // Check if we have at least one valid device
                if (count > 0 && typeLock.data != nullptr && typeLock.data[0] != nullptr) {

                    // Clear existing device name to prevent memory leaks
                    if (g_deviceName) {
                        SysFreeString(g_deviceName);
                        g_deviceName = nullptr;
                    }

                    // Store primary device identifier and LED count
                    g_deviceName = SysAllocString(typeLock.data[0]);

                    VARTYPE vtCount;
                    SafeArrayGetVartype(pLedCount, &vtCount);
                    g_totalLeds = GetIntFromSafeArray(countLock.data, vtCount, 0);

                    // Attempt to fetch a friendly name for logs
                    BSTR friendlyName = nullptr;
                    if (lpMLAPI_GetDeviceNameEx) {
                        int hrName = lpMLAPI_GetDeviceNameEx(g_deviceName, 0, &friendlyName);
                        if (hrName != MLAPI_OK) {
                            LogSDKError("MLAPI_GetDeviceNameEx", hrName);
                        }
                    }

                    // Log findings
                    char devInfo[LOG_BUFFER_SIZE];
                    snprintf(devInfo, sizeof(devInfo),
                        "[MysticFight] MSI device selected: %ls (Type: %ls) | Logical Areas: %d",
                        (friendlyName ? friendlyName : L"Unknown"), g_deviceName, g_totalLeds);
                    Log(devInfo);

                    if (friendlyName) SysFreeString(friendlyName);
                }
            }
        } // Locks released automatically

        SafeArrayDestroy(pDevType);
        SafeArrayDestroy(pLedCount);
    }
    else if (hrInfo != MLAPI_OK) {
        LogSDKError("MLAPI_GetDeviceInfo (Hardware Detection)", hrInfo);
    }
}

static void SwitchActiveProfile(HWND hWnd, int index) {
    if (index < 0 || index >= 5) return;
    {
        std::lock_guard<std::mutex> lock(g_cfgMutex);
        g_Global.activeProfileIndex = index;
    }

    SaveSettings();
    g_ConfigVersion++;
    g_lastDataSourceSearchRetry = 0;
    g_ResetHttp = true;
    g_activeSource = DataSource::Searching;
    g_asyncTemp = -1.0f;
    g_pendingStyleChange = true;

    SetEvent(g_hSensorEvent);
    ResetEvent(g_hSourceResolvedEvent);
    forceLEDRefresh();

    wchar_t msg[128];
    swprintf_s(msg, L"Switched to %ls", g_cfg.label);
    ShowNotification(hWnd, L"MysticFight", msg);
}


// ============================================================================
// WINDOW PROCEDURES & UI SUBCLASSING
// ============================================================================

WNDPROC oldEditProc;
static COLORREF g_CustomColors[16] = { 0 };

/**
 * Subclass for Color Edit boxes to trigger standard Color Picker
 */
LRESULT CALLBACK ColorEditSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    if (uMsg == WM_LBUTTONUP) {
        CHOOSECOLOR cc = { 0 };
        cc.lStructSize = sizeof(cc);
        cc.hwndOwner = GetParent(hWnd);
        cc.lpCustColors = g_CustomColors;
        cc.Flags = CC_RGBINIT | CC_FULLOPEN | CC_ANYCOLOR;

        wchar_t buf[10];
        GetWindowTextW(hWnd, buf, 10);
        if (IsValidHex(buf)) {
            cc.rgbResult = HexToColor(buf);
        }

        if (ChooseColor(&cc)) {
            wchar_t newHex[10];
            ColorToHex(cc.rgbResult, newHex, 10);
            SetWindowTextW(hWnd, newHex);
            SendMessage(GetParent(hWnd), WM_COMMAND, MAKEWPARAM(GetDlgCtrlID(hWnd), EN_CHANGE), (LPARAM)hWnd);
        }
    }
    return CallWindowProc(oldEditProc, hWnd, uMsg, wParam, lParam);
}

/**
 * Swaps UI visibility between Profile settings and Global Shortcuts
 */
static void ToggleSettingsLayer(HWND hDlg, bool showShortcuts) {
    int cmdShowProfiles = showShortcuts ? SW_HIDE : SW_SHOW;
    int cmdShowShortcuts = showShortcuts ? SW_SHOW : SW_HIDE;

    const int profileCtrls[] = {
        IDC_GRP_MSI, IDC_GRP_SOURCE, IDC_GRP_THRESHOLDS, IDC_GRP_ADVANCED,
        IDC_EDIT_LABEL, IDC_CHK_ACTIVE_PROFILE, IDC_COMBO_DEVICE, IDC_COMBO_AREA,
        IDC_SENSOR_ID, IDC_EDIT_SERVER, IDC_TEMP_LOW, IDC_TEMP_MED, IDC_TEMP_HIGH,
        IDC_HEX_LOW, IDC_HEX_MED, IDC_HEX_HIGH, IDC_BTN_RESET,
        IDC_COMBO_SENSOR_UPDATE, IDC_COMBO_LED_FPS, IDC_COMBO_SMOOTHING, IDC_BTN_RESET_ADVANCED,
        IDC_LBL_PROFILE_NAME, IDC_LBL_DEVICE, IDC_LBL_AREA, IDC_LBL_TEMP_SENS, IDC_LBL_WEB_SRV,
        IDC_LBL_LOW, IDC_LBL_MED, IDC_LBL_HIGH, IDC_LBL_C1, IDC_LBL_C2, IDC_LBL_C3,
        IDC_LBL_HEX1, IDC_LBL_HEX2, IDC_LBL_HEX3, IDC_LBL_ADV_SENSOR, IDC_LBL_ADV_LED, IDC_LBL_ADV_SMOOTH
    };

    const int shortcutCtrls[] = {
        IDC_GRP_SHORTCUTS, IDC_LBL_TOGGLE, IDC_HK_TOGGLE, IDC_LBL_P1, IDC_HK_P1,
        IDC_LBL_P2, IDC_HK_P2, IDC_LBL_P3, IDC_HK_P3, IDC_LBL_P4, IDC_HK_P4, IDC_LBL_P5, IDC_HK_P5, IDC_BTN_RESET_SHORTCUTS
    };

    for (int id : profileCtrls) { HWND hCtrl = GetDlgItem(hDlg, id); if (hCtrl) ShowWindow(hCtrl, cmdShowProfiles); }
    for (int id : shortcutCtrls) { HWND hCtrl = GetDlgItem(hDlg, id); if (hCtrl) ShowWindow(hCtrl, cmdShowShortcuts); }
}

static bool ValidateHotkeys(HWND hDlg) {
    WORD keys[6];
    keys[0] = (WORD)SendDlgItemMessage(hDlg, IDC_HK_TOGGLE, HKM_GETHOTKEY, 0, 0);
    keys[1] = (WORD)SendDlgItemMessage(hDlg, IDC_HK_P1, HKM_GETHOTKEY, 0, 0);
    keys[2] = (WORD)SendDlgItemMessage(hDlg, IDC_HK_P2, HKM_GETHOTKEY, 0, 0);
    keys[3] = (WORD)SendDlgItemMessage(hDlg, IDC_HK_P3, HKM_GETHOTKEY, 0, 0);
    keys[4] = (WORD)SendDlgItemMessage(hDlg, IDC_HK_P4, HKM_GETHOTKEY, 0, 0);
    keys[5] = (WORD)SendDlgItemMessage(hDlg, IDC_HK_P5, HKM_GETHOTKEY, 0, 0);

    for (int i = 0; i < 6; i++) {
        if (LOBYTE(keys[i]) == 0) continue;
        for (int j = i + 1; j < 6; j++) {
            if (keys[i] == keys[j]) {
                MessageBoxW(hDlg, L"Duplicate hotkey detected.", L"Error", MB_ICONWARNING);
                return false;
            }
        }
    }
    return true;
}

static void SaveHotkeysFromUI(HWND hDlg) {
    g_Global.hotkeys.toggleLEDs = (WORD)SendDlgItemMessage(hDlg, IDC_HK_TOGGLE, HKM_GETHOTKEY, 0, 0);
    g_Global.hotkeys.profile1 = (WORD)SendDlgItemMessage(hDlg, IDC_HK_P1, HKM_GETHOTKEY, 0, 0);
    g_Global.hotkeys.profile2 = (WORD)SendDlgItemMessage(hDlg, IDC_HK_P2, HKM_GETHOTKEY, 0, 0);
    g_Global.hotkeys.profile3 = (WORD)SendDlgItemMessage(hDlg, IDC_HK_P3, HKM_GETHOTKEY, 0, 0);
    g_Global.hotkeys.profile4 = (WORD)SendDlgItemMessage(hDlg, IDC_HK_P4, HKM_GETHOTKEY, 0, 0);
    g_Global.hotkeys.profile5 = (WORD)SendDlgItemMessage(hDlg, IDC_HK_P5, HKM_GETHOTKEY, 0, 0);
}

/**
 * Main Settings Dialog Procedure
 */
INT_PTR CALLBACK SettingsDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    static HBRUSH hBrushLow = NULL, hBrushMed = NULL, hBrushHigh = NULL;
    static int s_currentTab = 0;

    switch (message) {
    case WM_CTLCOLOREDIT: {
        HDC hdc = (HDC)wParam;
        HWND hCtrl = (HWND)lParam;
        int id = GetDlgCtrlID(hCtrl);
        HBRUSH hSelectedBrush = NULL;

        if (id == IDC_HEX_LOW) hSelectedBrush = hBrushLow;
        else if (id == IDC_HEX_MED) hSelectedBrush = hBrushMed;
        else if (id == IDC_HEX_HIGH) hSelectedBrush = hBrushHigh;

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
        HWND hMainWnd = GetParent(hDlg);
        if (hMainWnd) {
            UnregisterHotKey(hMainWnd, 1);
            for (int i = 101; i <= 105; i++) UnregisterHotKey(hMainWnd, i);
        }

        RECT rcDlg; GetWindowRect(hDlg, &rcDlg);
        int dwWidth = rcDlg.right - rcDlg.left;
        int dwHeight = rcDlg.bottom - rcDlg.top;
        int x = (GetSystemMetrics(SM_CXSCREEN) - dwWidth) / 2;
        int y = (GetSystemMetrics(SM_CYSCREEN) - dwHeight) / 2;
        SetWindowPos(hDlg, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE);

        LoadSettings();

        HWND hTab = GetDlgItem(hDlg, IDC_TAB_PROFILES);
        if (hTab) {
            TabCtrl_DeleteAllItems(hTab);
            TCITEMW tie = { 0 }; tie.mask = TCIF_TEXT;
            for (int i = 0; i < 5; i++) {
                tie.pszText = g_Global.profiles[i].label;
                SendMessage(hTab, TCM_INSERTITEMW, i, (LPARAM)&tie);
            }
            tie.pszText = (LPWSTR)L"Shortcuts";
            SendMessage(hTab, TCM_INSERTITEMW, 5, (LPARAM)&tie);
        }

        HWND hSensorUpdate = GetDlgItem(hDlg, IDC_COMBO_SENSOR_UPDATE);
        SendMessage(hSensorUpdate, CB_ADDSTRING, 0, (LPARAM)L"250 ms");
        SendMessage(hSensorUpdate, CB_ADDSTRING, 0, (LPARAM)L"500 ms");
        SendMessage(hSensorUpdate, CB_ADDSTRING, 0, (LPARAM)L"1000 ms");

        HWND hLedFPS = GetDlgItem(hDlg, IDC_COMBO_LED_FPS);
        SendMessage(hLedFPS, CB_ADDSTRING, 0, (LPARAM)L"25 FPS");
        SendMessage(hLedFPS, CB_ADDSTRING, 0, (LPARAM)L"20 FPS");
        SendMessage(hLedFPS, CB_ADDSTRING, 0, (LPARAM)L"15 FPS");

        HWND hSmoothing = GetDlgItem(hDlg, IDC_COMBO_SMOOTHING);
        SendMessage(hSmoothing, CB_ADDSTRING, 0, (LPARAM)L"Disabled");
        SendMessage(hSmoothing, CB_ADDSTRING, 0, (LPARAM)L"Slow");
        SendMessage(hSmoothing, CB_ADDSTRING, 0, (LPARAM)L"Normal");
        SendMessage(hSmoothing, CB_ADDSTRING, 0, (LPARAM)L"Fast");

        PopulateDeviceList(hDlg);

        SendDlgItemMessage(hDlg, IDC_HK_TOGGLE, HKM_SETHOTKEY, g_Global.hotkeys.toggleLEDs, 0);
        SendDlgItemMessage(hDlg, IDC_HK_P1, HKM_SETHOTKEY, g_Global.hotkeys.profile1, 0);
        SendDlgItemMessage(hDlg, IDC_HK_P2, HKM_SETHOTKEY, g_Global.hotkeys.profile2, 0);
        SendDlgItemMessage(hDlg, IDC_HK_P3, HKM_SETHOTKEY, g_Global.hotkeys.profile3, 0);
        SendDlgItemMessage(hDlg, IDC_HK_P4, HKM_SETHOTKEY, g_Global.hotkeys.profile4, 0);
        SendDlgItemMessage(hDlg, IDC_HK_P5, HKM_SETHOTKEY, g_Global.hotkeys.profile5, 0);

        s_currentTab = g_Global.activeProfileIndex;
        TabCtrl_SetCurSel(hTab, s_currentTab);

        if (s_currentTab == 5) ToggleSettingsLayer(hDlg, true);
        else {
            ToggleSettingsLayer(hDlg, false);
            LoadProfileToUI(hDlg, s_currentTab);
        }

        int safeTab = (s_currentTab < 5) ? s_currentTab : 0;
        hBrushLow = CreateSolidBrush(g_Global.profiles[safeTab].colorLow);
        hBrushMed = CreateSolidBrush(g_Global.profiles[safeTab].colorMed);
        hBrushHigh = CreateSolidBrush(g_Global.profiles[safeTab].colorHigh);

        oldEditProc = (WNDPROC)SetWindowLongPtr(GetDlgItem(hDlg, IDC_HEX_LOW), GWLP_WNDPROC, (LONG_PTR)ColorEditSubclassProc);
        SetWindowLongPtr(GetDlgItem(hDlg, IDC_HEX_MED), GWLP_WNDPROC, (LONG_PTR)ColorEditSubclassProc);
        SetWindowLongPtr(GetDlgItem(hDlg, IDC_HEX_HIGH), GWLP_WNDPROC, (LONG_PTR)ColorEditSubclassProc);

        if (ValidStartupTaskExists()) CheckDlgButton(hDlg, IDC_CHK_STARTUP, BST_CHECKED);

        return (INT_PTR)TRUE;
    }

    case WM_NOTIFY: {
        LPNMHDR pnm = (LPNMHDR)lParam;
        if (pnm->idFrom == IDC_TAB_PROFILES && pnm->code == TCN_SELCHANGE) {
            int newTab = TabCtrl_GetCurSel(pnm->hwndFrom);

            if (s_currentTab >= 0 && s_currentTab <= 4) SaveUIToProfile(hDlg, s_currentTab);
            else if (s_currentTab == 5) {
                if (!ValidateHotkeys(hDlg)) {
                    TabCtrl_SetCurSel(pnm->hwndFrom, 5);
                    return TRUE;
                }
                SaveHotkeysFromUI(hDlg);
            }

            s_currentTab = newTab;

            if (s_currentTab == 5) ToggleSettingsLayer(hDlg, true);
            else {
                ToggleSettingsLayer(hDlg, false);
                if (hBrushLow) DeleteObject(hBrushLow);
                if (hBrushMed) DeleteObject(hBrushMed);
                if (hBrushHigh) DeleteObject(hBrushHigh);
                hBrushLow = CreateSolidBrush(g_Global.profiles[s_currentTab].colorLow);
                hBrushMed = CreateSolidBrush(g_Global.profiles[s_currentTab].colorMed);
                hBrushHigh = CreateSolidBrush(g_Global.profiles[s_currentTab].colorHigh);
                LoadProfileToUI(hDlg, s_currentTab);
            }
        }
        break;
    }

    case WM_COMMAND: {
        int id = LOWORD(wParam);
        int code = HIWORD(wParam);

        if (id == IDC_CHK_ACTIVE_PROFILE && code == BN_CLICKED) {
            if (IsDlgButtonChecked(hDlg, IDC_CHK_ACTIVE_PROFILE)) {
                if (s_currentTab < 5) g_Global.activeProfileIndex = s_currentTab;
            }
            else {
                if (g_Global.activeProfileIndex == s_currentTab)
                    CheckDlgButton(hDlg, IDC_CHK_ACTIVE_PROFILE, BST_CHECKED);
            }
            return TRUE;
        }

        if (id == IDC_COMBO_DEVICE && code == CBN_SELCHANGE) {
            int idx = (int)SendMessage((HWND)lParam, CB_GETCURSEL, 0, 0);
            wchar_t* devType = (wchar_t*)SendMessage((HWND)lParam, CB_GETITEMDATA, idx, 0);
            if (devType) PopulateAreaList(hDlg, devType, 0);
            return TRUE;
        }

        if (code == EN_CHANGE) {
            if (id == IDC_HEX_LOW || id == IDC_HEX_MED || id == IDC_HEX_HIGH) {
                wchar_t buf[10]; GetDlgItemTextW(hDlg, id, buf, 10);
                if (IsValidHex(buf)) {
                    COLORREF newCol = HexToColor(buf);
                    HBRUSH newBrush = CreateSolidBrush(newCol);
                    if (id == IDC_HEX_LOW) { if (hBrushLow) DeleteObject(hBrushLow); hBrushLow = newBrush; }
                    else if (id == IDC_HEX_MED) { if (hBrushMed) DeleteObject(hBrushMed); hBrushMed = newBrush; }
                    else if (id == IDC_HEX_HIGH) { if (hBrushHigh) DeleteObject(hBrushHigh); hBrushHigh = newBrush; }
                    InvalidateRect(GetDlgItem(hDlg, id), NULL, TRUE);
                }
            }
            else if (id == IDC_EDIT_LABEL) {
                wchar_t newLabel[64]; GetDlgItemTextW(hDlg, IDC_EDIT_LABEL, newLabel, 64);
                HWND hTab = GetDlgItem(hDlg, IDC_TAB_PROFILES);
                if (hTab && s_currentTab < 5) {
                    TCITEMW tie = { 0 }; tie.mask = TCIF_TEXT; tie.pszText = newLabel;
                    SendMessage(hTab, TCM_SETITEMW, s_currentTab, (LPARAM)&tie);
                }
            }
            return TRUE;
        }

        if (id == IDOK) {
            if (s_currentTab <= 4) SaveUIToProfile(hDlg, s_currentTab);
            if (ValidateHotkeys(hDlg)) SaveHotkeysFromUI(hDlg);

            SetStartupTask(IsDlgButtonChecked(hDlg, IDC_CHK_STARTUP) == BST_CHECKED);
            SaveSettings();
            g_ConfigVersion++;
            g_ResetHttp = true;
            forceLEDRefresh();
            SetEvent(g_hSensorEvent);
            EndDialog(hDlg, IDOK);
            return TRUE;
        }

        if (id == IDCANCEL) {
            EndDialog(hDlg, IDCANCEL);
            return TRUE;
        }

        if (id == IDC_BTN_RESET_SHORTCUTS) {
            BYTE defMod = HOTKEYF_CONTROL | HOTKEYF_SHIFT | HOTKEYF_ALT;
            SendDlgItemMessage(hDlg, IDC_HK_TOGGLE, HKM_SETHOTKEY, MAKEWORD(0x4C, defMod), 0);
            for (int i = 0; i < 5; i++)
                SendDlgItemMessage(hDlg, IDC_HK_P1 + i, HKM_SETHOTKEY, MAKEWORD('1' + i, defMod), 0);
            return TRUE;
        }

        if (id == IDC_BTN_RESET) {
            SetDlgItemInt(hDlg, IDC_TEMP_LOW, 50, TRUE);
            SetDlgItemInt(hDlg, IDC_TEMP_MED, 70, TRUE);
            SetDlgItemInt(hDlg, IDC_TEMP_HIGH, 90, TRUE);
            SetDlgItemTextW(hDlg, IDC_HEX_LOW, L"#00FF00");
            SetDlgItemTextW(hDlg, IDC_HEX_MED, L"#FFFF00");
            SetDlgItemTextW(hDlg, IDC_HEX_HIGH, L"#FF0000");
            SendMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDC_HEX_LOW, EN_CHANGE), 0);
            SendMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDC_HEX_MED, EN_CHANGE), 0);
            SendMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDC_HEX_HIGH, EN_CHANGE), 0);
            return TRUE;
        }

        if (id == IDC_BTN_RESET_ADVANCED) {
            SendMessage(GetDlgItem(hDlg, IDC_COMBO_SENSOR_UPDATE), CB_SETCURSEL, 1, 0);
            SendMessage(GetDlgItem(hDlg, IDC_COMBO_LED_FPS), CB_SETCURSEL, 0, 0);
            SendMessage(GetDlgItem(hDlg, IDC_COMBO_SMOOTHING), CB_SETCURSEL, 2, 0);
            return TRUE;
        }
    } break;

    case WM_DESTROY:
        RegisterAppHotkeys(GetParent(hDlg));
        if (hBrushLow) DeleteObject(hBrushLow);
        if (hBrushMed) DeleteObject(hBrushMed);
        if (hBrushHigh) DeleteObject(hBrushHigh);
        ClearComboHeapData(GetDlgItem(hDlg, IDC_SENSOR_ID));
        ClearComboHeapData(GetDlgItem(hDlg, IDC_COMBO_DEVICE));
        SetWindowLongPtr(GetDlgItem(hDlg, IDC_HEX_LOW), GWLP_WNDPROC, (LONG_PTR)oldEditProc);
        SetWindowLongPtr(GetDlgItem(hDlg, IDC_HEX_MED), GWLP_WNDPROC, (LONG_PTR)oldEditProc);
        SetWindowLongPtr(GetDlgItem(hDlg, IDC_HEX_HIGH), GWLP_WNDPROC, (LONG_PTR)oldEditProc);
        break;
    }
    return (INT_PTR)FALSE;
}

/**
 * About Dialog Procedure
 */
INT_PTR CALLBACK AboutDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    static HFONT hFontLink = NULL;

    switch (message) {
    case WM_INITDIALOG: {
        HWND hLink = GetDlgItem(hDlg, IDC_GITHUB_LINK);
        SetClassLongPtr(hLink, GCLP_HCURSOR, (LONG_PTR)LoadCursor(NULL, IDC_HAND));

        RECT rcDlg; GetWindowRect(hDlg, &rcDlg);
        int dwWidth = rcDlg.right - rcDlg.left; int dwHeight = rcDlg.bottom - rcDlg.top;
        int x = (GetSystemMetrics(SM_CXSCREEN) - dwWidth) / 2; int y = (GetSystemMetrics(SM_CYSCREEN) - dwHeight) / 2;
        SetWindowPos(hDlg, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE);

        HICON hIcon = (HICON)LoadImage(GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_ICON1), IMAGE_ICON, 48, 48, LR_SHARED);
        SendMessage(GetDlgItem(hDlg, IDC_ABOUT_ICON), STM_SETIMAGE, IMAGE_ICON, (LPARAM)hIcon);

        std::wstring versionStr = L"MysticFight " + std::wstring(APP_VERSION);
        SetDlgItemTextW(hDlg, IDC_ABOUT_VERSION, versionStr.c_str());

        HFONT hFont = (HFONT)SendMessage(GetDlgItem(hDlg, IDC_GITHUB_LINK), WM_GETFONT, 0, 0);
        LOGFONT lf; GetObject(hFont, sizeof(LOGFONT), &lf);
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
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        if (LOWORD(wParam) == IDC_GITHUB_LINK) RunShellNonAdmin(L"https://github.com/tonikelope/MysticFight");
        break;
    case WM_DESTROY:
        if (hFontLink) { DeleteObject(hFontLink); hFontLink = NULL; }
        break;
    }
    return (INT_PTR)FALSE;
}


// ============================================================================
// SYSTEM CLEANUP & MESSAGING
// ============================================================================

/**
 * Releases all handles, libraries, and resets hardware state before exit
 */
static void FinalCleanup(HWND hWnd) {
    Log("[MysticFight] Starting cleaning...");

    if (g_hLibrary) {
        if (lpMLAPI_SetLedStyle) {
            _bstr_t bstrOff(L"Off");
            Config& activeProfile = g_Global.profiles[g_Global.activeProfileIndex];
            _bstr_t bstrTarget(activeProfile.targetDevice);
            int hrStyle = lpMLAPI_SetLedStyle(bstrTarget, activeProfile.targetLedIndex, bstrOff);
            if (hrStyle == MLAPI_OK) {
                Log("[MysticFight] LEDs power off");
            }
            else {
                LogSDKError("MLAPI_SetLedStyle (Cleanup)", hrStyle);
            }
        }

        if (lpMLAPI_Release) {
            int hrRelease = lpMLAPI_Release();
            if (hrRelease == MLAPI_OK) {
                Log("[MysticFight] MSI SDK Released");
            }
            else {
                LogSDKError("MLAPI_Release", hrRelease);
            }
        }
        FreeLibrary(g_hLibrary);
        g_hLibrary = NULL;
    }

    if (g_deviceName) { SysFreeString(g_deviceName); g_deviceName = NULL; }

    if (hWnd) {
        NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA), hWnd, 1 };
        Shell_NotifyIcon(NIM_DELETE, &nid);
        UnregisterHotKey(hWnd, 1);
        for (int i = 101; i <= 105; i++) UnregisterHotKey(hWnd, i);
        Log("[MysticFight] Global Hot Keys Unregistered");
    }

    CoUninitialize();

    if (g_hSensorEvent) { CloseHandle(g_hSensorEvent); g_hSensorEvent = NULL; }
    if (g_hSourceResolvedEvent) { CloseHandle(g_hSourceResolvedEvent); g_hSourceResolvedEvent = NULL; }
    if (g_hMutex) { ReleaseMutex(g_hMutex); CloseHandle(g_hMutex); g_hMutex = NULL; }

    Log("[MysticFight] Cleaning finished.");
}

/**
 * Main Application Window Procedure
 */
static LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_HOTKEY:
        if (wParam == 1) {
            g_LedsEnabled = !g_LedsEnabled;
            SetEvent(g_hSensorEvent);
            lastR = RGB_LED_REFRESH;
            PostMessage(hWnd, WM_NULL, 0, 0);

            wchar_t msg[128];
            swprintf_s(msg, L"Lights %ls", g_LedsEnabled ? L"ON" : L"OUT");
            ShowNotification(hWnd, L"MysticFight", msg);

            PlaySound(MAKEINTRESOURCE(g_LedsEnabled ? IDR_WAV_LIGHTS_ON : IDR_WAV_LIGHTS_OFF), GetModuleHandle(NULL), SND_RESOURCE | SND_ASYNC);
        }
        else if (wParam >= 101 && wParam <= 105) {
            SwitchActiveProfile(hWnd, (int)wParam - 101);
        }
        break;

    case WM_TRAYICON:
        if (lParam == WM_RBUTTONUP) {
            POINT curPoint; GetCursorPos(&curPoint);
            HMENU hMenu = CreatePopupMenu();

            if (hMenu) {
                // Profile Submenu Creation
                HMENU hProfileMenu = CreatePopupMenu();
                for (int i = 0; i < 5; i++) {
                    UINT flags = MF_STRING | (i == g_Global.activeProfileIndex ? MF_CHECKED : 0);
                    AppendMenuW(hProfileMenu, flags, ID_TRAY_PROFILE_BASE + i, g_Global.profiles[i].label);
                }

                AppendMenuW(hMenu, MF_POPUP, (UINT_PTR)hProfileMenu, L"Active Profile");
                AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
                InsertMenuW(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_CONFIG, L"Settings");
                InsertMenuW(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_ABOUT, L"About MysticFight");
                InsertMenuW(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_LOG, L"View Debug Log");
                InsertMenuW(hMenu, -1, MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
                InsertMenuW(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_EXIT, L"Exit MysticFight");

                SetForegroundWindow(hWnd);
                TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, curPoint.x, curPoint.y, 0, hWnd, NULL);
                DestroyMenu(hMenu);
            }
        }
        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == ID_TRAY_EXIT) g_Running = false;

        // Handle Profile Switching via Tray
        if (LOWORD(wParam) >= ID_TRAY_PROFILE_BASE && LOWORD(wParam) <= ID_TRAY_PROFILE_BASE + 4) {
            SwitchActiveProfile(hWnd, LOWORD(wParam) - ID_TRAY_PROFILE_BASE);
        }

        if (LOWORD(wParam) == ID_TRAY_LOG) {
            wchar_t fullLogPath[MAX_PATH]; GetFullPathNameW(LOG_FILENAME, MAX_PATH, fullLogPath, NULL);
            RunShellNonAdmin(fullLogPath);
        }

        if (!g_Resetting_sdk && LOWORD(wParam) == ID_TRAY_CONFIG) {
            if (g_SettingsOpen) {
                HWND hExisting = FindWindowW(L"#32770", L"Settings");
                if (hExisting) SetForegroundWindow(hExisting);
                break;
            }
            g_SettingsOpen = true;
            DialogBoxParamW(GetModuleHandle(NULL), MAKEINTRESOURCEW(IDD_SETTINGS), hWnd, (DLGPROC)SettingsDlgProc, 0);
            g_SettingsOpen = false;
        }

        if (LOWORD(wParam) == ID_TRAY_ABOUT) {
            if (g_AboutOpen) {
                HWND hExistingAbout = FindWindowW(L"#32770", L"About MysticFight");
                if (hExistingAbout) SetForegroundWindow(hExistingAbout);
                break;
            }
            g_AboutOpen = true;
            DialogBoxParamW(GetModuleHandle(NULL), MAKEINTRESOURCEW(IDD_ABOUT), hWnd, (DLGPROC)AboutDlgProc, 0);
            g_AboutOpen = false;
        }
        break;

    case WM_QUERYENDSESSION:
        if (!g_windows_shutdown) {
            Log("[MysticFight] Windows Shutdown detected...");
            g_windows_shutdown = true; g_Running = false;
        }
        return FALSE;

    case WM_CLOSE:
    case WM_DESTROY:
        g_Running = false; PostQuitMessage(0); return 0;
    }
    return DefWindowProc(hWnd, message, wParam, lParam);
}

/**
 * Dynamically binds MSI SDK functions from DLL
 */
static bool BindMSISDK(HMODULE hLib) {
    if (!hLib) return false;

    lpMLAPI_Initialize = (LPMLAPI_Initialize)GetProcAddress(hLib, "MLAPI_Initialize");
    lpMLAPI_GetDeviceInfo = (LPMLAPI_GetDeviceInfo)GetProcAddress(hLib, "MLAPI_GetDeviceInfo");
    lpMLAPI_SetLedColor = (LPMLAPI_SetLedColor)GetProcAddress(hLib, "MLAPI_SetLedColor");
    lpMLAPI_SetLedStyle = (LPMLAPI_SetLedStyle)GetProcAddress(hLib, "MLAPI_SetLedStyle");
    lpMLAPI_SetLedSpeed = (LPMLAPI_SetLedSpeed)GetProcAddress(hLib, "MLAPI_SetLedSpeed");
    lpMLAPI_Release = (LPMLAPI_Release)GetProcAddress(hLib, "MLAPI_Release");
    lpMLAPI_GetDeviceNameEx = (LPMLAPI_GetDeviceNameEx)GetProcAddress(hLib, "MLAPI_GetDeviceNameEx");
    lpMLAPI_GetLedInfo = (LPMLAPI_GetLedInfo)GetProcAddress(hLib, "MLAPI_GetLedInfo");

    if (!lpMLAPI_Initialize || !lpMLAPI_SetLedColor || !lpMLAPI_Release) {
        Log("[MysticFight] Error: Critical MSI SDK functions missing.");
        return false;
    }
    return true;
}
/**
 * Executes an SDK function and manages failure counters
 */
static bool SafeSDKCall(int result, const char* funcName) {
    if (result == MLAPI_OK) {
        g_sdkFailCount = 0; // Reset count on success
        return true;
    }

    LogSDKError(funcName, result);
    g_sdkFailCount++;

    if (g_sdkFailCount >= MAX_SDK_FAILURES && !g_Resetting_sdk) {
        Log("[MysticFight] Critical failure threshold reached. SDK Reset triggered.");
        g_Resetting_sdk = true;
        g_ResetStage = 0;
    }
    return false;
}

static void ShowLetsDanceNotification(HWND hWnd) {
    wchar_t startTitle[128];
    swprintf_s(startTitle, L"MysticFight %ls (by tonikelope)", APP_VERSION);
    ShowNotification(hWnd, startTitle, L"Let's dance baby");
}

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    // Basic Process Initialization
    INITCOMMONCONTROLSEX icce = { sizeof(icce), ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES };
    InitCommonControlsEx(&icce);
    SetPriorityClass(GetCurrentProcess(), HIGH_PRIORITY_CLASS);
    SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);

    if (FAILED(CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE))) return 1;

    // Instance Check
    g_hMutex = CreateMutexW(NULL, TRUE, L"Global\\MysticFight_Unique_Mutex");
    if (g_hMutex == NULL || GetLastError() == ERROR_ALREADY_EXISTS) {
        if (g_hMutex) {
            MessageBoxW(NULL, L"MysticFight is already running.", L"Information", MB_OK | MB_ICONINFORMATION);
            CloseHandle(g_hMutex);
        }
        CoUninitialize();
        return 0;
    }

    HWND hWnd = NULL;
    std::thread sThread;

    try {
        {
            _bstr_t bstrOff(L"Off");
            _bstr_t bstrSteady(L"Steady");
            _bstr_t bstrBreath(L"Breath");
            _bstr_t bstrDevice(L"");

            // Setup environment
            if (!ExtractMSIDLL() || !(g_hLibrary = LoadLibraryW(L"MysticLight_SDK.dll")))
                throw std::runtime_error("DLL preparation failed");

            if (!BindMSISDK(g_hLibrary)) throw std::runtime_error("SDK Binding failed");

            TrimLogFile();

            char versionMsg[128];
            snprintf(versionMsg, sizeof(versionMsg), "[MysticFight] MysticFight %ls started", APP_VERSION);
            Log(versionMsg);

            bool isFirstRun = (GetFileAttributesW(INI_FILE) == INVALID_FILE_ATTRIBUTES);
            LoadSettings();

            // Register Window Class
            WNDCLASSW wc = { 0 };
            wc.lpfnWndProc = WndProc;
            wc.hInstance = hInstance;
            wc.lpszClassName = L"MysticFight_Class";
            wc.hIcon = (HICON)LoadImageW(hInstance, MAKEINTRESOURCEW(IDI_ICON1), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
            RegisterClassW(&wc);
            hWnd = CreateWindowExW(0, wc.lpszClassName, L"MysticFight", 0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);

            // Tray Icon Setup
            NOTIFYICONDATAW nid = { sizeof(NOTIFYICONDATAW), hWnd, 1, NIF_ICON | NIF_MESSAGE | NIF_TIP, WM_TRAYICON };
            nid.hIcon = (HICON)LoadImageW(hInstance, MAKEINTRESOURCEW(IDI_ICON1), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
            swprintf_s(nid.szTip, L"MysticFight %ls - Initializing...", APP_VERSION);
            Shell_NotifyIconW(NIM_ADD, &nid);

            // Init Engine
            Log("[MysticFight] Initializing MSI SDK and LHM HTTP Sensor Thread...");

            // STARTUP LOGIC
            int hrInit = lpMLAPI_Initialize();

            if (hrInit == MLAPI_OK) {
                Log("[MysticFight] SDK Initialized successfully");
                MSIHwardwareDetection();
            }
            else {
                LogSDKError("MLAPI_Initialize (Startup)", hrInit);

                if (hrInit == MLAPI_NO_IMPLEMENTED || hrInit == MLAPI_ERROR) {
                    Log("[MysticFight] SDK service unavailable. Attempting auto-repair...");
                    g_Resetting_sdk = true;
                    g_ResetStage = 0;
                    g_ResetTimer = GetTickCount64();
                }
                else {
                    MessageBoxW(NULL, L"MSI SDK Critical Error. Please check MSI Center installation.", L"MysticFight", MB_OK | MB_ICONERROR);
                    throw std::runtime_error("MSI SDK Initializing failed");
                }
            }

            g_hSensorEvent = CreateEventW(NULL, FALSE, FALSE, NULL);
            g_hSourceResolvedEvent = CreateEventW(NULL, TRUE, FALSE, NULL);
            sThread = std::thread(SensorThread);

            // Await LHM connection
            WaitForSingleObject(g_hSourceResolvedEvent, LHM_BOOT_TIMEOUT);

            if (g_activeSource == DataSource::Searching) {
                MessageBoxW(NULL, L"Fatal Error: No temperature data source found (http://localhost:8085).", L"MysticFight - Startup Error", MB_OK | MB_ICONERROR);
                throw std::runtime_error("LHM Data Source not found. Is LibreHardwareMonitor running and its web server enabled?");
            }

            if (isFirstRun) SaveSettings();

            if (isFirstRun) DialogBoxParamW(hInstance, MAKEINTRESOURCEW(IDD_SETTINGS), hWnd, (DLGPROC)SettingsDlgProc, 0);

            RegisterAppHotkeys(hWnd);

            // Loop Logic State
            int lastProcessedVersion = -1;
            ULONGLONG nextFrameTime = 0;
            DWORD targetR = 0, targetG = 0, targetB = 0;
            MSG msg = { 0 };
            bool firstRunLoop = true;

            if (!g_Resetting_sdk) {
                UpdateStatus(hWnd, NULL);
                ShowLetsDanceNotification(hWnd);
            }

            // ============================================================================
            // MAIN APPLICATION LOOP
            // ============================================================================
            while (g_Running) {
                while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
                    if (msg.message == WM_QUIT) { g_Running = false; break; }
                    TranslateMessage(&msg); DispatchMessage(&msg);
                }
                if (!g_Running) break;

                Config cfgLocal;
                {
                    std::lock_guard<std::mutex> lock(g_cfgMutex);
                    cfgLocal = g_cfg;
                }

                if (g_ConfigVersion != lastProcessedVersion) {
                    bstrDevice = cfgLocal.targetDevice;
                    lastProcessedVersion = g_ConfigVersion;
                    forceLEDRefresh();
                    Log("[MysticFight] Config refreshed");
                }

                ULONGLONG currentTime = GetTickCount64();

                // ------------------------------------------------------------------------
                // SDK RECOVERY LOGIC
                // ------------------------------------------------------------------------
                if (g_Resetting_sdk) {
                    if (currentTime >= g_ResetTimer) {
                        switch (g_ResetStage) {
                        case 0:
                            Log("[MysticFight] Restarting MSI Service...");
                            UpdateStatus(hWnd, L"Restarting MSI Service...");
                            ControlScheduledTask(L"MSI Task Host - LEDKeeper2_Host", false);
                            KillProcessByName(L"LEDKeeper2.exe");
                            g_ResetTimer = currentTime + KILL_LEDKEEPER_TASK_WAIT_MS;
                            g_ResetStage = 1;
                            break;
                        case 1:
                            ControlScheduledTask(L"MSI Task Host - LEDKeeper2_Host", true);
                            g_ResetTimer = currentTime + RESET_LEDKEEPER_TASK_DELAY_MS;
                            g_ResetStage = 2;
                            break;
                        case 2:
                            Log("[MysticFight] Connecting to SDK...");
                            UpdateStatus(hWnd, L"Connecting to SDK...");

                            int hr = lpMLAPI_Initialize();

                            if (hr == MLAPI_OK) {
                                Log("[MysticFight] SDK Online.");
                                MSIHwardwareDetection(); // Reload arrays
                                g_Resetting_sdk = false;
                                g_resetCounter = 0;
                                g_pendingStyleChange = true;
                                forceLEDRefresh();
                                UpdateStatus(hWnd, NULL);
                                ShowLetsDanceNotification(hWnd);
                            }
                            else {
                                g_resetCounter++;
                                LogSDKError("MLAPI_Initialize", hr);

                                if (g_resetCounter >= SDK_MAX_RESETS) {
                                    MessageBoxW(NULL, L"MSI SDK Fatal Error. Please reinstall MSI Center.", L"MysticFight Error", MB_OK | MB_ICONERROR);
                                    g_Running = false;
                                }
                                else {
                                    g_ResetStage = 0;
                                    g_ResetTimer = currentTime + RESET_SDK_RETRY_DELAY_MS;
                                }
                            }
                            break;
                        }
                    }
                    MsgWaitForMultipleObjectsEx(0, NULL, 100, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
                    continue;
                }

                // ------------------------------------------------------------------------
                // COLOR LOGIC (WITH FPS COMPENSATION)
                // ------------------------------------------------------------------------
                if (g_activeSource == DataSource::Searching) {
                    if (g_pendingStyleChange) {
                        g_pendingStyleChange = false;
                        if (SafeSDKCall(lpMLAPI_SetLedStyle(bstrDevice, cfgLocal.targetLedIndex, bstrBreath), "SetLedStyle (Search)")) {
                            SafeSDKCall(lpMLAPI_SetLedColor(bstrDevice, cfgLocal.targetLedIndex, 255, 255, 255), "SetLedColor (Search)");
                        }
                        lastR = RGB_LED_REFRESH;
                    }
                }
                else if (!g_LedsEnabled) {
                    if (lastR != RGB_LEDS_OFF || g_pendingStyleChange) {
                        g_pendingStyleChange = false;
                        if (SafeSDKCall(lpMLAPI_SetLedStyle(bstrDevice, cfgLocal.targetLedIndex, bstrOff), "SetLedStyle (Off)")) {
                            lastR = RGB_LEDS_OFF;
                        }
                    }
                }
                else {
                    if (currentTime >= nextFrameTime) {
                        nextFrameTime = currentTime + (1000 / cfgLocal.ledRefreshFPS);

                        float rawTemp = g_asyncTemp;

                        if (rawTemp >= 0.0f) {
                            if (g_pendingStyleChange) { g_pendingStyleChange = false; forceLEDRefresh(); }

                            float temp = floorf(rawTemp * 2.0f + 0.5f) / 2.0f;

                            if (temp != lastTemp) {
                                lastTemp = temp;
                                float ratio = 0.0f;
                                COLORREF c1 = 0, c2 = 0;

                                // Threshold Logic
                                if (temp <= (float)cfgLocal.tempLow) {
                                    targetR = GetRValue(cfgLocal.colorLow); targetG = GetGValue(cfgLocal.colorLow); targetB = GetBValue(cfgLocal.colorLow);
                                }
                                else if (temp < (float)cfgLocal.tempMed) {
                                    float range = ((float)cfgLocal.tempMed - (float)cfgLocal.tempLow);
                                    ratio = range > 0.0f ? (temp - (float)cfgLocal.tempLow) / range : 0.0f;
                                    c1 = cfgLocal.colorLow; c2 = cfgLocal.colorMed;
                                }
                                else if (temp < (float)cfgLocal.tempHigh) {
                                    float range = ((float)cfgLocal.tempHigh - (float)cfgLocal.tempMed);
                                    ratio = range > 0.0f ? (temp - (float)cfgLocal.tempMed) / range : 0.0f;
                                    c1 = cfgLocal.colorMed; c2 = cfgLocal.colorHigh;
                                }
                                else {
                                    targetR = GetRValue(cfgLocal.colorHigh); targetG = GetGValue(cfgLocal.colorHigh); targetB = GetBValue(cfgLocal.colorHigh);
                                }

                                if (c1 != 0 || c2 != 0) {
                                    double r1 = (double)GetRValue(c1), g1 = (double)GetGValue(c1), b1 = (double)GetBValue(c1);
                                    double r2 = (double)GetRValue(c2), g2 = (double)GetGValue(c2), b2 = (double)GetBValue(c2);
                                    targetR = (DWORD)sqrt(r1 * r1 * (1.0 - ratio) + r2 * r2 * ratio);
                                    targetG = (DWORD)sqrt(g1 * g1 * (1.0 - ratio) + g2 * g2 * ratio);
                                    targetB = (DWORD)sqrt(b1 * b1 * (1.0 - ratio) + b2 * b2 * ratio);
                                }

                                if (firstRunLoop) { currR = (float)targetR; currG = (float)targetG; currB = (float)targetB; firstRunLoop = false; }
                            }

                            // FPS Compensation for Smoothing
                            float timeCompensation = 25.0f / (float)(cfgLocal.ledRefreshFPS > 0 ? cfgLocal.ledRefreshFPS : 25);
                            float adjustedFactor = cfgLocal.smoothingFactor * timeCompensation;
                            if (adjustedFactor > 1.0f) adjustedFactor = 1.0f;

                            // Apply Smoothing
                            currR += ((float)targetR - currR) * adjustedFactor;
                            currG += ((float)targetG - currG) * adjustedFactor;
                            currB += ((float)targetB - currB) * adjustedFactor;

                            DWORD sendR = (DWORD)roundf(currR), sendG = (DWORD)roundf(currG), sendB = (DWORD)roundf(currB);

                            if (sendR != lastR || sendG != lastG || sendB != lastB) {
                                bool success = true;
                                if (lastR == RGB_LEDS_OFF || lastR == RGB_LED_REFRESH) {
                                    success = SafeSDKCall(lpMLAPI_SetLedStyle(bstrDevice, cfgLocal.targetLedIndex, bstrSteady), "SetLedStyle (Steady)");
                                }

                                if (success) {
                                    if (SafeSDKCall(lpMLAPI_SetLedColor(bstrDevice, cfgLocal.targetLedIndex, sendR, sendG, sendB), "SetLedColor")) {
                                        lastR = sendR; lastG = sendG; lastB = sendB;
                                    }
                                }
                            }
                        }
                    }
                }

                DWORD dwWait;
                if (!g_LedsEnabled) dwWait = (DWORD)MAIN_LOOP_OFF_DELAY_MS;
                else {
                    ULONGLONG now = GetTickCount64();
                    if (nextFrameTime > now) {
                        dwWait = (DWORD)min((ULONGLONG)(1000 / cfgLocal.ledRefreshFPS), nextFrameTime - now);
                    }
                    else dwWait = 0;
                }

                MsgWaitForMultipleObjectsEx(0, NULL, dwWait, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
            }
        }
    }
    catch (const std::exception& e) {
        char errBuf[256];
        snprintf(errBuf, sizeof(errBuf), "[MysticFight] CRITICAL EXCEPTION: %s", e.what());
        Log(errBuf);
        MessageBoxA(NULL, errBuf, "MysticFight Crash", MB_OK | MB_ICONERROR);
    }
    catch (...) {
        Log("[MysticFight] CRITICAL UNKNOWN EXCEPTION occurred.");
        MessageBoxW(NULL, L"An unknown critical error occurred.", L"MysticFight Crash", MB_OK | MB_ICONERROR);
    }

    // Termination Procedure
    g_Running = false;
    SetEvent(g_hSensorEvent);

    if (sThread.joinable()) sThread.join();

    if (g_hConnect) WinHttpCloseHandle(g_hConnect);
    if (g_hSession) WinHttpCloseHandle(g_hSession);

    FinalCleanup(hWnd);
    Log("[MysticFight] BYE BYE");

    return 0;
}