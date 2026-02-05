/**
 * ============================================================================
 * MysticFight - RGB Control & Hardware Monitor Integration
 * ============================================================================
 *
 * Description:
 * Win32 application acting as a bridge between LibreHardwareMonitor (HTTP)
 * and MSI Mystic Light SDK. It dynamically adjusts RGB lighting based on
 * system temperature sensors with profile support and background persistence.
 *
 * Architecture: Synchronous UI with threaded sensor polling.
 *
 * ============================================================================
 */

#include "resource.h"
#include <algorithm>
#include <atomic>
#include <charconv>
#include <comdef.h>
#include <comip.h>
#include <commdlg.h>
#include <ctime>
#include <exdisp.h>
#include <fstream>
#include <functional>
#include <math.h>
#include <mutex>
#include <oleauto.h>
#include <shellapi.h>
#include <shlobj.h>
#include <string>
#include <string_view>
#include <taskschd.h>
#include <thread>
#include <tlhelp32.h>
#include <vector>
#include <windows.h>
#include <winhttp.h>

 // ============================================================================
 // LIBRARY LINKING
 // ============================================================================
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "winmm.lib")

// ============================================================================
// CONSTANTS & MACROS
// ============================================================================

const wchar_t* APP_VERSION = L"v2.59";
const wchar_t* LOG_FILENAME = L"debug.log";
const wchar_t* INI_FILE = L".\\config.ini";
const wchar_t* TASK_NAME = L"MysticFight";

// UI & Tray Identifiers
#define WM_TRAYICON             (WM_USER + 1)
#define ID_TRAY_EXIT            1001
#define ID_TRAY_CONFIG          2001
#define ID_TRAY_LOG             3001
#define ID_TRAY_ABOUT           4001

// Logic Sentinels
const DWORD RGB_LED_REFRESH = 999;  // Signals a mandatory style refresh
const DWORD RGB_LEDS_OFF = 1000; // Signals hardware is in OFF state

// Timings
const ULONGLONG MAIN_LOOP_OFF_DELAY_MS = 1000;
const ULONGLONG SENSOR_POLL_INTERVAL_MS = 500;
const ULONGLONG LHM_RETRY_DELAY_MS = 5000;
const ULONGLONG RESET_KILL_TASK_WAIT_MS = 2000;
const ULONGLONG RESET_RESTART_TASK_DELAY_MS = 5000;
const ULONGLONG HTTP_FAST_TIMEOUT = 1000;
const ULONGLONG HTTP_NORMAL_TIMEOUT = 60000;

const int HEX_COLOR_LEN = 7;
const int LOG_BUFFER_SIZE = 512;
const int SENSOR_ID_LEN = 256;

// ============================================================================
// DATA STRUCTURES
// ============================================================================

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

struct GlobalConfig {
    Config profiles[5];
    int activeProfileIndex;
};

// ============================================================================
// MSI SDK DEFINITIONS
// ============================================================================

typedef int (*LPMLAPI_Initialize)();
typedef int (*LPMLAPI_GetDeviceInfo)(SAFEARRAY** pDevType, SAFEARRAY** pLedCount);
typedef int (*LPMLAPI_GetDeviceNameEx)(BSTR type, DWORD index, BSTR* pDevName);
typedef int (*LPMLAPI_GetLedInfo)(BSTR, DWORD, BSTR*, SAFEARRAY**);
typedef int (*LPMLAPI_SetLedColor)(BSTR type, DWORD index, DWORD R, DWORD G, DWORD B);
typedef int (*LPMLAPI_SetLedStyle)(BSTR type, DWORD index, BSTR style);
typedef int (*LPMLAPI_SetLedSpeed)(BSTR type, DWORD index, DWORD level);
typedef int (*LPMLAPI_Release)();

// COM Smart Pointers for Task Scheduler
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
// GLOBAL STATE
// ============================================================================

GlobalConfig g_Global;
// Macro for backward compatibility
#define g_cfg g_Global.profiles[g_Global.activeProfileIndex]

std::mutex g_cfgMutex;
std::mutex g_httpMutex;

std::atomic<bool> g_ResetHttp(false);
std::atomic<DataSource> g_activeSource{ DataSource::Searching };
std::atomic<bool> g_Running{ true };
std::atomic<bool> g_windows_shutdown{ false };
std::atomic<bool> g_LedsEnabled{ true };
std::atomic<float> g_asyncTemp{ -1.0f };
_bstr_t g_target_device = L"";

HANDLE g_hSensorEvent = NULL;
HANDLE g_hSourceResolvedEvent = NULL;
HANDLE g_hMutex = NULL;

HINTERNET g_hSession = NULL;
HINTERNET g_hConnect = NULL;

ULONGLONG g_ResetTimer = 0;
bool g_Resetting_sdk = false;
int g_ResetStage = 0;
ULONGLONG g_lastDataSourceSearchRetry = 0;
std::atomic<int> g_httpConnectionDrops{ 0 };

std::atomic<DWORD> lastR{ RGB_LED_REFRESH }, lastG{ RGB_LED_REFRESH }, lastB{ RGB_LED_REFRESH };
std::atomic<float> currR{ -1.0f }, currG{ -1.0f }, currB{ -1.0f };
std::atomic<float> lastTemp{ -1.0f };

BSTR g_deviceName = NULL;
HMODULE g_hLibrary = NULL;
int g_totalLeds = 0;

// SDK Function Pointers
LPMLAPI_Initialize lpMLAPI_Initialize = nullptr;
LPMLAPI_GetDeviceInfo lpMLAPI_GetDeviceInfo = nullptr;
LPMLAPI_SetLedColor lpMLAPI_SetLedColor = nullptr;
LPMLAPI_SetLedStyle lpMLAPI_SetLedStyle = nullptr;
LPMLAPI_SetLedSpeed lpMLAPI_SetLedSpeed = nullptr;
LPMLAPI_Release lpMLAPI_Release = nullptr;
LPMLAPI_GetDeviceNameEx lpMLAPI_GetDeviceNameEx = nullptr;
LPMLAPI_GetLedInfo lpMLAPI_GetLedInfo = nullptr;

// ============================================================================
// FORWARD DECLARATIONS
// ============================================================================
static void LoadProfileToUI(HWND hDlg, int profileIndex);
static void SaveUIToProfile(HWND hDlg, int profileIndex);
void PopulateAreaList(HWND hDlg, const wchar_t* deviceType);
void PopulateSensorList(HWND hDlg, const wchar_t* webServerUrl = nullptr, const wchar_t* currentTargetID = nullptr);
void SwitchActiveProfile(HWND hWnd, int index);
static bool AutoSelectFirstSensor();

// ============================================================================
// UTILITY HELPERS
// ============================================================================

static void Log(const char* text) {
    static std::mutex logMutex;
    std::lock_guard<std::mutex> lock(logMutex);
    time_t now = time(0);
    char dt[26];
    ctime_s(dt, sizeof(dt), &now);
    dt[24] = '\0';
    std::ofstream out(LOG_FILENAME, std::ios::app);
    if (out) out << "[" << dt << "] " << text << "\n";
}

static void RunShellNonAdmin(const wchar_t* path) {
    ShellExecuteW(NULL, L"open", L"explorer.exe", path, NULL, SW_SHOWNORMAL);
}

static void KillProcessByName(const wchar_t* filename) {
    HANDLE hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnap == INVALID_HANDLE_VALUE) return;
    PROCESSENTRY32W pe; pe.dwSize = sizeof(PROCESSENTRY32W);
    if (Process32FirstW(hSnap, &pe)) {
        do {
            if (_wcsicmp(pe.szExeFile, filename) == 0) {
                HANDLE hProc = OpenProcess(PROCESS_TERMINATE, FALSE, pe.th32ProcessID);
                if (hProc) { TerminateProcess(hProc, 1); CloseHandle(hProc); }
            }
        } while (Process32NextW(hSnap, &pe));
    }
    CloseHandle(hSnap);
}

static bool IsValidHex(const wchar_t* hex) {
    if (!hex || wcslen(hex) != HEX_COLOR_LEN || hex[0] != L'#') return false;
    for (int i = 1; i < HEX_COLOR_LEN; i++) if (!iswxdigit(hex[i])) return false;
    return true;
}

static COLORREF HexToColor(const wchar_t* hex) {
    if (hex[0] == L'#') hex++;
    unsigned int r = 0, g = 0, b = 0;
    if (swscanf_s(hex, L"%02x%02x%02x", &r, &g, &b) == 3) return RGB(r, g, b);
    return RGB(0, 0, 0);
}

static void ColorToHex(COLORREF color, wchar_t* out, size_t size) {
    swprintf_s(out, size, L"#%02X%02X%02X", GetRValue(color), GetGValue(color), GetBValue(color));
}

static bool IsColorDark(COLORREF col) {
    double brightness = (GetRValue(col) * 299 + GetGValue(col) * 587 + GetBValue(col) * 114) / 1000.0;
    return brightness < 128.0;
}

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

static std::wstring ExtractJsonString(const std::string& block, const std::string& key) {
    std::string searchKey = "\"" + key + "\":\"";
    size_t start = block.find(searchKey);
    if (start == std::string::npos) return L"";
    start += searchKey.length();
    size_t end = block.find("\"", start);
    if (end == std::string::npos) return L"";
    std::string val = block.substr(start, end - start);
    int len = MultiByteToWideChar(CP_UTF8, 0, val.c_str(), -1, NULL, 0);
    if (len > 0) {
        std::vector<wchar_t> buf(len);
        MultiByteToWideChar(CP_UTF8, 0, val.c_str(), -1, &buf[0], len);
        return std::wstring(&buf[0]);
    }
    return L"";
}

static int GetIntFromSafeArray(void* pRawData, VARTYPE vt, int index) {
    if (!pRawData) return 0;
    switch (vt) {
    case VT_BSTR: { BSTR* pBstrArray = (BSTR*)pRawData; return _wtoi(pBstrArray[index]); }
    case VT_I4: case VT_INT: { long* pLongArray = (long*)pRawData; return (int)pLongArray[index]; }
    case VT_UI4: case VT_UINT: { unsigned int* pULongArray = (unsigned int*)pRawData; return (int)pULongArray[index]; }
    case VT_I2: { short* pShortArray = (short*)pRawData; return (int)pShortArray[index]; }
    default: return 0;
    }
}

static void TrimLogFile() {
    WIN32_FILE_ATTRIBUTE_DATA fad;
    if (!GetFileAttributesExW(LOG_FILENAME, GetFileExInfoStandard, &fad)) return;
    LARGE_INTEGER size; size.LowPart = fad.nFileSizeLow; size.HighPart = fad.nFileSizeHigh;
    if (size.QuadPart < 1024 * 1024) return;
    std::vector<std::string> lines; std::string line;
    { std::ifstream in(LOG_FILENAME); if (!in.is_open()) return; while (std::getline(in, line)) lines.push_back(line); }
    if (lines.size() > 500) {
        std::ofstream out(LOG_FILENAME, std::ios::trunc);
        if (out.is_open()) for (size_t i = lines.size() - 500; i < lines.size(); ++i) out << lines[i] << "\n";
    }
}

static void forceLEDRefresh() {
    lastTemp = -1.0f; currR = -1.0f; currG = -1.0f; currB = -1.0f;
    lastR = RGB_LED_REFRESH; lastG = RGB_LED_REFRESH; lastB = RGB_LED_REFRESH;
}

// ============================================================================
// RESOURCE EXTRACTION (Restored)
// ============================================================================

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

// ============================================================================
// HTTP & NETWORK LOGIC
// ============================================================================

static std::string FetchLHMJson(const wchar_t* serverUrl, int timeout) {
    std::lock_guard<std::mutex> lock(g_httpMutex);

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
    if (!hRequest) { WinHttpCloseHandle(g_hConnect); g_hConnect = NULL; return ""; }

    if (WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
        WinHttpReceiveResponse(hRequest, NULL)) {
        DWORD dwSize = 0, dwDownloaded = 0;
        do {
            dwSize = 0;
            if (!WinHttpQueryDataAvailable(hRequest, &dwSize)) { connectionFailed = true; break; }
            if (dwSize == 0) break;
            char chunk[4096];
            DWORD toRead = (dwSize > sizeof(chunk)) ? sizeof(chunk) : dwSize;
            if (WinHttpReadData(hRequest, (LPVOID)chunk, toRead, &dwDownloaded)) responseData.append(chunk, dwDownloaded);
            else { connectionFailed = true; break; }
        } while (dwSize > 0);
    }
    else { connectionFailed = true; }

    WinHttpCloseHandle(hRequest);

    if (connectionFailed) {
        Log("[MysticFight] HTTP Connection dropped. Resetting handle...");
        if (g_hConnect) { WinHttpCloseHandle(g_hConnect); g_hConnect = NULL; }
        return "";
    }
    return responseData;
}

static float ParseLHMJsonForTemp(const std::string& json, const wchar_t* sensorIDW) {
    if (json.empty() || !sensorIDW) return -1.0f;
    char sensorID[256];
    size_t converted;
    if (wcstombs_s(&converted, sensorID, sizeof(sensorID), sensorIDW, _TRUNCATE) != 0) return -1.0f;

    size_t idPos = json.find(sensorID);
    if (idPos == std::string::npos) return -1.0f;
    size_t objStart = json.rfind('{', idPos);
    size_t objEnd = json.find('}', idPos);
    if (objStart == std::string::npos || objEnd == std::string::npos || objStart > idPos) return -1.0f;

    std::string objBlock = json.substr(objStart, objEnd - objStart + 1);
    std::string valueKey = "\"Value\":";
    size_t valKeyPos = objBlock.find(valueKey);
    if (valKeyPos == std::string::npos) return -1.0f;
    size_t searchOffset = valKeyPos + valueKey.length();
    size_t dataStart = objBlock.find_first_of("0123456789-.,", searchOffset);
    if (dataStart == std::string::npos) return -1.0f;
    size_t dataEnd = objBlock.find_first_of("\",}", dataStart);
    if (dataEnd == std::string::npos) dataEnd = objBlock.length();

    std::string valStr = objBlock.substr(dataStart, dataEnd - dataStart);
    std::replace(valStr.begin(), valStr.end(), ',', '.');
    try { return std::stof(valStr); }
    catch (...) { return -1.0f; }
}

static void LogAllLHMTemperatureSensors() {
    Log("[MysticFight] --- DUMPING ALL LHM TEMPERATURE SENSORS ---");
    wchar_t localURL[256];
    {
        std::lock_guard<std::mutex> lock(g_cfgMutex);
        memcpy(localURL, g_cfg.webServerUrl, sizeof(g_cfg.webServerUrl));
    }
    std::string json = FetchLHMJson(localURL, HTTP_FAST_TIMEOUT);
    if (!json.empty()) {
        std::string typeKey = "\"Type\":\"Temperature\"";
        size_t pos = 0; bool foundAny = false;
        while ((pos = json.find(typeKey, pos)) != std::string::npos) {
            size_t blockStart = json.rfind('{', pos); size_t blockEnd = json.find('}', pos);
            if (blockStart != std::string::npos && blockEnd != std::string::npos) {
                std::string block = json.substr(blockStart, blockEnd - blockStart + 1);
                std::wstring name = ExtractJsonString(block, "Text");
                std::wstring id = ExtractJsonString(block, "SensorId");
                std::wstring valStr = ExtractJsonString(block, "Value");
                if (!id.empty()) {
                    std::wstring wVal = valStr; for (auto& c : wVal) if (c == L',') c = L'.';
                    float valFloat = (float)_wtof(wVal.c_str());
                    char buf[LOG_BUFFER_SIZE]; wchar_t wBuf[LOG_BUFFER_SIZE];
                    swprintf(wBuf, LOG_BUFFER_SIZE, L"[MysticFight] [HTTP] ID: %ls | Name: %ls | Value: %.1f", id.c_str(), name.c_str(), valFloat);
                    WideCharToMultiByte(CP_ACP, 0, wBuf, -1, buf, sizeof(buf), NULL, NULL);
                    Log(buf); foundAny = true;
                }
            }
            pos = blockEnd;
        }
        if (foundAny) Log("[MysticFight] --- END OF HTTP DUMP ---");
    }
}

// ============================================================================
// SENSOR THREAD LOGIC (Restored)
// ============================================================================

static float GetCPUTempFast() {
    ULONGLONG currentTime = GetTickCount64();
    float temp = -1.0f;
    wchar_t localSensorID[256]; wchar_t localURL[256];
    {
        std::lock_guard<std::mutex> lock(g_cfgMutex);
        memcpy(localSensorID, g_cfg.sensorID, sizeof(g_cfg.sensorID));
        memcpy(localURL, g_cfg.webServerUrl, sizeof(g_cfg.webServerUrl));
    }

    if (g_activeSource == DataSource::HTTP) {
        std::string json = FetchLHMJson(localURL, HTTP_NORMAL_TIMEOUT);
        if (json.empty()) {
            g_httpConnectionDrops++;
            if (g_httpConnectionDrops >= 2) {
                Log("[MysticFight] HTTP connection lost for 2 consecutive polls. Resetting to SEARCH mode.");
                g_activeSource = DataSource::Searching; g_httpConnectionDrops = 0;
            }
            return -1.0f;
        }
        g_httpConnectionDrops = 0;
        temp = ParseLHMJsonForTemp(json, localSensorID);
        return temp;
    }
    else {
        if (currentTime - g_lastDataSourceSearchRetry <= LHM_RETRY_DELAY_MS) return -1.0f;
        g_lastDataSourceSearchRetry = currentTime;
        Log("[MysticFight] Attempting HTTP data source discovery...");
        std::string json = FetchLHMJson(localURL, HTTP_NORMAL_TIMEOUT);
        if (!json.empty()) {
            g_activeSource = DataSource::HTTP;
            Log("[MysticFight] Data Source detected: HTTP.");
            SetEvent(g_hSourceResolvedEvent);
            temp = ParseLHMJsonForTemp(json, localSensorID);
            if (temp >= 0.0f) { LogAllLHMTemperatureSensors(); return temp; }
            return -1.0f;
        }
    }
    return -1.0f;
}

static void SensorThread() {
    while (g_Running) {
        if (g_LedsEnabled) {
            g_asyncTemp = GetCPUTempFast();
        }
        DWORD waitMs = INFINITE;
        if (g_LedsEnabled) {
            std::lock_guard<std::mutex> lock(g_cfgMutex);
            waitMs = (DWORD)g_cfg.sensorUpdateMS;
        }
        WaitForSingleObject(g_hSensorEvent, waitMs);
    }
}

// ============================================================================
// HARDWARE DETECTION (Restored)
// ============================================================================

static void MSIHwardwareDetection() {
    SAFEARRAY* pDevType = nullptr; SAFEARRAY* pLedCount = nullptr;
    if (lpMLAPI_GetDeviceInfo && lpMLAPI_GetDeviceInfo(&pDevType, &pLedCount) == 0 && pDevType && pLedCount) {
        BSTR* pTypes = nullptr; void* pCountsRaw = nullptr;
        VARTYPE vtCount; SafeArrayGetVartype(pLedCount, &vtCount);
        if (SUCCEEDED(SafeArrayAccessData(pDevType, (void**)&pTypes)) && SUCCEEDED(SafeArrayAccessData(pLedCount, &pCountsRaw))) {
            long lBound, uBound; SafeArrayGetLBound(pDevType, 1, &lBound); SafeArrayGetUBound(pDevType, 1, &uBound);
            if ((uBound - lBound + 1) > 0 && pTypes != nullptr && pTypes[0] != nullptr) {
                if (g_deviceName) SysFreeString(g_deviceName);
                g_deviceName = SysAllocString(pTypes[0]);
                g_totalLeds = GetIntFromSafeArray(pCountsRaw, vtCount, 0);
                BSTR friendlyName = NULL;
                if (lpMLAPI_GetDeviceNameEx) lpMLAPI_GetDeviceNameEx(g_deviceName, 0, &friendlyName);
                char devInfo[LOG_BUFFER_SIZE];
                snprintf(devInfo, sizeof(devInfo), "[MysticFight] %ls (Type: %ls) | Logical Areas: %d", (friendlyName ? friendlyName : L"Unknown Device"), g_deviceName, g_totalLeds);
                Log(devInfo);
                if (friendlyName) SysFreeString(friendlyName);

                Log("[MysticFight] Listing styles per area...");
                for (DWORD i = 0; i < (DWORD)g_totalLeds; i++) {
                    BSTR ledName = nullptr; SAFEARRAY* pStyles = nullptr;
                    if (lpMLAPI_GetLedInfo(g_deviceName, i, &ledName, &pStyles) == 0) {
                        char ledLine[LOG_BUFFER_SIZE]; snprintf(ledLine, sizeof(ledLine), "    [INDEX %lu] LED: %ls", i, (ledName ? ledName : L"Unknown"));
                        Log(ledLine);
                        if (pStyles) SafeArrayDestroy(pStyles);
                        if (ledName) SysFreeString(ledName);
                    }
                }
            }
            SafeArrayUnaccessData(pDevType); SafeArrayUnaccessData(pLedCount);
        }
        SafeArrayDestroy(pDevType); SafeArrayDestroy(pLedCount);
    }
}

// ============================================================================
// TASK SCHEDULER & CONFIG LOGIC
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
        if (start) { IRunningTaskPtr pRunningTask; return pTask->Run(_variant_t(), &pRunningTask); }
        else { return pTask->Stop(0); }
    }
    catch (const _com_error& e) { return e.Error(); }
}

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
    catch (const _com_error&) { return false; }
    return false;
}

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
        if (SUCCEEDED(hr)) Log("[MysticFight] Startup task removed");
        if (!run) return;
        ITaskDefinitionPtr pTask; pService->NewTask(0, &pTask);
        IPrincipalPtr pPrincipal; pTask->get_Principal(&pPrincipal);
        if (pPrincipal != nullptr) { pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST); pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN); }
        ITaskSettingsPtr pSettings; pTask->get_Settings(&pSettings);
        if (pSettings != nullptr) {
            pSettings->put_StartWhenAvailable(VARIANT_TRUE); pSettings->put_DisallowStartIfOnBatteries(VARIANT_FALSE);
            pSettings->put_StopIfGoingOnBatteries(VARIANT_FALSE); pSettings->put_ExecutionTimeLimit(_bstr_t(L"PT0S"));
            pSettings->put_AllowHardTerminate(VARIANT_TRUE);
        }
        ITriggerCollectionPtr pTriggers; pTask->get_Triggers(&pTriggers);
        ITriggerPtr pTrigger; pTriggers->Create(TASK_TRIGGER_LOGON, &pTrigger);
        ILogonTriggerPtr pLogonTrigger = pTrigger;
        if (pLogonTrigger != nullptr) { pLogonTrigger->put_Delay(_bstr_t(L"PT30S")); }
        IActionCollectionPtr pActions; pTask->get_Actions(&pActions);
        IActionPtr pAction; pActions->Create(TASK_ACTION_EXEC, &pAction);
        IExecActionPtr pExecAction = pAction;
        if (pExecAction != nullptr) { pExecAction->put_Path(_bstr_t(szPath)); pExecAction->put_WorkingDirectory(_bstr_t(szDir)); }
        IRegisteredTaskPtr pRegisteredTask;
        pRootFolder->RegisterTaskDefinition(_bstr_t(TASK_NAME), pTask, TASK_CREATE_OR_UPDATE, _variant_t(), _variant_t(), TASK_LOGON_INTERACTIVE_TOKEN, _variant_t(L""), &pRegisteredTask);
        Log("[MysticFight] New Startup task created");
    }
    catch (const _com_error& e) {
        char buffer[256]; snprintf(buffer, sizeof(buffer), "[MysticFight] TaskScheduler ERROR: HRESULT 0x%08X", (unsigned int)e.Error()); Log(buffer);
    }
}

void SaveSettings() {
    std::lock_guard<std::mutex> lock(g_cfgMutex);
    WritePrivateProfileStringW(L"Global", L"ActiveProfile", std::to_wstring(g_Global.activeProfileIndex).c_str(), INI_FILE);
    for (int i = 0; i < 5; i++) {
        std::wstring section = L"Settings"; if (i > 0) section += L"_" + std::to_wstring(i);
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
        wchar_t hL[10], hM[10], hH[10]; ColorToHex(p.colorLow, hL, 10); ColorToHex(p.colorMed, hM, 10); ColorToHex(p.colorHigh, hH, 10);
        WritePrivateProfileStringW(section.c_str(), L"ColorLow", hL, INI_FILE);
        WritePrivateProfileStringW(section.c_str(), L"ColorMed", hM, INI_FILE);
        WritePrivateProfileStringW(section.c_str(), L"ColorHigh", hH, INI_FILE);
        WritePrivateProfileStringW(section.c_str(), L"WebServerUrl", p.webServerUrl, INI_FILE);
    }
}

static bool AutoSelectFirstSensor() {
    wchar_t localURL[256];
    {
        std::lock_guard<std::mutex> lock(g_cfgMutex);
        memcpy(localURL, g_cfg.webServerUrl, sizeof(g_cfg.webServerUrl));
    }
    std::string json = FetchLHMJson(localURL, HTTP_FAST_TIMEOUT);
    if (json.empty()) { Log("[MysticFight] Auto-select failed: HTTP connection unavailable."); return false; }
    std::string typeKey = "\"Type\":\"Temperature\"";
    size_t pos = json.find(typeKey);
    if (pos != std::string::npos) {
        size_t blockStart = json.rfind('{', pos);
        size_t blockEnd = json.find('}', pos);
        if (blockStart != std::string::npos && blockEnd != std::string::npos) {
            std::string block = json.substr(blockStart, blockEnd - blockStart + 1);
            std::wstring id = ExtractJsonString(block, "SensorId");
            if (!id.empty()) {
                std::lock_guard<std::mutex> lock(g_cfgMutex);
                wcscpy_s(g_cfg.sensorID, id.c_str());
                WritePrivateProfileStringW(L"Settings", L"SensorID", g_cfg.sensorID, INI_FILE);
                char logBuf[LOG_BUFFER_SIZE]; snprintf(logBuf, sizeof(logBuf), "[MysticFight] Auto-selected default sensor (HTTP): %ls", g_cfg.sensorID);
                Log(logBuf);
                return true;
            }
        }
    }
    Log("[MysticFight] Auto-select failed: No temperature sensors found in HTTP JSON.");
    return false;
}

void LoadSettings() {
    g_Global.activeProfileIndex = GetPrivateProfileIntW(L"Global", L"ActiveProfile", 0, INI_FILE);
    if (g_Global.activeProfileIndex < 0 || g_Global.activeProfileIndex > 4) g_Global.activeProfileIndex = 0;

    for (int i = 0; i < 5; i++) {
        std::wstring section = L"Settings"; if (i > 0) section += L"_" + std::to_wstring(i);
        Config& p = g_Global.profiles[i];
        p.tempLow = GetPrivateProfileIntW(section.c_str(), L"TempLow", 50, INI_FILE);
        p.tempMed = GetPrivateProfileIntW(section.c_str(), L"TempMed", 70, INI_FILE);
        p.tempHigh = GetPrivateProfileIntW(section.c_str(), L"TempHigh", 90, INI_FILE);
        GetPrivateProfileStringW(section.c_str(), L"TargetDevice", L"MSI_MB", p.targetDevice, 256, INI_FILE);
        p.targetLedIndex = GetPrivateProfileIntW(section.c_str(), L"TargetLedIndex", 0, INI_FILE);
        wchar_t hL[10], hM[10], hH[10];
        GetPrivateProfileStringW(section.c_str(), L"ColorLow", L"#00FF00", hL, 10, INI_FILE);
        GetPrivateProfileStringW(section.c_str(), L"ColorMed", L"#FFFF00", hM, 10, INI_FILE);
        GetPrivateProfileStringW(section.c_str(), L"ColorHigh", L"#FF0000", hH, 10, INI_FILE);
        p.colorLow = HexToColor(hL); p.colorMed = HexToColor(hM); p.colorHigh = HexToColor(hH);
        GetPrivateProfileStringW(section.c_str(), L"SensorID", L"", p.sensorID, SENSOR_ID_LEN, INI_FILE);
        GetPrivateProfileStringW(section.c_str(), L"WebServerUrl", L"http://localhost:8085", p.webServerUrl, 256, INI_FILE);
        wchar_t defLabel[64]; swprintf_s(defLabel, L"Profile %d", i + 1);
        GetPrivateProfileStringW(section.c_str(), L"Label", defLabel, p.label, 64, INI_FILE);
        if (wcslen(p.label) == 0) wcscpy_s(p.label, defLabel);
        p.sensorUpdateMS = GetPrivateProfileIntW(section.c_str(), L"SensorUpdateMS", 500, INI_FILE);
        p.ledRefreshFPS = GetPrivateProfileIntW(section.c_str(), L"LedRefreshFPS", 25, INI_FILE);
        wchar_t sfBuf[32]; GetPrivateProfileStringW(section.c_str(), L"SmoothingFactor", L"0.150", sfBuf, 32, INI_FILE);
        p.smoothingFactor = (float)_wtof(sfBuf);
        if (p.colorLow == 0 && p.colorMed == 0 && p.colorHigh == 0) {
            p.colorLow = RGB(0, 255, 0); p.colorMed = RGB(255, 255, 0); p.colorHigh = RGB(255, 0, 0);
        }
        bool invalid = false;
        if (p.tempLow < 0 || p.tempHigh > 110 || p.tempLow >= p.tempMed || p.tempMed >= p.tempHigh || wcslen(p.targetDevice) == 0 || p.targetLedIndex < 0 || p.targetLedIndex > 255) invalid = true;
        if (invalid) {
            p.tempLow = 50; p.tempMed = 70; p.tempHigh = 90;
            p.colorLow = RGB(0, 255, 0); p.colorMed = RGB(255, 255, 0); p.colorHigh = RGB(255, 0, 0);
            wcscpy_s(p.targetDevice, L"MSI_MB"); p.targetLedIndex = 0; p.sensorID[0] = L'\0';
        }
    }
    if (wcslen(g_cfg.sensorID) == 0) AutoSelectFirstSensor();
}

static void LoadProfileToUI(HWND hDlg, int profileIndex) {
    Config& p = g_Global.profiles[profileIndex];
    SetDlgItemInt(hDlg, IDC_TEMP_LOW, p.tempLow, TRUE);
    SetDlgItemInt(hDlg, IDC_TEMP_MED, p.tempMed, TRUE);
    SetDlgItemInt(hDlg, IDC_TEMP_HIGH, p.tempHigh, TRUE);
    wchar_t hexBuf[10];
    ColorToHex(p.colorLow, hexBuf, 10); SetDlgItemTextW(hDlg, IDC_HEX_LOW, hexBuf);
    ColorToHex(p.colorMed, hexBuf, 10); SetDlgItemTextW(hDlg, IDC_HEX_MED, hexBuf);
    ColorToHex(p.colorHigh, hexBuf, 10); SetDlgItemTextW(hDlg, IDC_HEX_HIGH, hexBuf);
    SetDlgItemTextW(hDlg, IDC_EDIT_LABEL, p.label);
    HWND hSensorUpdate = GetDlgItem(hDlg, IDC_COMBO_SENSOR_UPDATE);
    if (p.sensorUpdateMS == 250) SendMessage(hSensorUpdate, CB_SETCURSEL, 0, 0);
    else if (p.sensorUpdateMS == 1000) SendMessage(hSensorUpdate, CB_SETCURSEL, 2, 0);
    else SendMessage(hSensorUpdate, CB_SETCURSEL, 1, 0);
    HWND hLedFPS = GetDlgItem(hDlg, IDC_COMBO_LED_FPS);
    if (p.ledRefreshFPS == 20) SendMessage(hLedFPS, CB_SETCURSEL, 1, 0);
    else if (p.ledRefreshFPS == 15) SendMessage(hLedFPS, CB_SETCURSEL, 2, 0);
    else SendMessage(hLedFPS, CB_SETCURSEL, 0, 0);
    HWND hSmoothing = GetDlgItem(hDlg, IDC_COMBO_SMOOTHING);
    if (p.smoothingFactor > 0.9f) SendMessage(hSmoothing, CB_SETCURSEL, 0, 0);
    else if (p.smoothingFactor < 0.1f) SendMessage(hSmoothing, CB_SETCURSEL, 1, 0);
    else if (p.smoothingFactor > 0.2f) SendMessage(hSmoothing, CB_SETCURSEL, 3, 0);
    else SendMessage(hSmoothing, CB_SETCURSEL, 2, 0);
    SetDlgItemTextW(hDlg, IDC_EDIT_SERVER, p.webServerUrl);
    int devCount = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_DEVICE), CB_GETCOUNT, 0, 0);
    for (int i = 0; i < devCount; i++) {
        wchar_t* devName = (wchar_t*)SendMessage(GetDlgItem(hDlg, IDC_COMBO_DEVICE), CB_GETITEMDATA, i, 0);
        if (devName && wcscmp(devName, p.targetDevice) == 0) { SendMessage(GetDlgItem(hDlg, IDC_COMBO_DEVICE), CB_SETCURSEL, i, 0); break; }
    }
    PopulateAreaList(hDlg, p.targetDevice);
    int areaCount = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_AREA), CB_GETCOUNT, 0, 0);
    for (int i = 0; i < areaCount; i++) {
        int idx = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_AREA), CB_GETITEMDATA, i, 0);
        if (idx == p.targetLedIndex) { SendMessage(GetDlgItem(hDlg, IDC_COMBO_AREA), CB_SETCURSEL, i, 0); break; }
    }
    PopulateSensorList(hDlg, p.webServerUrl, p.sensorID);
    int sensCount = (int)SendMessage(GetDlgItem(hDlg, IDC_SENSOR_ID), CB_GETCOUNT, 0, 0);
    bool foundSens = false;
    for (int i = 0; i < sensCount; i++) {
        wchar_t* sID = (wchar_t*)SendMessage(GetDlgItem(hDlg, IDC_SENSOR_ID), CB_GETITEMDATA, i, 0);
        if (sID && wcscmp(sID, p.sensorID) == 0) { SendMessage(GetDlgItem(hDlg, IDC_SENSOR_ID), CB_SETCURSEL, i, 0); foundSens = true; break; }
    }
    if (!foundSens && sensCount > 0) SendMessage(GetDlgItem(hDlg, IDC_SENSOR_ID), CB_SETCURSEL, 0, 0);
    CheckDlgButton(hDlg, IDC_CHK_ACTIVE_PROFILE, (profileIndex == g_Global.activeProfileIndex) ? BST_CHECKED : BST_UNCHECKED);
    InvalidateRect(GetDlgItem(hDlg, IDC_HEX_LOW), NULL, TRUE);
    InvalidateRect(GetDlgItem(hDlg, IDC_HEX_MED), NULL, TRUE);
    InvalidateRect(GetDlgItem(hDlg, IDC_HEX_HIGH), NULL, TRUE);
}

static void SaveUIToProfile(HWND hDlg, int profileIndex) {
    Config& p = g_Global.profiles[profileIndex];
    p.tempLow = GetDlgItemInt(hDlg, IDC_TEMP_LOW, NULL, TRUE);
    p.tempMed = GetDlgItemInt(hDlg, IDC_TEMP_MED, NULL, TRUE);
    p.tempHigh = GetDlgItemInt(hDlg, IDC_TEMP_HIGH, NULL, TRUE);
    wchar_t buf[256];
    GetDlgItemTextW(hDlg, IDC_HEX_LOW, buf, 10); p.colorLow = HexToColor(buf);
    GetDlgItemTextW(hDlg, IDC_HEX_MED, buf, 10); p.colorMed = HexToColor(buf);
    GetDlgItemTextW(hDlg, IDC_HEX_HIGH, buf, 10); p.colorHigh = HexToColor(buf);
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
    if (idxSensor == 0) p.sensorUpdateMS = 250; else if (idxSensor == 2) p.sensorUpdateMS = 1000; else p.sensorUpdateMS = 500;
    int idxFPS = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_LED_FPS), CB_GETCURSEL, 0, 0);
    if (idxFPS == 1) p.ledRefreshFPS = 20; else if (idxFPS == 2) p.ledRefreshFPS = 15; else p.ledRefreshFPS = 25;
    int idxSmooth = (int)SendMessage(GetDlgItem(hDlg, IDC_COMBO_SMOOTHING), CB_GETCURSEL, 0, 0);
    if (idxSmooth == 0) p.smoothingFactor = 1.0f; else if (idxSmooth == 1) p.smoothingFactor = 0.05f; else if (idxSmooth == 3) p.smoothingFactor = 0.4f; else p.smoothingFactor = 0.15f;
}

static void ShowNotification(HWND hWnd, HINSTANCE hInstance, const wchar_t* title, const wchar_t* info) {
    NOTIFYICONDATAW nid = { sizeof(NOTIFYICONDATAW), hWnd, 1, NIF_INFO };
    wcscpy_s(nid.szInfoTitle, _countof(nid.szInfoTitle), title);
    wcscpy_s(nid.szInfo, _countof(nid.szInfo), info);
    nid.dwInfoFlags = NIIF_USER | NIIF_LARGE_ICON;
    Shell_NotifyIconW(NIM_MODIFY, &nid);
}

void PopulateAreaList(HWND hDlg, const wchar_t* deviceType) {
    HWND hComboArea = GetDlgItem(hDlg, IDC_COMBO_AREA);
    SendMessage(hComboArea, CB_RESETCONTENT, 0, 0);
    if (!deviceType || !lpMLAPI_GetLedInfo || !lpMLAPI_GetDeviceInfo) { EnableWindow(hComboArea, FALSE); return; }

    SAFEARRAY* pDevType = nullptr; SAFEARRAY* pLedCount = nullptr;
    int currentDeviceLedCount = 0;

    if (lpMLAPI_GetDeviceInfo(&pDevType, &pLedCount) == 0 && pDevType && pLedCount) {
        BSTR* pTypes = nullptr; void* pCountsRaw = nullptr;
        VARTYPE vtCount; SafeArrayGetVartype(pLedCount, &vtCount);
        if (SUCCEEDED(SafeArrayAccessData(pDevType, (void**)&pTypes)) && SUCCEEDED(SafeArrayAccessData(pLedCount, &pCountsRaw))) {
            long lBound, uBound; SafeArrayGetLBound(pDevType, 1, &lBound); SafeArrayGetUBound(pDevType, 1, &uBound);
            long totalDevices = uBound - lBound + 1;
            for (long j = 0; j < totalDevices; j++) {
                if (wcscmp(pTypes[j], deviceType) == 0) { currentDeviceLedCount = GetIntFromSafeArray(pCountsRaw, vtCount, j); break; }
            }
            SafeArrayUnaccessData(pDevType); SafeArrayUnaccessData(pLedCount);
        }
        SafeArrayDestroy(pDevType); SafeArrayDestroy(pLedCount);

        if (currentDeviceLedCount <= 0) currentDeviceLedCount = 0;

        // --- CRITICAL FIX: ALLOCATE BSTR ---
        BSTR bstrDevice = SysAllocString(deviceType);
        if (bstrDevice) {
            for (DWORD i = 0; i < (DWORD)currentDeviceLedCount; i++) {
                BSTR ledName = nullptr; SAFEARRAY* pStyles = nullptr;
                if (lpMLAPI_GetLedInfo(bstrDevice, i, &ledName, &pStyles) == 0) {
                    int idx = (int)SendMessageW(hComboArea, CB_ADDSTRING, 0, (LPARAM)(ledName ? ledName : L"Unknown Area"));
                    SendMessage(hComboArea, CB_SETITEMDATA, idx, (LPARAM)i);
                    if (i == (DWORD)g_cfg.targetLedIndex) SendMessage(hComboArea, CB_SETCURSEL, idx, 0);
                    if (pStyles) SafeArrayDestroy(pStyles);
                    if (ledName) SysFreeString(ledName);
                }
            }
            SysFreeString(bstrDevice);
        }
    }

    int finalCount = (int)SendMessage(hComboArea, CB_GETCOUNT, 0, 0);
    if (finalCount <= 1) {
        if (finalCount == 1) SendMessage(hComboArea, CB_SETCURSEL, 0, 0);
        EnableWindow(hComboArea, FALSE);
    }
    else {
        if (SendMessage(hComboArea, CB_GETCURSEL, 0, 0) == CB_ERR) SendMessage(hComboArea, CB_SETCURSEL, 0, 0);
        EnableWindow(hComboArea, TRUE);
    }
}

void PopulateDeviceList(HWND hDlg) {
    HWND hComboDev = GetDlgItem(hDlg, IDC_COMBO_DEVICE);
    ClearComboHeapData(hComboDev);
    SAFEARRAY* pDevType = nullptr; SAFEARRAY* pLedCount = nullptr;

    if (lpMLAPI_GetDeviceInfo && lpMLAPI_GetDeviceInfo(&pDevType, &pLedCount) == 0) {
        BSTR* pTypes = nullptr; SafeArrayAccessData(pDevType, (void**)&pTypes);
        long lBound, uBound; SafeArrayGetLBound(pDevType, 1, &lBound); SafeArrayGetUBound(pDevType, 1, &uBound);
        long count = uBound - lBound + 1;
        for (long i = 0; i < count; i++) {
            BSTR typeInternal = pTypes[i];
            BSTR friendlyName = NULL;
            if (lpMLAPI_GetDeviceNameEx) lpMLAPI_GetDeviceNameEx(typeInternal, i, &friendlyName);
            wchar_t* pSafeStr = HeapDupString(typeInternal);
            int idx = (int)SendMessageW(hComboDev, CB_ADDSTRING, 0, (LPARAM)(friendlyName ? friendlyName : typeInternal));
            if (idx != CB_ERR) {
                void* oldData = (void*)SendMessage(hComboDev, CB_GETITEMDATA, idx, 0);
                if (oldData && oldData != (void*)CB_ERR) HeapFree(GetProcessHeap(), 0, oldData);
                SendMessage(hComboDev, CB_SETITEMDATA, idx, (LPARAM)pSafeStr);
                if (wcscmp(typeInternal, g_cfg.targetDevice) == 0) SendMessage(hComboDev, CB_SETCURSEL, idx, 0);
            }
            else { HeapFree(GetProcessHeap(), 0, pSafeStr); }
            if (friendlyName) SysFreeString(friendlyName);
        }
        SafeArrayUnaccessData(pDevType); SafeArrayDestroy(pDevType); SafeArrayDestroy(pLedCount);
        if (SendMessage(hComboDev, CB_GETCURSEL, 0, 0) == CB_ERR && SendMessage(hComboDev, CB_GETCOUNT, 0, 0) > 0) SendMessage(hComboDev, CB_SETCURSEL, 0, 0);
        int finalCount = (int)SendMessage(hComboDev, CB_GETCOUNT, 0, 0);
        if (finalCount <= 1) { if (finalCount == 1) SendMessage(hComboDev, CB_SETCURSEL, 0, 0); EnableWindow(hComboDev, FALSE); }
        else { if (SendMessage(hComboDev, CB_GETCURSEL, 0, 0) == CB_ERR) SendMessage(hComboDev, CB_SETCURSEL, 0, 0); EnableWindow(hComboDev, TRUE); }
    }
    else { EnableWindow(hComboDev, FALSE); Log("[MysticFight] SDK Error or No MSI Devices found."); }
}

void PopulateSensorList(HWND hDlg, const wchar_t* webServerUrl, const wchar_t* currentTargetID) {
    HWND hCombo = GetDlgItem(hDlg, IDC_SENSOR_ID);
    ClearComboHeapData(hCombo);
    wchar_t localTargetID[256];
    if (currentTargetID) wcscpy_s(localTargetID, currentTargetID);
    else { std::lock_guard<std::mutex> lock(g_cfgMutex); wcscpy_s(localTargetID, g_cfg.sensorID); }
    wchar_t localURL[256];
    if (webServerUrl) wcscpy_s(localURL, webServerUrl);
    else { std::lock_guard<std::mutex> lock(g_cfgMutex); wcscpy_s(localURL, g_cfg.webServerUrl); }
    std::string json = FetchLHMJson(localURL, HTTP_FAST_TIMEOUT);
    if (json.empty()) return;
    std::string typeKey = "\"Type\":\"Temperature\""; size_t pos = 0; bool foundAny = false;
    while ((pos = json.find(typeKey, pos)) != std::string::npos) {
        size_t blockEnd = json.find('}', pos); size_t blockStart = json.rfind('{', pos);
        if (blockEnd != std::string::npos && blockStart != std::string::npos) {
            std::string block = json.substr(blockStart, blockEnd - blockStart + 1);
            std::wstring name = ExtractJsonString(block, "Text");
            std::wstring id = ExtractJsonString(block, "SensorId");
            if (!name.empty() && !id.empty()) {
                if (SendMessage(hCombo, CB_FINDSTRINGEXACT, -1, (LPARAM)name.c_str()) == CB_ERR) {
                    wchar_t* pSafeStr = HeapDupString(id.c_str());
                    int idx = (int)SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)name.c_str());
                    if (idx != CB_ERR) {
                        SendMessage(hCombo, CB_SETITEMDATA, idx, (LPARAM)pSafeStr);
                        if (wcscmp(id.c_str(), localTargetID) == 0) { SendMessage(hCombo, CB_SETCURSEL, idx, 0); foundAny = true; }
                    }
                    else { HeapFree(GetProcessHeap(), 0, pSafeStr); }
                }
            }
        }
        pos = blockEnd;
    }
    if (!foundAny && SendMessage(hCombo, CB_GETCOUNT, 0, 0) > 0 && SendMessage(hCombo, CB_GETCURSEL, 0, 0) == CB_ERR) SendMessage(hCombo, CB_SETCURSEL, 0, 0);
    int finalCount = (int)SendMessage(hCombo, CB_GETCOUNT, 0, 0);
    if (finalCount <= 1) { if (finalCount == 1) SendMessage(hCombo, CB_SETCURSEL, 0, 0); EnableWindow(hCombo, FALSE); }
    else { if (SendMessage(hCombo, CB_GETCURSEL, 0, 0) == CB_ERR) SendMessage(hCombo, CB_SETCURSEL, 0, 0); EnableWindow(hCombo, TRUE); }
}

// ============================================================================
// WINDOW PROCEDURES
// ============================================================================

WNDPROC oldEditProc;
static COLORREF g_CustomColors[16] = { 0 };

LRESULT CALLBACK ColorEditSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    if (uMsg == WM_LBUTTONUP) {
        CHOOSECOLOR cc = { 0 }; cc.lStructSize = sizeof(cc); cc.hwndOwner = GetParent(hWnd);
        cc.lpCustColors = g_CustomColors; cc.Flags = CC_RGBINIT | CC_FULLOPEN | CC_ANYCOLOR;
        wchar_t buf[10]; GetWindowTextW(hWnd, buf, 10);
        if (IsValidHex(buf)) cc.rgbResult = HexToColor(buf);
        if (ChooseColor(&cc)) {
            wchar_t newHex[10]; ColorToHex(cc.rgbResult, newHex, 10); SetWindowTextW(hWnd, newHex);
            SendMessage(GetParent(hWnd), WM_COMMAND, MAKEWPARAM(GetDlgCtrlID(hWnd), EN_CHANGE), (LPARAM)hWnd);
        }
    }
    return CallWindowProc(oldEditProc, hWnd, uMsg, wParam, lParam);
}

INT_PTR CALLBACK SettingsDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    static HBRUSH hBrushLow = NULL, hBrushMed = NULL, hBrushHigh = NULL;
    static int s_currentTab = 0;
    switch (message) {
    case WM_CTLCOLOREDIT: {
        HDC hdc = (HDC)wParam; HWND hCtrl = (HWND)lParam; int id = GetDlgCtrlID(hCtrl); HBRUSH hSelectedBrush = NULL;
        if (id == IDC_HEX_LOW) hSelectedBrush = hBrushLow; if (id == IDC_HEX_MED) hSelectedBrush = hBrushMed; if (id == IDC_HEX_HIGH) hSelectedBrush = hBrushHigh;
        if (hSelectedBrush) {
            wchar_t buf[10]; GetWindowTextW(hCtrl, buf, 10);
            if (IsValidHex(buf)) {
                COLORREF c = HexToColor(buf); SetBkColor(hdc, c);
                if (IsColorDark(c)) SetTextColor(hdc, RGB(255, 255, 255)); else SetTextColor(hdc, RGB(0, 0, 0));
                return (INT_PTR)hSelectedBrush;
            }
        }
        return (INT_PTR)GetStockObject(WHITE_BRUSH);
    }
    case WM_INITDIALOG: {
        RECT rcDlg; GetWindowRect(hDlg, &rcDlg);
        int x = (GetSystemMetrics(SM_CXSCREEN) - (rcDlg.right - rcDlg.left)) / 2;
        int y = (GetSystemMetrics(SM_CYSCREEN) - (rcDlg.bottom - rcDlg.top)) / 2;
        SetWindowPos(hDlg, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE);
        LoadSettings();
        HWND hTab = GetDlgItem(hDlg, IDC_TAB_PROFILES);
        if (hTab) {
            TabCtrl_DeleteAllItems(hTab); TCITEMW tie = { 0 }; tie.mask = TCIF_TEXT;
            for (int i = 0; i < 5; i++) { tie.pszText = g_Global.profiles[i].label; SendMessage(hTab, TCM_INSERTITEMW, i, (LPARAM)&tie); }
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
        s_currentTab = g_Global.activeProfileIndex; TabCtrl_SetCurSel(hTab, s_currentTab);
        PopulateDeviceList(hDlg);
        LoadProfileToUI(hDlg, s_currentTab);
        hBrushLow = CreateSolidBrush(g_Global.profiles[s_currentTab].colorLow);
        hBrushMed = CreateSolidBrush(g_Global.profiles[s_currentTab].colorMed);
        hBrushHigh = CreateSolidBrush(g_Global.profiles[s_currentTab].colorHigh);
        oldEditProc = (WNDPROC)SetWindowLongPtr(GetDlgItem(hDlg, IDC_HEX_LOW), GWLP_WNDPROC, (LONG_PTR)ColorEditSubclassProc);
        SetWindowLongPtr(GetDlgItem(hDlg, IDC_HEX_MED), GWLP_WNDPROC, (LONG_PTR)ColorEditSubclassProc);
        SetWindowLongPtr(GetDlgItem(hDlg, IDC_HEX_HIGH), GWLP_WNDPROC, (LONG_PTR)ColorEditSubclassProc);
        DataSource currentSrc = g_activeSource.load();
        const wchar_t* sourceLabel = (currentSrc == DataSource::HTTP) ? L"LibreHardwareMonitor (Data source: HTTP)" : L"LibreHardwareMonitor (Searching...)";
        SetDlgItemTextW(hDlg, IDC_GRP_SOURCE, sourceLabel);
        if (ValidStartupTaskExists()) CheckDlgButton(hDlg, IDC_CHK_STARTUP, BST_CHECKED);
        return (INT_PTR)TRUE;
    }
    case WM_NOTIFY: {
        LPNMHDR pnm = (LPNMHDR)lParam;
        if (pnm->idFrom == IDC_TAB_PROFILES && pnm->code == TCN_SELCHANGE) {
            SaveUIToProfile(hDlg, s_currentTab);
            s_currentTab = TabCtrl_GetCurSel(pnm->hwndFrom);
            LoadProfileToUI(hDlg, s_currentTab);
            if (hBrushLow) DeleteObject(hBrushLow); if (hBrushMed) DeleteObject(hBrushMed); if (hBrushHigh) DeleteObject(hBrushHigh);
            hBrushLow = CreateSolidBrush(g_Global.profiles[s_currentTab].colorLow);
            hBrushMed = CreateSolidBrush(g_Global.profiles[s_currentTab].colorMed);
            hBrushHigh = CreateSolidBrush(g_Global.profiles[s_currentTab].colorHigh);
            InvalidateRect(hDlg, NULL, TRUE);
        }
        break;
    }
    case WM_COMMAND: {
        int id = LOWORD(wParam); int code = HIWORD(wParam);
        if (id == IDC_CHK_ACTIVE_PROFILE && code == BN_CLICKED) {
            if (IsDlgButtonChecked(hDlg, IDC_CHK_ACTIVE_PROFILE)) g_Global.activeProfileIndex = s_currentTab;
            else if (g_Global.activeProfileIndex == s_currentTab) CheckDlgButton(hDlg, IDC_CHK_ACTIVE_PROFILE, BST_CHECKED);
            return TRUE;
        }
        if (id == IDC_COMBO_DEVICE && code == CBN_SELCHANGE) {
            int idx = (int)SendMessage((HWND)lParam, CB_GETCURSEL, 0, 0);
            if (idx != CB_ERR) {
                wchar_t* devType = (wchar_t*)SendMessage((HWND)lParam, CB_GETITEMDATA, idx, 0);
                if (devType) PopulateAreaList(hDlg, devType);
            }
            return TRUE;
        }
        if (code == EN_CHANGE) {
            wchar_t buf[10]; GetDlgItemTextW(hDlg, id, buf, 10);
            if (IsValidHex(buf)) {
                COLORREF newCol = HexToColor(buf); HBRUSH newBrush = CreateSolidBrush(newCol);
                if (id == IDC_HEX_LOW) { if (hBrushLow) DeleteObject(hBrushLow); hBrushLow = newBrush; }
                else if (id == IDC_HEX_MED) { if (hBrushMed) DeleteObject(hBrushMed); hBrushMed = newBrush; }
                else if (id == IDC_HEX_HIGH) { if (hBrushHigh) DeleteObject(hBrushHigh); hBrushHigh = newBrush; }
                InvalidateRect(GetDlgItem(hDlg, id), NULL, TRUE);
            }
            return TRUE;
        }
        switch (id) {
        case IDOK: {
            SaveUIToProfile(hDlg, s_currentTab);
            Config& p = g_Global.profiles[s_currentTab];
            if (p.tempLow < 0 || p.tempHigh > 110 || p.tempLow >= p.tempMed || p.tempMed >= p.tempHigh) {
                MessageBoxW(hDlg, L"Invalid temperature range in current profile.", L"Error", MB_ICONWARNING); return TRUE;
            }
            SetStartupTask(IsDlgButtonChecked(hDlg, IDC_CHK_STARTUP) == BST_CHECKED);
            SaveSettings(); g_ResetHttp = true; g_asyncTemp = -1.0f; g_activeSource = DataSource::Searching; forceLEDRefresh();
            SetEvent(g_hSensorEvent);
            EndDialog(hDlg, IDOK); return TRUE;
        }
        case IDCANCEL: EndDialog(hDlg, IDCANCEL); return TRUE;
        case IDC_BTN_RESET:
            SetDlgItemInt(hDlg, IDC_TEMP_LOW, 50, TRUE); SetDlgItemInt(hDlg, IDC_TEMP_MED, 70, TRUE); SetDlgItemInt(hDlg, IDC_TEMP_HIGH, 90, TRUE);
            SetDlgItemTextW(hDlg, IDC_HEX_LOW, L"#00FF00"); SetDlgItemTextW(hDlg, IDC_HEX_MED, L"#FFFF00"); SetDlgItemTextW(hDlg, IDC_HEX_HIGH, L"#FF0000");
            SendMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDC_HEX_LOW, EN_CHANGE), (LPARAM)GetDlgItem(hDlg, IDC_HEX_LOW));
            SendMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDC_HEX_MED, EN_CHANGE), (LPARAM)GetDlgItem(hDlg, IDC_HEX_MED));
            SendMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDC_HEX_HIGH, EN_CHANGE), (LPARAM)GetDlgItem(hDlg, IDC_HEX_HIGH));
            return TRUE;
        case IDC_BTN_RESET_ADVANCED:
            SendMessage(GetDlgItem(hDlg, IDC_COMBO_SENSOR_UPDATE), CB_SETCURSEL, 1, 0); SendMessage(GetDlgItem(hDlg, IDC_COMBO_LED_FPS), CB_SETCURSEL, 0, 0); SendMessage(GetDlgItem(hDlg, IDC_COMBO_SMOOTHING), CB_SETCURSEL, 2, 0);
            return TRUE;
        case IDC_EDIT_LABEL:
            if (code == EN_CHANGE) {
                wchar_t newLabel[64]; GetDlgItemTextW(hDlg, IDC_EDIT_LABEL, newLabel, 64);
                HWND hTab = GetDlgItem(hDlg, IDC_TAB_PROFILES);
                if (hTab) { int curTab = TabCtrl_GetCurSel(hTab); if (curTab != -1) { TCITEMW tie = { 0 }; tie.mask = TCIF_TEXT; tie.pszText = newLabel; SendMessage(hTab, TCM_SETITEMW, curTab, (LPARAM)&tie); } }
            }
            return TRUE;
        }
        break;
    }
    case WM_DESTROY:
        if (hBrushLow) DeleteObject(hBrushLow); if (hBrushMed) DeleteObject(hBrushMed); if (hBrushHigh) DeleteObject(hBrushHigh);
        ClearComboHeapData(GetDlgItem(hDlg, IDC_SENSOR_ID)); ClearComboHeapData(GetDlgItem(hDlg, IDC_COMBO_DEVICE));
        SetWindowLongPtr(GetDlgItem(hDlg, IDC_HEX_LOW), GWLP_WNDPROC, (LONG_PTR)oldEditProc);
        SetWindowLongPtr(GetDlgItem(hDlg, IDC_HEX_MED), GWLP_WNDPROC, (LONG_PTR)oldEditProc);
        SetWindowLongPtr(GetDlgItem(hDlg, IDC_HEX_HIGH), GWLP_WNDPROC, (LONG_PTR)oldEditProc);
        break;
    }
    return (INT_PTR)FALSE;
}

INT_PTR CALLBACK AboutDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    static HFONT hFontLink = NULL;
    switch (message) {
    case WM_INITDIALOG: {
        RECT rcDlg; GetWindowRect(hDlg, &rcDlg);
        int x = (GetSystemMetrics(SM_CXSCREEN) - (rcDlg.right - rcDlg.left)) / 2;
        int y = (GetSystemMetrics(SM_CYSCREEN) - (rcDlg.bottom - rcDlg.top)) / 2;
        SetWindowPos(hDlg, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE);
        HICON hIcon = (HICON)LoadImage(GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_ICON1), IMAGE_ICON, 48, 48, LR_SHARED);
        SendMessage(GetDlgItem(hDlg, IDC_ABOUT_ICON), STM_SETIMAGE, IMAGE_ICON, (LPARAM)hIcon);
        std::wstring versionStr = L"MysticFight " + std::wstring(APP_VERSION);
        SetDlgItemTextW(hDlg, IDC_ABOUT_VERSION, versionStr.c_str());
        HFONT hFont = (HFONT)SendMessage(GetDlgItem(hDlg, IDC_GITHUB_LINK), WM_GETFONT, 0, 0);
        LOGFONT lf; GetObject(hFont, sizeof(LOGFONT), &lf); lf.lfUnderline = TRUE;
        hFontLink = CreateFontIndirect(&lf);
        SendMessage(GetDlgItem(hDlg, IDC_GITHUB_LINK), WM_SETFONT, (WPARAM)hFontLink, TRUE);
        return (INT_PTR)TRUE;
    }
    case WM_CTLCOLORSTATIC:
        if ((HWND)lParam == GetDlgItem(hDlg, IDC_GITHUB_LINK)) {
            SetTextColor((HDC)wParam, RGB(0, 102, 204)); SetBkMode((HDC)wParam, TRANSPARENT); return (INT_PTR)GetStockObject(NULL_BRUSH);
        }
        break;
    case WM_SETCURSOR:
        if ((HWND)wParam == GetDlgItem(hDlg, IDC_GITHUB_LINK)) { SetCursor(LoadCursor(NULL, IDC_HAND)); return TRUE; }
        break;
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) { if (hFontLink) { DeleteObject(hFontLink); hFontLink = NULL; } EndDialog(hDlg, LOWORD(wParam)); return (INT_PTR)TRUE; }
        if (LOWORD(wParam) == IDC_GITHUB_LINK) RunShellNonAdmin(L"https://github.com/tonikelope/MysticFight");
        break;
    case WM_DESTROY: if (hFontLink) { DeleteObject(hFontLink); hFontLink = NULL; } break;
    }
    return (INT_PTR)FALSE;
}

void SwitchActiveProfile(HWND hWnd, int index) {
    if (index < 0 || index >= 5) return;
    { std::lock_guard<std::mutex> lock(g_cfgMutex); g_Global.activeProfileIndex = index; }
    SaveSettings();
    g_asyncTemp = -1.0f; g_ResetHttp = true; g_activeSource = DataSource::Searching; forceLEDRefresh();
    wchar_t msg[128]; swprintf_s(msg, L"Switched to %ls", g_cfg.label);
    ShowNotification(hWnd, GetModuleHandle(NULL), L"MysticFight", msg);
    char logMsg[128]; snprintf(logMsg, sizeof(logMsg), "[MysticFight] Hotkey triggered: Switched to Profile %d", index + 1); Log(logMsg);
}

static void FinalCleanup(HWND hWnd) {
    Log("[MysticFight] Starting cleaning...");
    if (g_hLibrary) {
        if (g_deviceName && lpMLAPI_SetLedStyle) { _bstr_t bstrOff(L"Off"); lpMLAPI_SetLedStyle(g_deviceName, 0, bstrOff); Log("[MysticFight] LEDs power off"); }
        if (lpMLAPI_Release) { lpMLAPI_Release(); Log("[MysticFight] MSI SDK Released"); }
        FreeLibrary(g_hLibrary); g_hLibrary = NULL;
    }
    if (g_deviceName) { SysFreeString(g_deviceName); g_deviceName = NULL; }
    if (hWnd) {
        NOTIFYICONDATA nid = { sizeof(NOTIFYICONDATA), hWnd, 1 }; Shell_NotifyIcon(NIM_DELETE, &nid);
        UnregisterHotKey(hWnd, 1); UnregisterHotKey(hWnd, 101); UnregisterHotKey(hWnd, 102); UnregisterHotKey(hWnd, 103);
    }
    if (g_hMutex) { ReleaseMutex(g_hMutex); CloseHandle(g_hMutex); g_hMutex = NULL; }
    CoUninitialize();
    if (g_hConnect) { WinHttpCloseHandle(g_hConnect); g_hConnect = NULL; }
    if (g_hSession) { WinHttpCloseHandle(g_hSession); g_hSession = NULL; }
    CloseHandle(g_hSensorEvent);
    Log("[MysticFight] Cleaning finished.");
}

// ============================================================================
// MAIN LOOP & ENTRY POINT
// ============================================================================

static LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_HOTKEY:
        if (wParam == 1) {
            g_LedsEnabled = !g_LedsEnabled;
            if (g_LedsEnabled) PlaySound(MAKEINTRESOURCE(IDR_WAV_LIGHTS_ON), GetModuleHandle(NULL), SND_RESOURCE | SND_ASYNC);
            else PlaySound(MAKEINTRESOURCE(IDR_WAV_LIGHTS_OFF), GetModuleHandle(NULL), SND_RESOURCE | SND_ASYNC);
            SetEvent(g_hSensorEvent); lastR = RGB_LED_REFRESH;
        }
        else if (wParam >= 101 && wParam <= 105) SwitchActiveProfile(hWnd, (int)wParam - 101);
        break;
    case WM_TRAYICON:
        if (lParam == WM_RBUTTONUP) {
            POINT curPoint; GetCursorPos(&curPoint); HMENU hMenu = CreatePopupMenu();
            if (hMenu) {
                InsertMenuW(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_CONFIG, L"Settings");
                InsertMenuW(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_ABOUT, L"About MysticFight");
                InsertMenuW(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_LOG, L"View Debug Log");
                InsertMenuW(hMenu, -1, MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
                InsertMenuW(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_EXIT, L"Exit MysticFight");
                SetForegroundWindow(hWnd); TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, curPoint.x, curPoint.y, 0, hWnd, NULL); DestroyMenu(hMenu);
            }
        }
        break;
    case WM_COMMAND:
        if (LOWORD(wParam) == ID_TRAY_EXIT) g_Running = false;
        if (LOWORD(wParam) == ID_TRAY_LOG) { wchar_t fullLogPath[MAX_PATH]; GetFullPathNameW(LOG_FILENAME, MAX_PATH, fullLogPath, NULL); RunShellNonAdmin(fullLogPath); }
        if (LOWORD(wParam) == ID_TRAY_CONFIG) { DialogBoxParamW(GetModuleHandle(NULL), MAKEINTRESOURCEW(IDD_SETTINGS), hWnd, (DLGPROC)SettingsDlgProc, 0); }
        if (LOWORD(wParam) == ID_TRAY_ABOUT) { DialogBoxParamW(GetModuleHandle(NULL), MAKEINTRESOURCEW(IDD_ABOUT), hWnd, (DLGPROC)AboutDlgProc, 0); }
        break;
    case WM_QUERYENDSESSION:
        if (!g_windows_shutdown) { Log("[MysticFight] Windows Shutdown detected..."); g_windows_shutdown = true; g_Running = false; }
        return FALSE;
    case WM_CLOSE: case WM_DESTROY: g_Running = false; PostQuitMessage(0); return 0;
    }
    return DefWindowProc(hWnd, message, wParam, lParam);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    SetPriorityClass(GetCurrentProcess(), HIGH_PRIORITY_CLASS); SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (FAILED(hr)) { Log("[MysticFight] Critical: COM Initialization failed."); return 1; }

    g_hMutex = CreateMutexW(NULL, TRUE, L"Global\\MysticFight_Unique_Mutex");
    if (g_hMutex == NULL || GetLastError() == ERROR_ALREADY_EXISTS) {
        if (g_hMutex) { MessageBoxW(NULL, L"MysticFight is already running.", L"Information", MB_OK | MB_ICONINFORMATION); CloseHandle(g_hMutex); }
        CoUninitialize(); return 0;
    }

    if (!ExtractMSIDLL() || !(g_hLibrary = LoadLibraryW(L"MysticLight_SDK.dll"))) {
        Log("[MysticFight] FATAL: Component preparation failed.");
        MessageBoxW(NULL, L"Critical Error: Required components (DLL) could not be prepared.", L"MysticFight - Error", MB_OK | MB_ICONERROR);
        if (g_hMutex) { ReleaseMutex(g_hMutex); CloseHandle(g_hMutex); } CoUninitialize(); return 1;
    }

    lpMLAPI_Initialize = (LPMLAPI_Initialize)GetProcAddress(g_hLibrary, "MLAPI_Initialize");
    lpMLAPI_GetDeviceInfo = (LPMLAPI_GetDeviceInfo)GetProcAddress(g_hLibrary, "MLAPI_GetDeviceInfo");
    lpMLAPI_SetLedColor = (LPMLAPI_SetLedColor)GetProcAddress(g_hLibrary, "MLAPI_SetLedColor");
    lpMLAPI_SetLedStyle = (LPMLAPI_SetLedStyle)GetProcAddress(g_hLibrary, "MLAPI_SetLedStyle");
    lpMLAPI_SetLedSpeed = (LPMLAPI_SetLedSpeed)GetProcAddress(g_hLibrary, "MLAPI_SetLedSpeed");
    lpMLAPI_Release = (LPMLAPI_Release)GetProcAddress(g_hLibrary, "MLAPI_Release");
    lpMLAPI_GetDeviceNameEx = (LPMLAPI_GetDeviceNameEx)GetProcAddress(g_hLibrary, "MLAPI_GetDeviceNameEx");
    lpMLAPI_GetLedInfo = (LPMLAPI_GetLedInfo)GetProcAddress(g_hLibrary, "MLAPI_GetLedInfo");

    TrimLogFile();
    char versionMsg[128]; snprintf(versionMsg, sizeof(versionMsg), "[MysticFight] MysticFight %ls started", APP_VERSION); Log(versionMsg);

    bool isFirstRun = (GetFileAttributesW(INI_FILE) == INVALID_FILE_ATTRIBUTES);
    LoadSettings(); g_target_device = g_cfg.targetDevice;

    wchar_t hL[10], hM[10], hH[10]; ColorToHex(g_cfg.colorLow, hL, 10); ColorToHex(g_cfg.colorMed, hM, 10); ColorToHex(g_cfg.colorHigh, hH, 10);
    char startupCfg[LOG_BUFFER_SIZE]; snprintf(startupCfg, sizeof(startupCfg), "[MysticFight] Config Loaded - Device: %ls | LED Area index: %d | Sensor: %ls | Low: %dºC (%ls) | Med: %dºC (%ls) | High: %dºC (%ls)", g_cfg.targetDevice, g_cfg.targetLedIndex, g_cfg.sensorID, g_cfg.tempLow, hL, g_cfg.tempMed, hM, g_cfg.tempHigh, hH); Log(startupCfg);

    wchar_t windowTitle[100]; swprintf_s(windowTitle, L"MysticFight %ls (by tonikelope)", APP_VERSION);
    WNDCLASSW wc = { 0 }; wc.lpfnWndProc = WndProc; wc.hInstance = hInstance; wc.lpszClassName = L"MysticFight_Class";
    wc.hIcon = (HICON)LoadImageW(hInstance, MAKEINTRESOURCEW(101), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
    RegisterClassW(&wc);
    HWND hWnd = CreateWindowExW(0, wc.lpszClassName, windowTitle, 0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);

    Log("[MysticFight] Initializing MSI SDK and HTTP Sensor Thread (Parallel Startup)...");
    g_hSensorEvent = CreateEventW(NULL, FALSE, FALSE, NULL);
    g_hSourceResolvedEvent = CreateEventW(NULL, TRUE, FALSE, NULL);
    std::thread sThread(SensorThread);

    ULONGLONG startTime = GetTickCount64(); bool sdkReady = false;
    while (GetTickCount64() - startTime < 5000) {
        if (lpMLAPI_Initialize && lpMLAPI_Initialize() == 0) { sdkReady = true; Log("[MysticFight] SDK Initialized successfully at startup"); break; }
        Sleep(500);
    }

    if (!sdkReady) {
        Log("[MysticFight] FATAL: MSI SDK Initialization failed after timeout.");
        MessageBoxW(NULL, L"The MSI Mystic Light SDK is not responding.\n\nPlease ensure MSI Center and Mystic Light are installed correctly.", L"MysticFight - SDK Error", MB_OK | MB_ICONERROR);
        if (g_hMutex) { ReleaseMutex(g_hMutex); CloseHandle(g_hMutex); g_hMutex = NULL; }
        g_Running = false; SetEvent(g_hSensorEvent); if (sThread.joinable()) sThread.join();
        if (g_hSensorEvent) CloseHandle(g_hSensorEvent); if (g_hSourceResolvedEvent) CloseHandle(g_hSourceResolvedEvent);
        CoUninitialize(); return 1;
    }

    if (sdkReady) MSIHwardwareDetection();
    ULONGLONG elapsed = GetTickCount64() - startTime; DWORD waitTime = (elapsed >= 5000) ? 0 : (5000 - (DWORD)elapsed);
    WaitForSingleObject(g_hSourceResolvedEvent, waitTime);

    if (g_activeSource == DataSource::Searching) {
        Log("[MysticFight] FATAL: No valid data source (HTTP) found during startup timeout.");
        MessageBoxW(NULL, L"Fatal Error: No temperature data source found (LibreHardwareMonitor HTTP).\n\nPlease ensure the Web Server is enabled in LHM options and use the default settings (0.0.0.0 and port 8085).", L"MysticFight - Startup Error", MB_OK | MB_ICONERROR);
        g_Running = false; SetEvent(g_hSensorEvent); sThread.join(); return 1;
    }

    if (isFirstRun) SaveSettings();

    NOTIFYICONDATAW nid = { sizeof(NOTIFYICONDATAW), hWnd, 1, NIF_ICON | NIF_MESSAGE | NIF_TIP, WM_TRAYICON };
    nid.hIcon = (HICON)LoadImageW(hInstance, MAKEINTRESOURCEW(IDI_ICON1), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
    swprintf_s(nid.szTip, _countof(nid.szTip), L"MysticFight %ls (by tonikelope)", APP_VERSION); Shell_NotifyIconW(NIM_ADD, &nid);

    if (isFirstRun) DialogBoxParamW(hInstance, MAKEINTRESOURCEW(IDD_SETTINGS), hWnd, (DLGPROC)SettingsDlgProc, 0);

    RegisterHotKey(hWnd, 1, MOD_CONTROL | MOD_SHIFT | MOD_ALT, 0x4C);
    RegisterHotKey(hWnd, 101, MOD_CONTROL | MOD_SHIFT | MOD_ALT, '1');
    RegisterHotKey(hWnd, 102, MOD_CONTROL | MOD_SHIFT | MOD_ALT, '2');
    RegisterHotKey(hWnd, 103, MOD_CONTROL | MOD_SHIFT | MOD_ALT, '3');
    RegisterHotKey(hWnd, 104, MOD_CONTROL | MOD_SHIFT | MOD_ALT, '4');
    RegisterHotKey(hWnd, 105, MOD_CONTROL | MOD_SHIFT | MOD_ALT, '5');
    ShowNotification(hWnd, hInstance, windowTitle, L"Let's dance baby");

    _bstr_t bstrOff(L"Off"); _bstr_t bstrBreath(L"Breath"); _bstr_t bstrSteady(L"Steady");
    DWORD targetR = 0, targetG = 0, targetB = 0;
    MSG msg = { 0 }; bool lhmAlive = true; bool firstRun = true; int status = 0;

    while (g_Running) {
        while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
            if (msg.message == WM_QUIT) { g_Running = false; break; }
            TranslateMessage(&msg); DispatchMessage(&msg);
        }
        if (!g_Running) break;
        ULONGLONG currentTime = GetTickCount64();

        if (g_Resetting_sdk) {
            if (currentTime >= g_ResetTimer) {
                switch (g_ResetStage) {
                case 0: Log("[MysticFight] Reset Stage 0: Cleaning MSI..."); ControlScheduledTask(L"MSI Task Host - LEDKeeper2_Host", false); KillProcessByName(L"LEDKeeper2.exe"); g_ResetTimer = currentTime + RESET_KILL_TASK_WAIT_MS; g_ResetStage = 1; break;
                case 1: Log("[MysticFight] Reset Stage 1: Restarting MSI..."); ControlScheduledTask(L"MSI Task Host - LEDKeeper2_Host", true); g_ResetTimer = currentTime + RESET_RESTART_TASK_DELAY_MS; g_ResetStage = 2; break;
                case 2: Log("[MysticFight] Reset Stage 2: Re-initializing SDK..."); if (lpMLAPI_Initialize && lpMLAPI_Initialize() == 0) { MSIHwardwareDetection(); if (lpMLAPI_SetLedStyle) lpMLAPI_SetLedStyle(g_deviceName, 0, bstrSteady); } g_ResetTimer = currentTime + RESET_RESTART_TASK_DELAY_MS; g_Resetting_sdk = false; forceLEDRefresh(); break;
                }
            }
        }
        else if (g_activeSource == DataSource::Searching) {
            if (lhmAlive) {
                lhmAlive = false;
                if (lpMLAPI_SetLedStyle) status = lpMLAPI_SetLedStyle(g_target_device, g_cfg.targetLedIndex, bstrBreath);
                if (status == 0 && lpMLAPI_SetLedColor) status = lpMLAPI_SetLedColor(g_target_device, g_cfg.targetLedIndex, 255, 255, 255);
            }
        }
        else if (g_LedsEnabled) {
            float rawTemp = g_asyncTemp.load();
            if (rawTemp >= 0.0f) {
                if (!lhmAlive) { lhmAlive = true; Log("[MysticFight] Connection established/recovered."); forceLEDRefresh(); }
                float temp = floorf(rawTemp * 2.0f + 0.5f) / 2.0f;
                if (temp != lastTemp) {
                    lastTemp = temp;
                    float ratio = 0.0f; COLORREF c1 = 0, c2 = 0;
                    if (temp <= (float)g_cfg.tempLow) { targetR = GetRValue(g_cfg.colorLow); targetG = GetGValue(g_cfg.colorLow); targetB = GetBValue(g_cfg.colorLow); }
                    else if (temp < (float)g_cfg.tempMed) {
                        ratio = (temp - (float)g_cfg.tempLow) / ((float)g_cfg.tempMed - (float)g_cfg.tempLow);
                        if (ratio < 0.0f) ratio = 0.0f; if (ratio > 1.0f) ratio = 1.0f;
                        c1 = g_cfg.colorLow; c2 = g_cfg.colorMed;
                        targetR = (DWORD)sqrt((double)GetRValue(c1) * GetRValue(c1) * (1.0 - ratio) + (double)GetRValue(c2) * GetRValue(c2) * ratio);
                        targetG = (DWORD)sqrt((double)GetGValue(c1) * GetGValue(c1) * (1.0 - ratio) + (double)GetGValue(c2) * GetGValue(c2) * ratio);
                        targetB = (DWORD)sqrt((double)GetBValue(c1) * GetBValue(c1) * (1.0 - ratio) + (double)GetBValue(c2) * GetBValue(c2) * ratio);
                    }
                    else if (temp < (float)g_cfg.tempHigh) {
                        ratio = (temp - (float)g_cfg.tempMed) / ((float)g_cfg.tempHigh - (float)g_cfg.tempMed);
                        if (ratio < 0.0f) ratio = 0.0f; if (ratio > 1.0f) ratio = 1.0f;
                        c1 = g_cfg.colorMed; c2 = g_cfg.colorHigh;
                        targetR = (DWORD)sqrt((double)GetRValue(c1) * GetRValue(c1) * (1.0 - ratio) + (double)GetRValue(c2) * GetRValue(c2) * ratio);
                        targetG = (DWORD)sqrt((double)GetGValue(c1) * GetGValue(c1) * (1.0 - ratio) + (double)GetGValue(c2) * GetGValue(c2) * ratio);
                        targetB = (DWORD)sqrt((double)GetBValue(c1) * GetBValue(c1) * (1.0 - ratio) + (double)GetBValue(c2) * GetBValue(c2) * ratio);
                    }
                    else { targetR = GetRValue(g_cfg.colorHigh); targetG = GetGValue(g_cfg.colorHigh); targetB = GetBValue(g_cfg.colorHigh); }
                    if (firstRun) { currR = (float)targetR; currG = (float)targetG; currB = (float)targetB; firstRun = false; }
                }
            }
            float r = currR.load(); float g = currG.load(); float b = currB.load();
            r += ((float)targetR - r) * g_cfg.smoothingFactor; g += ((float)targetG - g) * g_cfg.smoothingFactor; b += ((float)targetB - b) * g_cfg.smoothingFactor;
            currR.store(r); currG.store(g); currB.store(b);
            DWORD sendR = (DWORD)currR; DWORD sendG = (DWORD)currG; DWORD sendB = (DWORD)currB;
            if (sendR != lastR || sendG != lastG || sendB != lastB) {
                if (lastR == RGB_LEDS_OFF || lastR == RGB_LED_REFRESH) { if (lpMLAPI_SetLedStyle) status = lpMLAPI_SetLedStyle(g_target_device, g_cfg.targetLedIndex, bstrSteady); }
                if (status == 0 && lpMLAPI_SetLedColor) status = lpMLAPI_SetLedColor(g_target_device, g_cfg.targetLedIndex, sendR, sendG, sendB);
                if (status != 0) { Log("[MysticFight] SDK SetColor failed. Triggering Reset..."); g_Resetting_sdk = true; g_ResetStage = 0; g_ResetTimer = 0; }
                else { lastR = sendR; lastG = sendG; lastB = sendB; }
            }
        }
        else {
            if (lastR != RGB_LEDS_OFF) {
                int status = 0; if (lpMLAPI_SetLedStyle) status = lpMLAPI_SetLedStyle(g_deviceName, 0, bstrOff);
                if (status != 0) { g_Resetting_sdk = true; g_ResetStage = 0; g_ResetTimer = 0; }
                else { lastR = RGB_LEDS_OFF; }
            }
        }
        MsgWaitForMultipleObjects(0, NULL, FALSE, g_LedsEnabled ? (DWORD)(1000 / g_cfg.ledRefreshFPS) : (DWORD)MAIN_LOOP_OFF_DELAY_MS, QS_ALLINPUT);
    }

    if (g_windows_shutdown) ShutdownBlockReasonCreate(hWnd, L"Mystic Fight Shutdown...");
    SetEvent(g_hSensorEvent); if (sThread.joinable()) sThread.join();
    FinalCleanup(hWnd); Log("[MysticFight] BYE BYE");
    if (g_windows_shutdown) ShutdownBlockReasonDestroy(hWnd);
    return 0;
}