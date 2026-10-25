#include "RazerOrnataController.h"

#include <windows.h>
#include <devpkey.h>
#include <hidsdi.h>
#include <setupapi.h>

#include <algorithm>
#include <cwctype>
#include <vector>

#pragma comment(lib, "hid.lib")
#pragma comment(lib, "setupapi.lib")

namespace {
constexpr std::size_t kReportSize = 90;
constexpr std::size_t kFeatureReportSize = kReportSize + 1; // leading report ID
constexpr std::uint8_t kTransactionId = 0x1F;
constexpr std::uint8_t kBacklightLed = 0x05;
constexpr int kMaxAttempts = 5;
constexpr DWORD kInterTransferDelayMs = 1;
constexpr DWORD kRetryDelayMs = 10;
constexpr DEVPROPKEY kDeviceContainerId = {
    { 0x8C7ED206, 0x3F8A, 0x4827, { 0xB3, 0xAB, 0xAE, 0x9E, 0x1F, 0xAE, 0xFC, 0x6C } },
    2
};

constexpr std::uint8_t kStatusBusy = 0x01;
constexpr std::uint8_t kStatusSuccessful = 0x02;
constexpr std::uint8_t kStatusNotSupported = 0x05;
using ControlReport = std::array<std::uint8_t, kReportSize>;

struct DeviceProfile {
    USHORT productId;
    std::size_t zoneCount;
    std::uint8_t preferredInterface;
};

constexpr std::array<DeviceProfile, 6> kSupportedDeviceProfiles = {{
    { RazerOrnataController::kProductIdTartarusPro, 21, 2 },
    { RazerOrnataController::kProductIdOrnataV3Alternate, 10, 2 },
    { RazerOrnataController::kProductIdOrnataV3X, 1, 1 },
    { RazerOrnataController::kProductIdOrnataV3, 10, 2 },
    { RazerOrnataController::kProductIdOrnataV3XAlternate, 1, 1 },
    { RazerOrnataController::kProductIdOrnataV3Tenkeyless, 8, 2 },
}};

constexpr const DeviceProfile* ProfileForProduct(const USHORT productId) {
    for (const auto& profile : kSupportedDeviceProfiles) {
        if (profile.productId == productId) return &profile;
    }
    return nullptr;
}

constexpr void FinalizeControlReport(ControlReport& report) {
    std::uint8_t crc = 0;
    for (std::size_t i = 2; i <= 87; ++i) crc ^= report[i];
    report[88] = crc;
    report[89] = 0;
}

constexpr ControlReport BuildCustomFrameReport(
    const std::array<RazerOrnataController::Rgb, RazerOrnataController::kMaxZoneCount>& colors,
    const std::size_t zoneCount) {
    ControlReport frame = {};
    frame[1] = kTransactionId;
    frame[5] = 0x47;
    frame[6] = 0x0F;
    frame[7] = 0x03;
    frame[10] = 0x00; // row
    frame[11] = 0x00; // first zone
    frame[12] = static_cast<std::uint8_t>(zoneCount - 1); // last zone
    for (std::size_t i = 0; i < zoneCount; ++i) {
        const std::size_t offset = 13 + (i * 3);
        frame[offset] = colors[i].red;
        frame[offset + 1] = colors[i].green;
        frame[offset + 2] = colors[i].blue;
    }
    FinalizeControlReport(frame);
    return frame;
}

constexpr ControlReport BuildCustomEffectReport() {
    ControlReport activate = {};
    activate[1] = kTransactionId;
    activate[5] = 0x0C;
    activate[6] = 0x0F;
    activate[7] = 0x02;
    activate[10] = 0x08; // extended CUSTOM effect; variable storage = 0
    FinalizeControlReport(activate);
    return activate;
}

constexpr ControlReport BuildSpectrumEffectReport(const std::uint8_t led) {
    ControlReport spectrum = {};
    spectrum[1] = kTransactionId;
    spectrum[5] = 0x06;
    spectrum[6] = 0x0F;
    spectrum[7] = 0x02;
    spectrum[8] = 0x01; // VARSTORE: restore the standalone device effect
    spectrum[9] = led;
    spectrum[10] = 0x03; // OpenRazer extended-matrix Spectrum Cycling
    FinalizeControlReport(spectrum);
    return spectrum;
}

constexpr ControlReport BuildLightsOffEffectReport(const std::uint8_t led) {
    ControlReport off = {};
    off[1] = kTransactionId;
    off[5] = 0x06;
    off[6] = 0x0F;
    off[7] = 0x02;
    off[8] = 0x01; // VARSTORE: persist the released standalone state
    off[9] = led;
    off[10] = 0x00; // OpenRazer extended-matrix NONE effect
    FinalizeControlReport(off);
    return off;
}

constexpr std::array<RazerOrnataController::Rgb, RazerOrnataController::kMaxZoneCount>
    kBlackProtocolTest = {};
constexpr ControlReport kBlackV3FrameProtocolTest =
    BuildCustomFrameReport(kBlackProtocolTest, 10);
constexpr ControlReport kBlackV3XFrameProtocolTest = BuildCustomFrameReport(kBlackProtocolTest, 1);
constexpr ControlReport kBlackV3TenkeylessFrameProtocolTest =
    BuildCustomFrameReport(kBlackProtocolTest, 8);
constexpr ControlReport kBlackTartarusProFrameProtocolTest =
    BuildCustomFrameReport(kBlackProtocolTest, 21);
constexpr ControlReport kCustomEffectProtocolTest = BuildCustomEffectReport();
constexpr ControlReport kSpectrumEffectProtocolTest = BuildSpectrumEffectReport(kBacklightLed);
constexpr ControlReport kLightsOffEffectProtocolTest = BuildLightsOffEffectReport(kBacklightLed);
static_assert(kReportSize == 90 && kFeatureReportSize == 91);
static_assert(kBlackV3FrameProtocolTest[5] == 0x47 &&
    kBlackV3FrameProtocolTest[10] == 0 && kBlackV3FrameProtocolTest[12] == 9 &&
    kBlackV3FrameProtocolTest[88] == 0x42);
static_assert(kBlackV3XFrameProtocolTest[5] == 0x47 &&
    kBlackV3XFrameProtocolTest[10] == 0 && kBlackV3XFrameProtocolTest[12] == 0 &&
    kBlackV3XFrameProtocolTest[88] == 0x4B);
static_assert(kBlackV3TenkeylessFrameProtocolTest[5] == 0x47 &&
    kBlackV3TenkeylessFrameProtocolTest[10] == 0 &&
    kBlackV3TenkeylessFrameProtocolTest[12] == 7 &&
    kBlackV3TenkeylessFrameProtocolTest[88] == 0x4C);
static_assert(kBlackTartarusProFrameProtocolTest[5] == 0x47 &&
    kBlackTartarusProFrameProtocolTest[10] == 0 &&
    kBlackTartarusProFrameProtocolTest[12] == 20 &&
    kBlackTartarusProFrameProtocolTest[88] == 0x5F);
static_assert(kCustomEffectProtocolTest[5] == 0x0C &&
    kCustomEffectProtocolTest[10] == 0x08 && kCustomEffectProtocolTest[88] == 0x09);
static_assert(kSpectrumEffectProtocolTest[5] == 0x06 &&
    kSpectrumEffectProtocolTest[8] == 0x01 && kSpectrumEffectProtocolTest[9] == 0x05 &&
    kSpectrumEffectProtocolTest[10] == 0x03 &&
    kSpectrumEffectProtocolTest[88] == 0x0C);
static_assert(kLightsOffEffectProtocolTest[5] == 0x06 &&
    kLightsOffEffectProtocolTest[8] == 0x01 && kLightsOffEffectProtocolTest[9] == 0x05 &&
    kLightsOffEffectProtocolTest[10] == 0x00 &&
    kLightsOffEffectProtocolTest[88] == 0x0F);

bool IsSupportedProduct(const USHORT productId) {
    return ProfileForProduct(productId) != nullptr;
}

std::size_t ZoneCountForProduct(const USHORT productId) {
    const auto* profile = ProfileForProduct(productId);
    return profile ? profile->zoneCount : 0;
}

std::wstring ToLower(std::wstring value) {
    std::transform(value.begin(), value.end(), value.begin(),
        [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
    return value;
}

bool IsPreferredInterface(const std::wstring& path, const USHORT productId) {
    const auto* profile = ProfileForProduct(productId);
    if (!profile) return false;
    const std::wstring lower = ToLower(path);
    // OpenRazer addresses the supported V3/V3 Alt/V3 TKL/Tartarus Pro models
    // on interface 2; the single-zone V3 X/V3 X Alt variants use interface 1.
    const wchar_t* marker = profile->preferredInterface == 1 ? L"&mi_01" : L"&mi_02";
    return lower.find(marker) != std::wstring::npos;
}

std::wstring FormatWin32Error(const wchar_t* operation, DWORD error) {
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"%ls failed (Win32 %lu)", operation,
        static_cast<unsigned long>(error));
    return buffer;
}

HANDLE OpenHidPath(const std::wstring& path) {
    constexpr DWORD accessAttempts[] = { GENERIC_READ | GENERIC_WRITE, GENERIC_WRITE, 0 };
    for (DWORD access : accessAttempts) {
        HANDLE handle = CreateFileW(path.c_str(), access,
            FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr);
        if (handle != INVALID_HANDLE_VALUE) return handle;
    }
    return INVALID_HANDLE_VALUE;
}

struct EnumerationResult {
    std::vector<RazerOrnataController::DetectedDevice> devices;
};

std::wstring ContainerKey(HDEVINFO deviceInfo, SP_DEVINFO_DATA& deviceData,
    const USHORT productId, const std::wstring& productName) {
    GUID container = {};
    DEVPROPTYPE type = 0;
    DWORD required = 0;
    if (SetupDiGetDevicePropertyW(deviceInfo, &deviceData, &kDeviceContainerId,
        &type, reinterpret_cast<PBYTE>(&container), sizeof(container), &required, 0) &&
        type == DEVPROP_TYPE_GUID && !IsEqualGUID(container, GUID_NULL)) {
        wchar_t key[80] = {};
        swprintf_s(key, L"{%08lX-%04hX-%04hX-%02hhX%02hhX-%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX}|PID_%04hX",
            container.Data1, container.Data2, container.Data3,
            container.Data4[0], container.Data4[1], container.Data4[2], container.Data4[3],
            container.Data4[4], container.Data4[5], container.Data4[6], container.Data4[7],
            productId);
        return key;
    }

    // The container identifier is normally available. This fallback remains
    // deterministic, but identical devices without a container may be merged.
    wchar_t fallback[384] = {};
    swprintf_s(fallback, L"VID_1532&PID_%04X|%ls", productId, productName.c_str());
    return fallback;
}

EnumerationResult EnumerateCandidates(std::wstring& error) {
    EnumerationResult result;
    GUID hidGuid = {};
    HidD_GetHidGuid(&hidGuid);

    HDEVINFO deviceInfo = SetupDiGetClassDevsW(&hidGuid, nullptr, nullptr,
        DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
    if (deviceInfo == INVALID_HANDLE_VALUE) {
        error = FormatWin32Error(L"SetupDiGetClassDevsW", GetLastError());
        return result;
    }

    for (DWORD index = 0;; ++index) {
        SP_DEVICE_INTERFACE_DATA interfaceData = {};
        interfaceData.cbSize = sizeof(interfaceData);
        if (!SetupDiEnumDeviceInterfaces(deviceInfo, nullptr, &hidGuid, index, &interfaceData)) {
            const DWORD enumError = GetLastError();
            if (enumError != ERROR_NO_MORE_ITEMS) {
                error = FormatWin32Error(L"SetupDiEnumDeviceInterfaces", enumError);
            }
            break;
        }

        DWORD requiredSize = 0;
        SetupDiGetDeviceInterfaceDetailW(deviceInfo, &interfaceData, nullptr, 0,
            &requiredSize, nullptr);
        if (requiredSize < sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA_W)) continue;

        std::vector<std::uint8_t> storage(requiredSize);
        auto* detail = reinterpret_cast<SP_DEVICE_INTERFACE_DETAIL_DATA_W*>(storage.data());
        detail->cbSize = sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA_W);
        SP_DEVINFO_DATA deviceData = {};
        deviceData.cbSize = sizeof(deviceData);
        if (!SetupDiGetDeviceInterfaceDetailW(deviceInfo, &interfaceData, detail,
            requiredSize, nullptr, &deviceData)) {
            continue;
        }

        const std::wstring path = detail->DevicePath;
        // Avoid touching unrelated HID collections. Identity and the exact
        // model-specific control interface are verified below using read-only
        // inspection; no feature report is sent during detection.
        if (ToLower(path).find(L"vid_1532") == std::wstring::npos) continue;

        HANDLE handle = CreateFileW(path.c_str(), 0, FILE_SHARE_READ | FILE_SHARE_WRITE,
            nullptr, OPEN_EXISTING, 0, nullptr);
        if (handle == INVALID_HANDLE_VALUE) continue;

        HIDD_ATTRIBUTES attributes = {};
        attributes.Size = sizeof(attributes);
        PHIDP_PREPARSED_DATA preparsed = nullptr;
        HIDP_CAPS caps = {};
        const bool attributesOk = HidD_GetAttributes(handle, &attributes) == TRUE;
        const bool capsOk = HidD_GetPreparsedData(handle, &preparsed) == TRUE &&
            HidP_GetCaps(preparsed, &caps) == HIDP_STATUS_SUCCESS;

        const bool compatibleInterface = attributesOk &&
            attributes.VendorID == RazerOrnataController::kVendorId &&
            IsSupportedProduct(attributes.ProductID) && capsOk &&
            caps.FeatureReportByteLength == kFeatureReportSize;

        wchar_t productBuffer[256] = {};
        if (compatibleInterface) {
            if (!HidD_GetProductString(handle, productBuffer, sizeof(productBuffer))) {
                productBuffer[0] = L'\0';
            }
        }

        if (preparsed) HidD_FreePreparsedData(preparsed);
        CloseHandle(handle);

        // Unsupported models and supported PIDs without the validated control
        // collection are intentionally omitted from the public inventory.
        if (!compatibleInterface) continue;

        std::wstring productName = productBuffer;
        if (productName.empty()) {
            wchar_t fallbackName[64] = {};
            swprintf_s(fallbackName, L"Razer device (PID 0x%04X)", attributes.ProductID);
            productName = fallbackName;
        }
        const std::wstring deviceKey = ContainerKey(deviceInfo, deviceData,
            attributes.ProductID, productName);
        auto deviceIt = std::find_if(result.devices.begin(), result.devices.end(),
            [&](const RazerOrnataController::DetectedDevice& device) {
                return device.deviceKey == deviceKey;
            });
        if (deviceIt == result.devices.end()) {
            RazerOrnataController::DetectedDevice device;
            device.deviceKey = deviceKey;
            device.productId = attributes.ProductID;
            device.productName = productName;
            result.devices.push_back(std::move(device));
            deviceIt = std::prev(result.devices.end());
        }
        else if (productBuffer[0] != L'\0' &&
            deviceIt->productName.starts_with(L"Razer device (PID")) {
            deviceIt->productName = productBuffer;
        }

        // A 91-byte feature report is authoritative. The upstream interface is
        // only a tie-breaker because Windows can expose the control collection
        // on another MI_XX path (observed on Ornata V3 X PID 0x02A2).
        if (!deviceIt->controllable || IsPreferredInterface(path, attributes.ProductID)) {
            deviceIt->controllable = true;
            deviceIt->controlPath = path;
        }
    }

    SetupDiDestroyDeviceInfoList(deviceInfo);
    std::sort(result.devices.begin(), result.devices.end(),
        [](const RazerOrnataController::DetectedDevice& left,
            const RazerOrnataController::DetectedDevice& right) {
            if (left.productName != right.productName) return left.productName < right.productName;
            return left.deviceKey < right.deviceKey;
        });
    return result;
}

bool ResponseMatches(const std::array<std::uint8_t, kFeatureReportSize>& feature,
    const std::array<std::uint8_t, kReportSize>& request) {
    // feature[0] is the HID report ID. The Razer response starts at feature[1].
    std::uint8_t crc = 0;
    for (std::size_t i = 3; i <= 88; ++i) crc ^= feature[i];
    if (crc != feature[89]) return false;

    return feature[3] == request[2] && feature[4] == request[3] &&
        feature[7] == request[6] && feature[8] == request[7];
}
} // namespace

RazerOrnataController::~RazerOrnataController() {
    Close();
}

RazerOrnataController::DetectionResult RazerOrnataController::Detect() {
    DetectionResult result;
    std::wstring error;
    const auto enumeration = EnumerateCandidates(error);
    result.devices = enumeration.devices;
    result.error = error;
    const auto selected = std::find_if(result.devices.begin(), result.devices.end(),
        [](const DetectedDevice& device) { return device.controllable; });
    if (selected == result.devices.end()) return result;

    result.detected = true;
    result.productId = selected->productId;
    result.productName = selected->productName;
    result.devicePath = selected->controlPath;
    return result;
}

bool RazerOrnataController::Open(const DetectedDevice& device) {
    Close();
    if (!device.controllable || !IsSupportedProduct(device.productId) ||
        device.controlPath.empty()) {
        lastError_ = L"Rejected unverified Razer HID identity";
        return false;
    }

    HANDLE handle = OpenHidPath(device.controlPath);
    if (handle == INVALID_HANDLE_VALUE) {
        lastError_ = FormatWin32Error(L"CreateFileW", GetLastError());
        return false;
    }

    HIDD_ATTRIBUTES attributes = {};
    attributes.Size = sizeof(attributes);
    PHIDP_PREPARSED_DATA preparsed = nullptr;
    HIDP_CAPS caps = {};
    const bool identityOk = HidD_GetAttributes(handle, &attributes) == TRUE &&
        attributes.VendorID == kVendorId &&
        attributes.ProductID == device.productId &&
        IsSupportedProduct(attributes.ProductID);
    const bool capsOk = HidD_GetPreparsedData(handle, &preparsed) == TRUE &&
        HidP_GetCaps(preparsed, &caps) == HIDP_STATUS_SUCCESS &&
        caps.FeatureReportByteLength == kFeatureReportSize;
    if (preparsed) HidD_FreePreparsedData(preparsed);

    if (!identityOk || !capsOk) {
        CloseHandle(handle);
        lastError_ = L"Razer HID identity changed between detection and open";
        return false;
    }

    handle_ = handle;
    productId_ = device.productId;
    customEffectActive_ = false;
    haveLastFrame_ = false;
    lastError_.clear();
    return true;
}

void RazerOrnataController::Close() {
    HANDLE handle = static_cast<HANDLE>(handle_);
    if (handle != INVALID_HANDLE_VALUE && handle != nullptr) CloseHandle(handle);
    handle_ = INVALID_HANDLE_VALUE;
    productId_ = 0;
    customEffectActive_ = false;
    haveLastFrame_ = false;
}

bool RazerOrnataController::IsOpen() const {
    const HANDLE handle = static_cast<HANDLE>(handle_);
    return handle != INVALID_HANDLE_VALUE && handle != nullptr;
}

bool RazerOrnataController::SetUniformColor(std::uint8_t red, std::uint8_t green,
    std::uint8_t blue) {
    std::array<Rgb, kMaxZoneCount> colors = {};
    colors.fill({ red, green, blue });
    return SetZoneColors(colors);
}

bool RazerOrnataController::SetZoneColors(const std::array<Rgb, kMaxZoneCount>& colors) {
    if (!IsOpen()) {
        lastError_ = L"Razer HID handle is not open";
        return false;
    }
    if (haveLastFrame_ && colors == lastFrame_) return true;

    // These explicitly whitelisted OpenRazer recipes share transaction 0x1F
    // and one extended-matrix row. The last column is model-specific.
    const std::size_t zoneCount = ZoneCountForProduct(productId_);
    if (zoneCount == 0 || zoneCount > kMaxZoneCount) {
        lastError_ = L"Razer HID product has no whitelisted lighting geometry";
        return false;
    }
    auto frame = BuildCustomFrameReport(colors, zoneCount);
    if (!SendReport(frame)) return false;

    if (!customEffectActive_) {
        auto activate = BuildCustomEffectReport();
        if (!SendReport(activate)) return false;
        customEffectActive_ = true;
    }

    lastFrame_ = colors;
    haveLastFrame_ = true;
    return true;
}

bool RazerOrnataController::SetSpectrumCycling() {
    if (!IsOpen()) {
        lastError_ = L"Razer HID handle is not open";
        return false;
    }
    auto spectrum = BuildSpectrumEffectReport(kBacklightLed);
    if (!SendReport(spectrum)) return false;
    customEffectActive_ = false;
    haveLastFrame_ = false;
    return true;
}

bool RazerOrnataController::SetLightsOff() {
    if (!IsOpen()) {
        lastError_ = L"Razer HID handle is not open";
        return false;
    }
    auto off = BuildLightsOffEffectReport(kBacklightLed);
    if (!SendReport(off)) return false;
    customEffectActive_ = false;
    haveLastFrame_ = false;
    return true;
}

bool RazerOrnataController::IsWhitelistedReport(
    const std::array<std::uint8_t, kReportSize>& report) const {
    if (report[0] != 0 || report[1] != kTransactionId || report[2] != 0 ||
        report[3] != 0 || report[4] != 0 || report[89] != 0) {
        return false;
    }

    std::uint8_t crc = 0;
    for (std::size_t i = 2; i <= 87; ++i) crc ^= report[i];
    if (crc != report[88]) return false;

    if (report[5] == 0x47 && report[6] == 0x0F && report[7] == 0x03) {
        const std::size_t zoneCount = ZoneCountForProduct(productId_);
        if (zoneCount == 0 || zoneCount > kMaxZoneCount) return false;
        if (report[8] != 0 || report[9] != 0 || report[10] != 0 ||
            report[11] != 0 || report[12] != zoneCount - 1) return false;
        for (std::size_t i = 13 + (zoneCount * 3); i <= 87; ++i) {
            if (report[i] != 0) return false;
        }
        return true;
    }

    if (report[5] == 0x0C && report[6] == 0x0F && report[7] == 0x02) {
        if (report[8] != 0 || report[9] != 0 || report[10] != 0x08) return false;
        for (std::size_t i = 11; i <= 87; ++i) {
            if (report[i] != 0) return false;
        }
        return true;
    }

    if (report[5] == 0x06 && report[6] == 0x0F && report[7] == 0x02) {
        if (report[8] != 0x01 || report[9] != kBacklightLed ||
            (report[10] != 0x00 && report[10] != 0x03)) return false;
        for (std::size_t i = 11; i <= 87; ++i) {
            if (report[i] != 0) return false;
        }
        return true;
    }

    return false;
}

bool RazerOrnataController::SendReport(std::array<std::uint8_t, kReportSize>& report) {
    if (!IsOpen() || !IsWhitelistedReport(report)) {
        lastError_ = L"Blocked non-whitelisted Razer report";
        return false;
    }

    HANDLE handle = static_cast<HANDLE>(handle_);
    DWORD finalError = ERROR_GEN_FAILURE;
    for (int attempt = 0; attempt < kMaxAttempts; ++attempt) {
        std::array<std::uint8_t, kFeatureReportSize> feature = {};
        std::copy(report.begin(), report.end(), feature.begin() + 1);
        if (!HidD_SetFeature(handle, feature.data(), static_cast<ULONG>(feature.size()))) {
            finalError = GetLastError();
            Sleep(kRetryDelayMs);
            continue;
        }

        Sleep(kInterTransferDelayMs);
        feature.fill(0);
        if (!HidD_GetFeature(handle, feature.data(), static_cast<ULONG>(feature.size()))) {
            finalError = GetLastError();
            Sleep(kRetryDelayMs);
            continue;
        }

        if (!ResponseMatches(feature, report)) {
            lastError_ = L"Razer response did not match the request";
            Sleep(kRetryDelayMs);
            continue;
        }

        const std::uint8_t status = feature[1];
        if (status == kStatusSuccessful || status == kStatusBusy) {
            lastError_.clear();
            return true;
        }
        if (status == kStatusNotSupported) {
            lastError_ = L"Razer device rejected a whitelisted command as unsupported";
            return false;
        }
        wchar_t statusError[96] = {};
        swprintf_s(statusError, L"Razer device returned status 0x%02X", status);
        lastError_ = statusError;
        Sleep(kRetryDelayMs);
    }

    if (lastError_.empty()) lastError_ = FormatWin32Error(L"Razer HID transfer", finalError);
    return false;
}
