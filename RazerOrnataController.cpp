#include "RazerOrnataController.h"

#include <windows.h>
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
constexpr int kMaxAttempts = 5;
constexpr DWORD kInterTransferDelayMs = 1;
constexpr DWORD kRetryDelayMs = 10;

constexpr std::uint8_t kStatusBusy = 0x01;
constexpr std::uint8_t kStatusSuccessful = 0x02;
constexpr std::uint8_t kStatusNotSupported = 0x05;
using ControlReport = std::array<std::uint8_t, kReportSize>;

constexpr void FinalizeControlReport(ControlReport& report) {
    std::uint8_t crc = 0;
    for (std::size_t i = 2; i <= 87; ++i) crc ^= report[i];
    report[88] = crc;
    report[89] = 0;
}

constexpr ControlReport BuildCustomFrameReport(
    const std::array<RazerOrnataController::Rgb, RazerOrnataController::kZoneCount>& colors,
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

constexpr std::array<RazerOrnataController::Rgb, RazerOrnataController::kZoneCount>
    kBlackProtocolTest = {};
constexpr ControlReport kBlackV3FrameProtocolTest =
    BuildCustomFrameReport(kBlackProtocolTest, RazerOrnataController::kZoneCount);
constexpr ControlReport kBlackV3XFrameProtocolTest = BuildCustomFrameReport(kBlackProtocolTest, 1);
constexpr ControlReport kCustomEffectProtocolTest = BuildCustomEffectReport();
static_assert(kReportSize == 90 && kFeatureReportSize == 91);
static_assert(kBlackV3FrameProtocolTest[5] == 0x47 &&
    kBlackV3FrameProtocolTest[10] == 0 && kBlackV3FrameProtocolTest[12] == 9 &&
    kBlackV3FrameProtocolTest[88] == 0x42);
static_assert(kBlackV3XFrameProtocolTest[5] == 0x47 &&
    kBlackV3XFrameProtocolTest[10] == 0 && kBlackV3XFrameProtocolTest[12] == 0 &&
    kBlackV3XFrameProtocolTest[88] == 0x4B);
static_assert(kCustomEffectProtocolTest[5] == 0x0C &&
    kCustomEffectProtocolTest[10] == 0x08 && kCustomEffectProtocolTest[88] == 0x09);

bool IsSupportedProduct(const USHORT productId) {
    return productId == RazerOrnataController::kProductIdOrnataV3 ||
        productId == RazerOrnataController::kProductIdOrnataV3Alternate ||
        productId == RazerOrnataController::kProductIdOrnataV3X ||
        productId == RazerOrnataController::kProductIdOrnataV3XAlternate;
}

bool IsOrnataV3X(const USHORT productId) {
    return productId == RazerOrnataController::kProductIdOrnataV3X ||
        productId == RazerOrnataController::kProductIdOrnataV3XAlternate;
}

std::size_t ZoneCountForProduct(const USHORT productId) {
    return IsOrnataV3X(productId) ? 1 : RazerOrnataController::kZoneCount;
}

std::wstring ToLower(std::wstring value) {
    std::transform(value.begin(), value.end(), value.begin(),
        [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
    return value;
}

bool IsPreferredInterface(const std::wstring& path, const USHORT productId) {
    const std::wstring lower = ToLower(path);
    // OpenRazer addresses Ornata V3/V3 Alt on interface 2, while the
    // single-zone Ornata V3 X/V3 X Alt variants use interface 1.
    const wchar_t* marker = IsOrnataV3X(productId) ? L"&mi_01" : L"&mi_02";
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

struct Candidate {
    std::wstring path;
    USHORT productId = 0;
    std::wstring productName;
};

std::vector<Candidate> EnumerateCandidates(std::wstring& error) {
    std::vector<Candidate> candidates;
    GUID hidGuid = {};
    HidD_GetHidGuid(&hidGuid);

    HDEVINFO deviceInfo = SetupDiGetClassDevsW(&hidGuid, nullptr, nullptr,
        DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
    if (deviceInfo == INVALID_HANDLE_VALUE) {
        error = FormatWin32Error(L"SetupDiGetClassDevsW", GetLastError());
        return candidates;
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
        if (!SetupDiGetDeviceInterfaceDetailW(deviceInfo, &interfaceData, detail,
            requiredSize, nullptr, nullptr)) {
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

        wchar_t productBuffer[256] = {};
        if (attributesOk && IsSupportedProduct(attributes.ProductID)) {
            if (!HidD_GetProductString(handle, productBuffer, sizeof(productBuffer))) {
                productBuffer[0] = L'\0';
            }
        }

        if (preparsed) HidD_FreePreparsedData(preparsed);
        CloseHandle(handle);

        if (!attributesOk || attributes.VendorID != RazerOrnataController::kVendorId ||
            !IsSupportedProduct(attributes.ProductID) || !capsOk ||
            caps.FeatureReportByteLength != kFeatureReportSize) {
            continue;
        }

        Candidate candidate;
        candidate.path = path;
        candidate.productId = attributes.ProductID;
        candidate.productName = productBuffer[0] ? productBuffer :
            (IsOrnataV3X(attributes.ProductID) ? L"Razer Ornata V3 X" : L"Razer Ornata V3");
        candidates.push_back(std::move(candidate));
    }

    SetupDiDestroyDeviceInfoList(deviceInfo);
    // Match openrazer-win's Windows selection rule: a 91-byte feature report
    // identifies a usable Razer control collection. The upstream USB interface
    // index is only a preference because Windows can expose that feature report
    // on a different MI_XX collection (observed on Ornata V3 X PID 0x02A2).
    std::stable_sort(candidates.begin(), candidates.end(),
        [](const Candidate& left, const Candidate& right) {
            const bool leftPreferred = IsPreferredInterface(left.path, left.productId);
            const bool rightPreferred = IsPreferredInterface(right.path, right.productId);
            return leftPreferred && !rightPreferred;
        });
    return candidates;
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
    const auto candidates = EnumerateCandidates(error);
    if (candidates.empty()) {
        result.error = error;
        return result;
    }

    const Candidate& selected = candidates.front();
    result.detected = true;
    result.productId = selected.productId;
    result.productName = selected.productName;
    result.devicePath = selected.path;
    return result;
}

bool RazerOrnataController::Open(const DetectionResult& detection) {
    Close();
    if (!detection.detected || !IsSupportedProduct(detection.productId) ||
        detection.devicePath.empty()) {
        lastError_ = L"Rejected unverified Razer HID identity";
        return false;
    }

    HANDLE handle = OpenHidPath(detection.devicePath);
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
        attributes.ProductID == detection.productId &&
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
    productId_ = detection.productId;
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
    std::array<Rgb, kZoneCount> colors = {};
    colors.fill({ red, green, blue });
    return SetZoneColors(colors);
}

bool RazerOrnataController::SetZoneColors(const std::array<Rgb, kZoneCount>& colors) {
    if (!IsOpen()) {
        lastError_ = L"Razer HID handle is not open";
        return false;
    }
    if (haveLastFrame_ && colors == lastFrame_) return true;

    // OpenRazer's Ornata recipe: transaction 0x1F, extended matrix row 0,
    // model-specific columns (0..9 for V3, 0 for V3 X), followed by the
    // volatile custom-frame effect selector.
    auto frame = BuildCustomFrameReport(colors, ZoneCountForProduct(productId_));
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
