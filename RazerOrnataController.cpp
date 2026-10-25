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
    const std::array<RazerOrnataController::Rgb, RazerOrnataController::kZoneCount>& colors) {
    ControlReport frame = {};
    frame[1] = kTransactionId;
    frame[5] = 0x47;
    frame[6] = 0x0F;
    frame[7] = 0x03;
    frame[10] = 0x00; // row
    frame[11] = 0x00; // first zone
    frame[12] = 0x09; // last zone
    for (std::size_t i = 0; i < colors.size(); ++i) {
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
constexpr ControlReport kBlackFrameProtocolTest = BuildCustomFrameReport(kBlackProtocolTest);
constexpr ControlReport kCustomEffectProtocolTest = BuildCustomEffectReport();
static_assert(kReportSize == 90 && kFeatureReportSize == 91);
static_assert(kBlackFrameProtocolTest[5] == 0x47 &&
    kBlackFrameProtocolTest[10] == 0 && kBlackFrameProtocolTest[12] == 9 &&
    kBlackFrameProtocolTest[88] == 0x42);
static_assert(kCustomEffectProtocolTest[5] == 0x0C &&
    kCustomEffectProtocolTest[10] == 0x08 && kCustomEffectProtocolTest[88] == 0x09);

bool IsSupportedProduct(const USHORT productId) {
    return productId == RazerOrnataController::kProductIdOrnataV3 ||
        productId == RazerOrnataController::kProductIdOrnataV3Alternate;
}

std::wstring ToLower(std::wstring value) {
    std::transform(value.begin(), value.end(), value.begin(),
        [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
    return value;
}

bool IsPreferredInterface(const std::wstring& path) {
    const std::wstring lower = ToLower(path);
    return lower.find(L"&mi_02") != std::wstring::npos;
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
        // Upstream addresses both Ornata V3 PIDs on USB interface 2. Requiring
        // MI_02 as well as the exact feature size avoids selecting a keyboard
        // input collection accidentally.
        if (!IsPreferredInterface(path)) continue;

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
        candidate.productName = productBuffer[0] ? productBuffer : L"Razer Ornata V3";
        candidates.push_back(std::move(candidate));
    }

    SetupDiDestroyDeviceInfoList(deviceInfo);
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
        detection.devicePath.empty() || !IsPreferredInterface(detection.devicePath)) {
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
    customEffectActive_ = false;
    haveLastFrame_ = false;
    lastError_.clear();
    return true;
}

void RazerOrnataController::Close() {
    HANDLE handle = static_cast<HANDLE>(handle_);
    if (handle != INVALID_HANDLE_VALUE && handle != nullptr) CloseHandle(handle);
    handle_ = INVALID_HANDLE_VALUE;
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

    // OpenRazer's Ornata V3 recipe: transaction 0x1F, extended matrix row 0,
    // columns 0..9, followed by the volatile custom-frame effect selector.
    auto frame = BuildCustomFrameReport(colors);
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
    const std::array<std::uint8_t, kReportSize>& report) {
    if (report[0] != 0 || report[1] != kTransactionId || report[2] != 0 ||
        report[3] != 0 || report[4] != 0 || report[89] != 0) {
        return false;
    }

    std::uint8_t crc = 0;
    for (std::size_t i = 2; i <= 87; ++i) crc ^= report[i];
    if (crc != report[88]) return false;

    if (report[5] == 0x47 && report[6] == 0x0F && report[7] == 0x03) {
        if (report[8] != 0 || report[9] != 0 || report[10] != 0 ||
            report[11] != 0 || report[12] != 9) return false;
        for (std::size_t i = 43; i <= 87; ++i) {
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
