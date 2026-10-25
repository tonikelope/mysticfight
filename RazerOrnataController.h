#pragma once

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>

class RazerOrnataController {
public:
    static constexpr std::uint16_t kVendorId = 0x1532;
    static constexpr std::uint16_t kProductIdOrnataV3 = 0x02A1;
    static constexpr std::uint16_t kProductIdOrnataV3Alternate = 0x028F;
    static constexpr std::size_t kZoneCount = 10;

    struct DetectionResult {
        bool detected = false;
        std::uint16_t productId = 0;
        std::wstring productName;
        std::wstring devicePath;
        std::wstring error;
    };

    struct Rgb {
        std::uint8_t red = 0;
        std::uint8_t green = 0;
        std::uint8_t blue = 0;

        bool operator==(const Rgb&) const = default;
    };

    RazerOrnataController() = default;
    ~RazerOrnataController();

    RazerOrnataController(const RazerOrnataController&) = delete;
    RazerOrnataController& operator=(const RazerOrnataController&) = delete;

    // Read-only HID enumeration. This method never sends a feature report.
    static DetectionResult Detect();

    bool Open(const DetectionResult& detection);
    void Close();
    bool IsOpen() const;

    // The only write operations exposed by this class. Both use OpenRazer's
    // volatile custom-frame commands; no firmware, bootloader or persistent
    // storage commands are implemented.
    bool SetUniformColor(std::uint8_t red, std::uint8_t green, std::uint8_t blue);
    bool SetZoneColors(const std::array<Rgb, kZoneCount>& colors);

    const std::wstring& LastError() const { return lastError_; }

private:
    bool SendReport(std::array<std::uint8_t, 90>& report);
    static bool IsWhitelistedReport(const std::array<std::uint8_t, 90>& report);

    void* handle_ = reinterpret_cast<void*>(static_cast<std::intptr_t>(-1));
    bool customEffectActive_ = false;
    bool haveLastFrame_ = false;
    std::array<Rgb, kZoneCount> lastFrame_{};
    std::wstring lastError_;
};
