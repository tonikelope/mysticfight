#pragma once

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

class RazerOrnataController {
public:
    static constexpr std::uint16_t kVendorId = 0x1532;
    static constexpr std::uint16_t kProductIdOrnataV3 = 0x02A1;
    static constexpr std::uint16_t kProductIdOrnataV3Alternate = 0x028F;
    static constexpr std::uint16_t kProductIdOrnataV3X = 0x0294;
    static constexpr std::uint16_t kProductIdOrnataV3XAlternate = 0x02A2;
    static constexpr std::uint16_t kProductIdOrnataV3Tenkeyless = 0x02A3;
    static constexpr std::uint16_t kProductIdTartarusPro = 0x0244;
    static constexpr std::size_t kMaxZoneCount = 21;

    struct DetectedDevice {
        std::wstring deviceKey;
        std::uint16_t productId = 0;
        std::wstring productName;
        std::wstring controlPath;
        bool controllable = false;

        bool operator==(const DetectedDevice&) const = default;
    };

    struct DetectionResult {
        bool detected = false;
        std::uint16_t productId = 0;
        std::wstring productName;
        std::wstring devicePath;
        std::wstring error;
        std::vector<DetectedDevice> devices;
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

    bool Open(const DetectedDevice& device);
    void Close();
    bool IsOpen() const;

    // Temperature output uses only OpenRazer's volatile custom-frame commands.
    // Persistent writes are limited to the two documented release states used
    // when MysticFight relinquishes a device it actually controlled.
    bool SetUniformColor(std::uint8_t red, std::uint8_t green, std::uint8_t blue);
    bool SetZoneColors(const std::array<Rgb, kMaxZoneCount>& colors);
    // Explicitly applies the documented persistent NONE (lights off) effect.
    bool SetLightsOff();
    // Explicitly restores the documented persistent Spectrum Cycling effect.
    bool SetSpectrumCycling();

    const std::wstring& LastError() const { return lastError_; }

private:
    bool SendReport(std::array<std::uint8_t, 90>& report);
    bool IsWhitelistedReport(const std::array<std::uint8_t, 90>& report) const;

    void* handle_ = reinterpret_cast<void*>(static_cast<std::intptr_t>(-1));
    std::uint16_t productId_ = 0;
    bool customEffectActive_ = false;
    bool haveLastFrame_ = false;
    std::array<Rgb, kMaxZoneCount> lastFrame_{};
    std::wstring lastError_;
};
