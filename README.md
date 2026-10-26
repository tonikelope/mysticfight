<div align="justify">
  
# MysticFight

<img src="docs/assets/rgb.png" width="64" height="64" align="right" />

Temperature-driven RGB control for MSI Mystic Light and explicitly supported Razer HID devices.

## ❓ Why MysticFight?
Let's be honest: the official **Mystic Light** software is often unreliable, and "fails more than a fairground shotgun". 

The main issue with Mystic Light CPU temperature profile: **it fails randomly after a system reboot**, forcing you to manually open MSI Center and navigate to the Mystic Light section every single time just to get it working again. It seems that they fix it in one update but break it in the next, in a endless loop. This is simply intolerable. 

Fortunately, while Mystic Light "client" is a mess, people who programmed the **SDK** actually did a good job. **MysticFight** bypasses the buggy MysticLight interface and talks directly to the SDK.

## Features
* **Real-time Monitoring:** Temperature tracking via HTTP.
* **VERY Lightweight:** No heavy dashboards, just a simple tray app that stays out of your way.
* **Customizable Device and LED area:** Choose the one you prefer. <i>Please note that MSI Mystic Light SDK may not recognise all areas available on your device (you will have to wait for MSI to update the public SDK)</i>.
* **Independent controllers per profile:** Control the MSI motherboard and any individually selected compatible Razer device. New profiles start with no RGB device selected.
* **Native Razer HID support:** Direct Windows HID control with no Synapse, Chroma Runtime, OpenRGB, extra service, or kernel driver. The validated models are Ornata V3, Ornata V3 X, Ornata V3 Tenkeyless and Tartarus Pro (exact PIDs below).
* **True multi-device Razer control:** Every compatible physical Razer HID device is shown by its reported product name and PID, and tracked internally by a stable device key. Profiles can independently select none, one, or several supported devices. Models outside the explicit compatibility table are omitted from Settings and never receive commands.
* **General or per-device thermal schemes:** Every profile has a General scheme, enabled for all selected controllers by default for full backward compatibility. Each scheme includes its data source, sensor, temperature thresholds and colours. Disable the General option to configure those values independently in the MSI and per-device Razer tabs. Unselecting a controller hides its tab without deleting its saved scheme.
* **Configurable device hand-off:** When a profile stops using an MSI or Razer device, MysticFight exits normally, or Windows commits a shutdown/restart, Advanced settings can leave it powered off or return it to an autonomous animated effect (an exact MSI SDK-advertised style, or Razer Spectrum Cycling). The default is lights off, and MysticFight never hands off a device it did not successfully control first.
* **Independent refresh rates:** MSI and every selected, supported Razer device can independently use 15, 20, or 25 FPS while consuming its assigned sensor. The Razer scheduler keeps separate smoothing and write deadlines per physical device, so different sensors, colour schemes and FPS values do not interfere. FPS limits transition speed: stable colours generate no repeated lighting writes or high-frequency rendering.
* **Customizable Data Source and Temperature Sensor:** LibreHardwareMonitor and its selected temperature sensor can be shared through General or assigned independently to each selected controller. The LibreHardwareMonitor server URL remains one global Advanced setting.
* **Customizable Temperature and Colors Thresholds:** RMS interpolation with adaptive colour-transition smoothing, shared through General or overridden per selected controller.
* **Dynamic profiles:** Five profiles are created by default, with support for up to 20. Add profiles from the **New profile** tab or delete an inactive profile with an explicit irreversible-action confirmation; the active profile and the final remaining profile are protected. Creation and deletion are saved immediately. Profile tabs scroll independently when labels no longer fit, while New profile, Shortcuts and Advanced remain pinned on the right. The New profile tab disappears at 20 profiles and returns after a deletion. Tray switching supports every profile, while configurable global shortcuts remain available for the first five positions.
* **Night-Mode:** You can set a custom global shortcut keys to power-off/on LEDs.
* **SDK auto-reset:** Automatic reset mechanism for the Mystic Light service in case of failure.
  <p align="center"><img src="docs/assets/mf_settings.png" width="500" /></p>

### Timing and idle behaviour

* With the lights enabled, LibreHardwareMonitor is still sampled at the configured 250, 500 (default), or 1000 ms interval. With the lights globally disabled, the existing low-power interval is 5 seconds.
* One `/data.json` download and one JSON parse serve every sensor required by the active profile during each polling cycle. Assigning different sensors to MSI and multiple Razer devices does not multiply HTTP requests.
* The configured 15, 20, or 25 FPS is the maximum colour-transition cadence (approximately 66, 50, or 40 ms per step), not a continuous device-write rate. Smoothing continues at that cadence until the visible RGB reaches its target.
* Each Razer device maintains its own interpolation state. Two selected devices can therefore follow different thermal thresholds and colours without coupling their frame schedules.
* Once a colour is stable, MSI and Razer stop high-frequency rendering and send no repeated lighting command. A sensor/source change, configuration change, device deadline, retry, or shutdown wakes the relevant worker immediately.
* The MSI supervisor still sends a non-lighting heartbeat after 60 seconds without SDK progress while active, or after 5 minutes while lights/MSI output are disabled. Heartbeat log messages remain throttled to one every 10 minutes.
* Razer performs read-only device detection and any throttled KVM reopen attempt on a 5-second deadline. Detection itself never sends a lighting command.
  
## ⚠️ Requirements
For this tool to work, you MUST have the following installed/running:

1. **MSI Center (only for MSI control):** [Download here](https://www.msi.com/Landing/MSI-Center). (Just installed, its not required to load with Windows. But remember you must install and enable the **Mystic Light** module inside it to provide the underlying drivers for SDK).
   * Disable **both** options in Mystic Light config: overwrite third part RGB and power saving mode. (It is highly recommended to also disable automatic updates for MSI Center and Mystic Light).
2. **LibreHardwareMonitor:** [Download here](https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases/latest). Must be running (minimized on tray) to provide temperature data via Web Server (You must enable Remote Web Server in LHM options menu). <i>Note: Yes, I know there is a LibreHardwareMonitorLib available, but I don't have the time or inclination to mess around with CLR DLL wrappers such when LHM client works perfectly well for this task, is lightweight and useful for other popular monitoring applications such as RainMeter</i>.

### Supported Razer devices

| Device | USB PID | Lighting geometry |
| --- | --- | --- |
| Razer Ornata V3 | `0x02A1` | 10 zones |
| Razer Ornata V3 (Alternate) | `0x028F` | 10 zones |
| Razer Ornata V3 X | `0x0294` | 1 zone |
| Razer Ornata V3 X (Alternate) | `0x02A2` | 1 zone |
| Razer Ornata V3 Tenkeyless | `0x02A3` | 8 zones |
| Razer Tartarus Pro | `0x0244` | 21 zones |

Compatibility means temperature-driven uniform colour plus the configurable Off/Spectrum release behaviour. It does not imply support for every feature exposed by the device. Other Razer models are deliberately omitted from Settings and are never written to until their complete command recipe, transport interface and lighting geometry have been validated.

This table is the currently implemented and validated scope, not a permanent limitation of MysticFight or Razer hardware. More Razer models can be added through explicit per-PID profiles after their required commands, HID interface, feature-report length and lighting geometry have been verified. Until such a profile exists, the model stays outside the UI and cannot receive commands.

The exact Razer protocol scope and safety decisions are documented in [RAZER_HID_PROTOCOL.md](RAZER_HID_PROTOCOL.md).

## Acknowledgements

The documented Razer HID command structures used by MysticFight were informed by the excellent reverse-engineering work of the [OpenRazer project](https://github.com/openrazer/openrazer) and cross-checked against [openrazer-win](https://github.com/Ar4ikov/openrazer-win). MysticFight does not bundle or require either project. On Windows it communicates directly with only the explicitly supported devices above.

<p align="center"><img src="docs/assets/pepinazo.jpeg"/></p>
</div>
