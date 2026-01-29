# MysticFight

<img src="rgb.png" width="64" height="64" align="right" />

MSI Mystic Light CPU Temperature Profile Replacement.

## ❓ Why MysticFight?
Let's be honest: the official **Mystic Light** software is often unreliable, and "fails more than a fairground shotgun". 

The main issue with Mystic Light CPU temperature profile: **it fails randomly after a system reboot**, forcing you to manually open MSI Center and navigate to the Mystic Light section every single time just to get it working again. It seems that they fix it in one update but break it in the next, in a endless loop. This is simply intolerable. 

Fortunately, while Mystic Light "client" is a mess, people who programmed the **SDK** actually did a good job. **MysticFight** bypasses the buggy MysticLight interface and talks directly to the SDK.

## Features
* **Real-time Monitoring:** Temperature tracking via WMI.
* **VERY Lightweight:** No heavy dashboards; just a simple tray app that stays out of your way.
* **Customizable Temperature Sensor:** Choose the one you prefer.
* **Customizable Temperature and Colors Thresholds:** LERP RGB algorithm.
* **Night-Mode with Global Hotkey:** `Ctrl + Alt + Shift + L` to power-off/on LEDs instantly.
  
## ⚠️ Requirements
For this tool to work, you MUST have the following installed/running:

1. **MSI Center:** [Download here](https://www.msi.com/Landing/MSI-Center). (Just installed, its not required to load with Windows. But remember you must install and enable the **Mystic Light** module inside it to provide the underlying drivers for SDK).
   * Disable **both** options in Mystic Light config: overwrite third part RGB and power saving mode.
3. **Mystic Light SDK:** MysticLight_SDK.dll is included inside MysticFight.exe
4. **LibreHardwareMonitor 0.9.4:** [Download here](https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases/download/v0.9.4/LibreHardwareMonitor-net472.zip). Must be running (minimized on tray) to provide temperature data via WMI. (Yes, i know there is a LibreHardwareMonitorLib available, but I don't have the time or inclination to mess around with CLR DLL wrappers such when WMI works perfectly well for this task and LHM client is lightweight and useful for other monitoring applications such as RainMeter).

This is practically a proof of concept for everything that the Mystic Light SDK can do with a little imagination. Carpe diem!
