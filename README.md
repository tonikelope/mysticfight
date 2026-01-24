# MysticFight (by tonikelope)

<img src="rgb.png" width="64" height="64" align="right" />

A lightweight Windows tool to control MSI RGB lighting based on hardware temperatures.

## ❓ Why MysticFight?
Let's be honest: the official **Mystic Light** software is often unreliable, and "fails more than a fairground shotgun". 

The main issue is Mystic Light cpu temperature monitoring profile: **it fails constantly after a system reboot**, forcing you to manually open MSI Center and navigate to the Mystic Light section every single time just to get it working again. It seems that they fix it in one update but break it in the next, in a endless loop. This is simply intolerable. 

Fortunately, while the MSI UI is a mess, people who programmed the **SDK** actually did a good job. **MysticFight** bypasses the buggy MysticLight interface and talks directly to the SDK.

## Features
* **Real-time Monitoring:** Temperature tracking via WMI.
* **Lightweight:** No heavy dashboards; just a simple tray app that stays out of your way.
* **Customizable Temperature Sensor:** Choose the one you prefer.
* **Customizable Thresholds:** Green (Cool), Yellow (Warm), Red (Hot).
* **Alert Effect:** Automatic 'Lightning' style during high-temp alerts.
* **Night-Mode with Global Hotkey:** `Ctrl + Alt + Shift + L` to power-off/on LEDs instantly.
  
## ⚠️ Requirements
For this tool to work, you MUST have the following installed/running:

1. **MSI Center:** [Download here](https://www.msi.com/Landing/MSI-Center). (Remember you must install and enable the **Mystic Light** module inside it to provide the underlying drivers for SDK).
   * Disable **both** options in Mystic Light config: overwrite third part RGB and disable on suspend.
3. **MysticLight_SDK.dll:** Included in this repo. Must stay in the same folder as MysticFight.exe [Official Link with SDK doc](https://www.msi.com/Landing/mystic-light-rgb-gaming-pc/download).
4. **LibreHardwareMonitor:** [Download here](https://github.com/LibreHardwareMonitor/LibreHardwareMonitor). Must be running (minimized on tray) to provide temperature data via WMI.

## 🚀 How to Run at Startup
Since the app requires Administrator privileges for SDK and WMI access, use **Windows Task Scheduler**:

1. **General Tab:** Name it `MysticFight` and check **Run with highest privileges**.
2. **Triggers Tab:** Set to **At log on**.
3. **Actions Tab:** Start a program -> Select `MysticFight.exe`.
   * **IMPORTANT:** In the **"Start in (optional)"** field, paste the full path to the folder (e.g., `C:\Tools\MysticFight\`). If you leave this blank, the app won't find the DLL or your config.
4. It's recommended to **delay 30 seconds** to give the SDK time to initialise correctly.
