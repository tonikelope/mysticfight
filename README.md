# MysticFight (by tonikelope)

<img src="rgb.png" width="64" height="64" align="right" />

A lightweight Windows tool to control MSI RGB lighting based on hardware temperatures.

## ❓ Why MysticFight?
Let's be honest: the official **MSI Center / Mystic Light** software is often bloated, unreliable, and "fails more than a fairground shotgun".

The main issue is MSI's built-in temperature monitoring script: **it fails constantly after a system reboot**, forcing you to manually open MSI Center and navigate to the Mystic Light section every single time just to get it working again. This is simply intolerable.

Fortunately, while the MSI UI is a mess, the engineers who programmed the **SDK** actually did a great job. **MysticFight** bypasses the buggy MSI interface and talks directly to the SDK, ensuring:

* **No Manual Intervention:** It works immediately after login without opening MSI Center.
* **Rock Solid:** Uses the reliable SDK backbone combined with LibreHardwareMonitor data.
* **Lightweight:** No heavy dashboards; just a simple tray app that stays out of your way.
* 
## ⚠️ Critical Requirements
For this tool to work, you MUST have the following installed:

1. **MSI Center:** [Download here](https://www.msi.com/Landing/MSI-Center). You must install the **Mystic Light** module inside it to provide the underlying drivers.
2. **MysticLight_SDK.dll:** Included in this repo. Must stay in the same folder as the EXE. [Official Link](https://www.msi.com/Landing/mystic-light-rgb-gaming-pc/download).
3. **LibreHardwareMonitor:** [Download here](https://github.com/LibreHardwareMonitor/LibreHardwareMonitor). Must be running (minimized) to provide temperature data via WMI.

## 🚀 How to Run at Startup
Since the app requires Administrator privileges for SDK and WMI access, use **Windows Task Scheduler**:

1. **General Tab:** Name it `MysticFight` and check **Run with highest privileges**.
2. **Triggers Tab:** Set to **At log on**.
3. **Actions Tab:** Start a program -> Select `MysticFight.exe`.
   * **IMPORTANT:** In the **"Start in (optional)"** field, paste the full path to the folder (e.g., `C:\Tools\MysticFight\`). If you leave this blank, the app won't find the DLL or your config.

## Features
* **Real-time Monitoring:** Temperature tracking via WMI.
* **Customizable Thresholds:** Green (Cool), Yellow (Warm), Red (Hot).
* **Alert Effect:** Automatic 'Lightning' style during high-temp alerts.
* **Global Hotkey:** `Ctrl + Alt + Shift + L` to toggle LEDs instantly.

