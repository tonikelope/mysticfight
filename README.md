# MysticFight (by tonikelope)

<img src="rgb.png" width="64" height="64" align="right" />

A lightweight Windows tool to control MSI RGB lighting based on hardware temperatures.

## ❓ Why MysticFight?
Let's be honest: the official **MSI Center / Mystic Light** software is often bloated, unreliable, and "fails more than a fairground shotgun" (*falla más que una escopeta de feria*). It frequently resets your color profiles after a reboot, gets stuck, or consumes unnecessary system resources just to show a simple temperature color.

I built **MysticFight** to solve this:
* **Lightweight:** It uses almost zero CPU/RAM compared to the heavy MSI suite.
* **Reliability:** It won't "forget" your colors or lose the connection after a restart.
* **Smart:** It does one thing perfectly—keeping your hardware safe and your RGB synced.
* **Stealth:** No annoying splash screens or dashboards; it lives in your system tray.

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

