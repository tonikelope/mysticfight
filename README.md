# MysticFight (by tonikelope)

<img src="rgb.png" width="64" height="64" align="right" />

A lightweight Windows tool to control MSI RGB lighting based on hardware temperatures.

## Requirements
* **MysticLight_SDK.dll** (Included in the project).
* **LibreHardwareMonitor** (Must be running for temperature sensing).
* An MSI-compatible motherboard or GPU.

## Features
* **Real-time Monitoring:** Temperature tracking via WMI.
* **Customizable Thresholds:** Set your own limits for Green, Yellow, and Red states.
* **Alert Effect:** Automatic 'Lightning' style during high-temperature alerts.
* **System Tray:** Runs quietly in the background.
* **Global Hotkey:** `Ctrl + Alt + Shift + L` to toggle LEDs on/off instantly.

## Installation & Setup
1. Download the release files and keep them in the same folder.
2. Run `MysticFight.exe`.
3. Right-click the tray icon and select **Settings** to pick your temperature sensor.

---

## 🚀 How to Run at Startup (Recommended)
Since the app requires Administrator privileges to interact with the MSI SDK and WMI, the best way to start it automatically is via **Windows Task Scheduler**:

1. Open **Task Scheduler** and click **Create Task...**
2. **General Tab:** * Name it `MysticFight`.
   * Check **Run with highest privileges** (Required for SDK access).
3. **Triggers Tab:** * Click **New** and select **At log on**.
4. **Actions Tab:** * Click **New**, select **Start a program**.
   * Browse to your `MysticFight.exe`.
   * **Crucial:** In the "Start in (optional)" field, paste the path to the folder where the EXE is located (e.g., `C:\Tools\MysticFight\`). If you leave this blank, the app won't find the DLL or the config file.
5. **Settings Tab:**
   * Uncheck "Stop the task if it runs longer than...".
