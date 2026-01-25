# MysticFight

<img src="rgb.png" width="64" height="64" align="right" />

MSI Mystic Light CPU Temperature Profile Replacement.

## ❓ Why MysticFight?
Let's be honest: the official **Mystic Light** software is often unreliable, and "fails more than a fairground shotgun". 

The main issue with Mystic Light cpu temperature profile: **it fails randomly after a system reboot**, forcing you to manually open MSI Center and navigate to the Mystic Light section every single time just to get it working again. It seems that they fix it in one update but break it in the next, in a endless loop. This is simply intolerable. 

Fortunately, while Mystic Light "client" is a mess, people who programmed the **SDK** actually did a good job. **MysticFight** bypasses the buggy MysticLight interface and talks directly to the SDK.

## Features
* **Real-time Monitoring:** Temperature tracking via WMI.
* **Lightweight:** No heavy dashboards; just a simple tray app that stays out of your way.
* **Customizable Temperature Sensor:** Choose the one you prefer.
* **Customizable Temperature Thresholds:** Green-Yellow-Red
* **Alert Effect:** Automatic 'Lightning' style during high-temp alerts.
* **Night-Mode with Global Hotkey:** `Ctrl + Alt + Shift + L` to power-off/on LEDs instantly.
  
## ⚠️ Requirements
For this tool to work, you MUST have the following installed/running:

1. **MSI Center:** [Download here](https://www.msi.com/Landing/MSI-Center). (Just installed, its not required to load with Windows. But remember you must install and enable the **Mystic Light** module inside it to provide the underlying drivers for SDK).
   * Disable **both** options in Mystic Light config: overwrite third part RGB and disable on suspend.
3. **MysticLight_SDK.dll:** Included in this repo. Must stay in the same folder as MysticFight.exe [Official Link with SDK doc](https://www.msi.com/Landing/mystic-light-rgb-gaming-pc/download).
4. **LibreHardwareMonitor:** [Download here](https://github.com/LibreHardwareMonitor/LibreHardwareMonitor). Must be running (minimized on tray) to provide temperature data via WMI.

## 🚀 How to Run at Startup

Since the app requires **Administrator privileges** for SDK and WMI access, you must use the **Windows Task Scheduler** instead of the standard Startup folder:

1.  **Create Task:** Open the **Task Scheduler** and select **Create Task...** (on the right panel).
2.  **General Tab:**
    * Name it `MysticFight`.
    * **CRITICAL:** Check the box **Run with highest privileges**. (Otherwise, SDK/WMI access will fail).
3.  **Triggers Tab:**
    * Click **New...** and select **At log on** in the top dropdown.
    * Under "Advanced settings", check **Delay task for:** and set it to **30 seconds**. (This gives the MSI SDK services enough time to initialize after boot).
4.  **Actions Tab:**
    * Click **New...** -> **Start a program**.
    * **Program/script:** Browse and select your `MysticFight.exe`.
    * **Start in:** ⚠️ **REQUIRED.** Paste the full path to the folder containing the exe.
        * *Example:* `C:\Tools\MysticFight\`
        * *Why?* If left blank, the app won't find `MysticLight_SDK.dll` or your configuration files.
5.  **Conditions Tab:**
    * Uncheck **Stop if the computer switches to battery power** (if you are on a laptop).
> **Quick Test:** Once created, right-click the task in the "Task Scheduler Library" and click **Run**. If the tray icon appears and the LEDs respond, the configuration is perfect.

This is practically a proof of concept for everything that the Mystic Light SDK can do with a little imagination. Carpe diem!
