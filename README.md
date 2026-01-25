# MysticFight

<img src="rgb.png" width="64" height="64" align="right" />

MSI Mystic Light CPU Temperature Profile Replacement.

## ❓ Why MysticFight?
Let's be honest: the official **Mystic Light** software is often unreliable, and "fails more than a fairground shotgun". 

The main issue with Mystic Light CPU temperature profile: **it fails randomly after a system reboot**, forcing you to manually open MSI Center and navigate to the Mystic Light section every single time just to get it working again. It seems that they fix it in one update but break it in the next, in a endless loop. This is simply intolerable. 

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

## 🚀 How to Run at Startup (Step-by-Step)

Since the app requires **Administrator privileges** for SDK and WMI access, it must be launched via **Windows Task Scheduler**. Follow these precise settings:

### 1. Create the Task
* Open the **Start Menu**, type `Task Scheduler`, and press **Enter**.
* On the right panel, click **Create Task...** (Do not use "Basic Task").

### 2. Tab-by-Tab Configuration

#### **General Tab**
* **Name:** `MysticFight`
* **Security options:** Check **Run with highest privileges**. 👈 *Essential for SDK access.*
* **Configure for:** Select **Windows 10** or **Windows 11**.

#### **Triggers Tab**
1. Click **New...**
2. **Begin the task:** Select `At log on`.
3. **Advanced settings:**
   * Check **Delay task for:** and set to `30 seconds`. (Gives MSI services time to start).
   * Check **Enabled**.
4. Click **OK**.

#### **Actions Tab**
1. Click **New...** -> **Action:** `Start a program`.
2. **Program/script:** Click *Browse* and select your `MysticFight.exe`.
3. **Start in (optional):** ⚠️ **MANDATORY:** Paste the full path to the folder where the EXE is located (e.g., `C:\Tools\MysticFight\`).
   * *Note: Do not use quotes in this field.*
4. Click **OK**.

#### **Conditions Tab**
* **Power:** Uncheck **Start the task only if the computer is on AC power**.
* **Power:** Uncheck **Stop if the computer switches to battery power**.

#### **Settings Tab**
* Check **Allow task to be run on demand**. 👈 *Allows you to start it manually if needed.*
* Uncheck **Stop the task if it runs longer than**. 👈 *Critical: Otherwise Windows kills the app after 3 days.*
* Check **If the running task does not end when requested, force it to stop**.
  
> [!TIP]
> **Verification:** Once saved, right-click the task in the library and select **Run**. If the icon appears in the tray and the LEDs respond, you've nailed it.

This is practically a proof of concept for everything that the Mystic Light SDK can do with a little imagination. Carpe diem!
