# Windows-Audio-and-Display-Baseline-Enforcer

A deterministic audio and display baseline enforcement system for shared Windows PCs.

## Overview

### Target Audience

- Administrators of shared computers installed in conference rooms, classrooms, or similar environments.
- Admins or owners of computers installed as part of complex AV systems, including home theaters.
- Admins or owners of computers used as kiosks or digital signage.

### Goals

This system is designed to keep audio and display configuration stable and predictable.

- Prevent Windows from selecting the wrong audio devices or display configuration.
- Prevent Windows from "guessing" or defaulting to last connected device/configuration state by deterministically recalling known-good settings at login and logout.
- Prevent AV configuration drift caused by Windows updates.
- Allow users to temporarily make custom audio/display changes, then restore a standard configuration for the next user.

Windows and users often change audio and display configuration in shared environments. This system restores a known-good baseline at login and logout.

> Optional: This system can also deploy Sysinternals [BGInfo](https://learn.microsoft.com/en-us/sysinternals/downloads/bginfo) to write information over the desktop wallpaper at login. This is useful for showing who is logged in and for displaying asset or service tag information that helps users submit support tickets.

### App-Based Deployment and Monitoring

Use the included Windows App to deploy and monitor this project with minimal permissions, dependencies, and hassle. No SCCM, Intune, programming languages, or other device management systems are needed.

![image of app](/images/app_image.png)

### Modular Architecture

This system is modular, so you can choose which features and installers to deploy. You can either directly run scripts under `\installer_scripts` , or use the recommended Windows Flutter app to install and monitor a selected set of scripts across multiple computers.

### Requirements

If you plan on deploying remotely, make sure your workstation has the proper permissions and WinRM is working by running the below in PowerShell:

`Test-WSMan -ComputerName <name of a remote PC>`

## Quickstart: Precompiled Install Wizard (recommended)

1. Open the [latest GitHub Release](https://github.com/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer/releases/latest).
2. Download and run `Windows-Audio-and-Display-Baseline-Enforcer-Orchestrator-<version>-Setup.exe`

***...and that's it! If this method works for you, feel free to stop reading here and start using the app.***

Note: This method will likely trigger a Windows Smart Screen warning, but you can safely proceed to run the program. If you don't trust the .exe or if you want to bypass Smart Screen altogether, you can follow the instructions below to compile your own exe from source.

### Compile the app yourself (optional)

Use this path when you want a local release build from source. It requires Git, the Flutter SDK, Visual Studio with the Desktop development with C++ workload, and a clone of this repository.

```powershell
git clone https://github.com/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer.git
cd .\Windows-Audio-and-Display-Baseline-Enforcer\deployment_orchestrator_app
flutter pub get
flutter build windows --release
```

The compiled bundle is written to `deployment_orchestrator_app\build\windows\x64\runner\Release`. Keep `Windows-Audio-and-Display-Baseline-Enforcer-Orchestrator.exe`, the DLLs, the `data` directory, and the other generated files together. Also keep the bundle inside the repository or place `installer_scripts`, `utility_scripts`, and `BGInfo` beside it; the Flutter executable alone is not a portable build.

Official tagged builds use [the Windows build workflow](.github/workflows/build-windows.yml) and [the Inno Setup definition](installer/windows-setup.iss) to package the complete bundle and PowerShell files into one downloadable installer.

### Develop and test the app

Use this path when changing the Dart UI, orchestration code, PowerShell scripts, or tests. After cloning the repository and installing the build prerequisites, run:

```powershell
cd .\deployment_orchestrator_app
flutter pub get
flutter run -d windows
```

Before submitting changes, format and validate the project:

```powershell
dart format lib test
flutter analyze
flutter test
```

Development builds locate `installer_scripts` and `utility_scripts` from the repository automatically. See the [app README](deployment_orchestrator_app/README.md) for additional implementation and runtime details.

## Individual Script Deployment Method

This method is best used for testing or very small deployments. With this method, there is no additional setup or dependencies required. Simply run the PowerShell script you need from the `installer_scripts` folder. When prompted for a target PC, provide a remote hostname or `localhost` to install locally.

It also may be possible to load or modify these scripts for use in Intune or other management programs.

See [Individual Scripts](#individual-scripts) for details about what each script does.

## Individual Scripts

The below can either be deployed by the main deployment script, or can be run individually against either the localhost or a remote PC with WinRM.

### Audio Device Cmdlets Installer

Installs a custom fork of [AudioDeviceCmdlets](https://github.com/mefranklin6/AudioDeviceCmdlets), which is maintained specifically for this system. (Thank you to all of those who made AudioDeviceCmdlets possible!)

It performs the following:

1. Installs the fork of AudioDeviceCmdlets.
2. Installs a local script to save audio settings (I/O devices, recording/playback volume, and mute states).
3. Installs a startup script that recalls saved settings at user login. If no settings are saved, no action is taken.
4. Writes a log to `C:\ProgramData\CTS\AudioDeviceStartup.log`.

#### AudioDeviceCmdlets Usage

If you are installing multiple tools and running `Cleanup.ps1` last, continue to the [Cleanup section](#cleanup-script). If you are only installing AudioDeviceCmdlets, the installer places `SAVE_AUDIO_SETTINGS.bat` on the desktop. Configure audio devices/levels as desired, then run that file to save the configuration.

### DisplayConfig Installer

Installs a fork of [`DisplayConfig`](https://github.com/mefranklin6/DisplayConfig), created by [MartinGC94](https://github.com/MartinGC94). The fork is maintained specifically for this system.

It performs the following:

1. Installs a pinned `DisplayConfig` module version from GitHub release assets.
2. Copies display save/recall scripts to `C:\ProgramData\CTS`.
3. Installs a startup launcher at `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\cts_display_startup.bat`.
4. Prints a prompt to run the display profile save script after installation.

#### DisplayConfig Usage

If you are installing multiple tools and running `Cleanup.ps1` last, continue to the [Cleanup section](#cleanup-script), otherwise follow the below steps:

- After install, run `C:\ProgramData\CTS\run_display_config_save_profile.bat` on the target machine to capture a known-good display profile.
- On future logins, the startup script recalls that saved display profile.

### BGInfo Installer

This installer deploys Sysinternals [BGInfo](https://learn.microsoft.com/en-us/sysinternals/downloads/bginfo) assets from the repo and configures startup execution.

It performs the following:

1. Copies the discovered BGInfo executable, config, and background image to `C:\ProgramData\CTS`.
2. Installs startup launcher `cts_bginfo_startup.bat` in the Windows Startup folder.
3. Replaces the BGInfo config each run to keep the deployed profile current.

#### BGInfo Usage

- Place the required BGInfo assets in a directory under `BGInfo`, then select that directory in the app.
- That folder must contain exactly one `BGInfo64.exe`, exactly one `.bgi` file, and exactly one supported image file (`.jpg`, `.jpeg`, `.png`, `.bmp`, or `.gif`). You can keep multiple folders for different backgrounds and styles, but only the folder selected in the app is deployed.
- Include `InstallBGInfo.ps1` in `pwsh_scripts` only on systems where you want BGInfo applied at login.

![bginfo desktop example](/images/bginfo.png)

### Cleanup Script

If more than one script was executed per machine, make sure `Cleanup.ps1` runs after the installers. This script consolidates installer artifacts into a single `SAVE_AV_SETTINGS.bat` file on the Public Desktop to save both audio and display settings, and creates one optimized startup script to recall saved settings in the proper order. This script will also apply your BGInfo settings if specified.

The `SAVE_AV_SETTINGS.bat` file is placed on the Public Desktop, requires admin rights, and self-destructs after running. For edits or reruns, a persistent copy is stored in `C:\ProgramData\CTS`.

#### Cleanup Results

- Consolidates separate startup launchers into one ordered startup batch file.
- Preserves display recall before audio recall and BGInfo execution.
- Removes the standalone `SAVE_AUDIO_SETTINGS.bat` desktop file when settings are consolidated into `SAVE_AV_SETTINGS.bat`.
- Stores persistent recall-first launchers batch files and SAVE_AV_SETTINGS.bat in `C:\ProgramData\CTS`.

### Shortcut Installer Script

Adds `Log Out` and `Reboot` shortcuts to the public desktop, which recall proper AV settings before proceeding. These shortcuts requires all of the above scripts except the BGInfo script to already be installed.

### Uninstall Script

Removes all shortcuts, cmdlets, files, and settings from the target PC.

## Notes

- The startup script is fast and lightweight, but Windows may take several seconds after login to execute Startup-folder items. Users may also briefly see a blank command prompt window (which is immediately minimized) before the saved settings are applied.
- It is best practice to hide the power options in the start menu and direct users to the `Reboot` and `Log Out` desktop shortcuts so AV settings are recalled at logout.
- If using the precompiled .exe, Windows may show a Smart Screen warning. If you don't trust the executable, you can bypass this warning and [compile the program yourself from source](#compile-the-app-yourself-optional).

## AI Disclosure

The Powershell scripts that are the core of the backend, and the old Python orchestrator files that became the basis for the app, were either developed before AI became useful, or AI was used as a tool with strict human review. These are the files that actually modify the PC's, and they have been thoroughly tested and reviewed. These systems have been in-production without issue.

The new front-end GUI, or 'App' was almost entirely 'vibe coded', but tested, and the parts that touch anything important were reviewed manually. Front-end, aesthetic, and UX elements are developed quickly by prompting AI, and these less important aspects are not as strictly reviewed. AI has also developed integration tests for changes to the GUI.

## Release Changelog

### v3.1.1

18 September 2026

- Allow for setting the BGInfo folder from any valid location

### v3.1.0

17 September 2026

- Refinements to the installation and uninstallation process
- Consistent use of naming in the app

### v3.0.0b

17 September 2026 Beta

- Major changes: Configuration, deployment, and monitoring are now achieved through a GUI app, significantly reducing complexity and hassle on administrators and reducing the barrier to entry.

New features that the app brings:

- One-click installation and configuration
- Monitoring with built-in retry
- Batch Uninstalling
- Development tests
- Update checking
- Log filtering
- Reporting
- Ability to use multiple target files

### v2.0.3

8 June 2026

- Add uninstaller utility script
- Internal refactor: desktop shortcut creation is its' own script now.
- Update the readme, including changes made previously

### v2.0.1

1 June 2026

- Bugfix error in BGInfo recall script
- Change default value of BGINFO_INSTALL to False in config example and add context.

### v2.0.0

4 May 2026

- Backwards incompatible change: Move feature flags and script-specific parameters to a shared configuration file (removed in a later app release).
- Bump DisplayConfig version to 6.0.1
- Modify `Reboot` and `Log Out` shortcuts to instantly display the user a message saying they will be logged out shortly.

### v1.1.1

10 April 2026

- Bug fix: record null value for audio levels if device does not exist (like no recording device)
- Harden: Elevate audio save script to admin if not already
- Harden: audio level recall. Check if value is non-numeric first.

### v1.1.0

27 March 2026

- Added changelog
- New feature: Add `Reboot` and `Log Out` shortcuts, which recall saved Audio and Display settings first.

### v1.0.0

26 March 2026

- Initial feature-complete release
