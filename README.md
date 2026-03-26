# Windows-Audio-and-Display-Baseline-Enforcer

A deterministic audio and display baseline enforcement system for shared Windows PCs.

## Target Audience

- Administrators of shared computers installed in conference rooms, classrooms, or similar environments.
- Admins or owners of computers installed as part of complex AV systems, including home theaters.
- Admins or owners of computers used as kiosks or digital signage.

## Goals

This system is designed to keep audio and display configuration stable and predictable.

- Prevent Windows from selecting the wrong audio devices or display configuration.
- Prevent Windows from "guessing" or defaulting to last connected device/configuration state by deterministically recalling known-good settings at login.
- Prevent AV configuration drift caused by Windows updates.
- Allow users to temporarily make custom audio/display changes, then restore a standard configuration for the next user.

Windows and users often change audio and display configuration in shared environments. This system restores a known-good baseline at login.

> Optional: This system can also deploy Sysinternals [BGInfo](https://learn.microsoft.com/en-us/sysinternals/downloads/bginfo) to write information over the desktop wallpaper at login. This is useful for showing who is logged in and for displaying asset or service tag information that helps users submit support tickets.

## Usage and Requirements

Both the deployment script and individual scripts can target local or remote PC's via WinRM. If you plan on deploying remotely, make sure your workstation has the proper permissions and WinRM is working:

`Test-WSMan -ComputerName <name of a remote PC>`

If you plan to run individual scripts directly, no additional setup is required. Run the PowerShell script you need from the `installer_scripts` folder. When prompted for a target PC, provide a remote hostname or `localhost`.

If you plan to deploy to multiple PCs at once:

- Install Python 3.14 or later on your workstation. This project currently uses only the standard library.
- Follow the [Remote Deploy](#remote-deployment-script) instructions below.

If you plan to deploy BGInfo, place the latest `BGInfo64.exe`, one `.bgi` file, and one background image in the folder configured by `$folder` in `InstallBGInfo.ps1`. The script scans `BGInfo\<folder>` and requires exactly one match for each asset type.

## Modular Architecture

This system is modular, so you can choose which features and installers to deploy. You can either run scripts under `installer_scripts` directly, or use `00_remote_deploy.py` to install a selected set of scripts across multiple computers.

## Remote Deployment Script

The deploy script processes multiple target PCs concurrently. For each target PC, it runs the selected installer scripts in the order listed in `pwsh_scripts`. For normal deployments, this is the only file you need to run.

### Usage

0. Clone this repository to your admin workstation. The 'main' branch is the most up-to-date but may not always be fully tested, so you may wish to use a [release version](https://github.com/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer/releases). You can find a changelog at the end of this readme.
1. Create `targets.txt` in the repository root (use `targets.txt.example` as a reference).
2. Add one PC target per line in `targets.txt`.
3. Edit the `pwsh_scripts` list in `00_remote_deploy.py` to include only what you want to deploy.

    Example (BGInfo skipped):

    ```python
    pwsh_scripts = [
        "installer_scripts\\./InstallAudioDeviceCmdlets.ps1",
        "installer_scripts\\./InstallDisplayConfig.ps1",
        # "installer_scripts\\./InstallBGInfo.ps1",
        "installer_scripts\\./Cleanup.ps1",  # Cleanup must be last
    ]
    ```

Run:

```powershell
cd <to your repo root>
python .\00_remote_deploy.py
```

#### Deployment Notes

- Each line in `targets.txt` should contain one hostname.
- Keep `Cleanup.ps1` last in the `pwsh_scripts` list.
- Remove or comment out installers you do not want to ship in your environment.

## Individual Scripts

The below can either be deployed by the main deployment script, or can be run individually against either the localhost or a remote PC with WinRM.

### Audio Device Cmdlets Installer

Installs a custom fork of [AudioDeviceCmdlets](https://github.com/mefranklin6/AudioDeviceCmdlets), which is maintained specifically for this system. (Thank you to all of those who made AudioDeviceCmdlets possible!)

It performs the following:

1. Installs the fork of AudioDeviceCmdlets.
2. Installs a local script to save audio settings (I/O devices, recording volume, playback volume).
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

- Place the required BGInfo assets in `BGInfo\<folder>`, where `<folder>` matches the `$folder` value in `InstallBGInfo.ps1`.
- That folder must contain exactly one `BGInfo64.exe`, exactly one `.bgi` file, and exactly one supported image file (`.jpg`, `.jpeg`, `.png`, `.bmp`, or `.gif`).
- Include `InstallBGInfo.ps1` in `pwsh_scripts` only on systems where you want BGInfo applied at login.

### Cleanup Script

If more than one script was executed per machine, make sure `Cleanup.ps1` runs last. This script consolidates installer artifacts into a single `SAVE_AV_SETTINGS.bat` file on the Public Desktop to save both audio and display settings, and creates one optimized startup script to recall saved settings in the proper order. This script will also apply your BGInfo settings if specified.

The `SAVE_AV_SETTINGS.bat` file is placed on the Public Desktop, requires admin rights, and self-destructs after running. For edits or reruns, a persistent copy is stored in `C:\ProgramData\CTS`.

#### Cleanup Results

- Consolidates separate startup launchers into one ordered startup batch file.
- Preserves display recall before audio recall and BGInfo execution.
- Removes the standalone `SAVE_AUDIO_SETTINGS.bat` desktop file when settings are consolidated into `SAVE_AV_SETTINGS.bat`.
- Creates `Log Out` and  `Reboot` shortcuts on the Public Desktop, which recall the AV settings first.
- Stores persistent recall-first launchers batch files and SAVE_AV_SETTINGS.bat in `C:\ProgramData\CTS`.

## Notes

The startup script is fast and lightweight, but Windows may take several seconds after login to execute Startup-folder items. Users may also briefly see a blank command prompt window (which is immediately minimized) before the saved settings are applied.

## Changelog

### v1.0.0

26 March 2026

- Initial feature-complete release

### v1.1.0

27 March 2026

- Added changelog
- New feature: Add `Reboot` and `Log Out` shortcuts, which recall saved Audio and Display settings first.
