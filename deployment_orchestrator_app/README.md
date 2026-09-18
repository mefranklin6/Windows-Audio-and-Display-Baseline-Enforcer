# Windows Audio and Display Baseline Enforcer Orchestrator

## Install and run the released app

Download `Windows-Audio-and-Display-Baseline-Enforcer-Orchestrator-<version>-Setup.exe` from the [latest GitHub Release](https://github.com/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer/releases/latest), then run it. The installer includes the compiled application, Flutter runtime, and PowerShell/support files. A repository checkout, Python, and Flutter are not required.

The installer places the application and required Flutter runtime files in `Program Files\Windows Audio and Display Baseline Enforcer Orchestrator`. Scripts, targets, BGInfo assets, logs, and settings are stored in `%APPDATA%\Windows Audio and Display Baseline Enforcer Orchestrator`. It creates a Start menu entry and can create a desktop shortcut. Use the update button in the application toolbar to compare the installed version with the latest GitHub Release and open the newer installer.

## Run from source

Install the Flutter SDK with Windows desktop support and Visual Studio with the **Desktop development with C++** workload. Clone the repository, then run the application from this directory:

```powershell
git clone https://github.com/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer.git
cd .\Windows-Audio-and-Display-Baseline-Enforcer\deployment_orchestrator_app
flutter pub get
flutter run -d windows
```

Source builds search parent directories for `installer_scripts` and `utility_scripts\MonitorTarget.ps1`. If the application cannot find them, set **Application files** in Settings to the repository root.

## Build a local release

From this directory:

```powershell
flutter pub get
flutter build windows --release
```

The raw Windows bundle is created at `build\windows\x64\runner\Release`. Keep every generated file in that directory together; `Windows-Audio-and-Display-Baseline-Enforcer-Orchestrator.exe` alone is not portable. To run that bundle, keep it within the repository or place `installer_scripts`, `utility_scripts`, and `BGInfo` beside it.

To make the same single-file installer used by CI, install [Inno Setup](https://jrsoftware.org/isinfo.php) and run this from `deployment_orchestrator_app` after building:

```powershell
$repoRoot = (Resolve-Path ..).Path
$releaseDir = (Resolve-Path .\build\windows\x64\runner\Release).Path
iscc.exe "/DAppVersion=1.0.0" "/DSourceDir=$releaseDir" "/DRepoRoot=$repoRoot" "/DOutputDir=$repoRoot\dist" ..\installer\windows-setup.iss
```

Replace `1.0.0` with the version being built. The installer is written to `dist` at the repository root.

## Develop and test

Fetch dependencies and launch the app with hot reload:

```powershell
flutter pub get
flutter run -d windows
```

Before committing changes, format and validate the code:

```powershell
dart format lib test
flutter analyze
flutter test
```

## Publish a release

The [Windows build workflow](../.github/workflows/build-windows.yml) runs on pushed tags in the form `vMAJOR.MINOR.PATCH` or `MAJOR.MINOR.PATCH`. It tests the project, builds the Windows application, packages the installer with Inno Setup, creates or updates the matching GitHub Release, and uploads both the installer and its SHA-256 checksum.

```powershell
git tag v1.0.0
git push origin v1.0.0
```

## Runtime requirements and files

The deployment workstation needs Windows PowerShell plus the administrative permissions, WinRM connectivity, and administrative-share access required by the root [README](../README.md). Target computers need internet access when the deployment scripts install pinned PowerShell modules from GitHub.

Target files must be UTF-8 text with one hostname per line. Blank lines and lines beginning with `#` are allowed. Settings are stored in `%APPDATA%\Windows Audio and Display Baseline Enforcer Orchestrator\settings.json`; timestamped deployment and monitoring logs are written to `logs` under the application-files directory.

For BGInfo deployments, select any accessible folder containing exactly one `BGInfo64.exe`, one `.bgi` configuration file, and one supported background image (`.jpg`, `.jpeg`, `.png`, `.bmp`, or `.gif`).
