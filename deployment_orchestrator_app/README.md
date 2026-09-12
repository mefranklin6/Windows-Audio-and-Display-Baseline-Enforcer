# Deployment Orchestrator app

A small Windows Flutter interface for `Deployment_Orchestrator.py`.

The app lets an administrator:

- load all feature defaults from the repository's `config.py`;
- override audio recall, display recall, BGInfo, the BGInfo folder, and desktop shortcuts for one run;
- set the maximum number of concurrently processed PCs;
- preview, edit, and save the repository's `targets.txt`, or enter targets directly without changing that file;
- switch between light and dark themes (dark is the default);
- filter deployment output by PC and load the complete timestamped log in the app;
- review per-script completion statuses in the detailed log;
- retry a deployment for an individual PC that completed with issues.
- monitor CTS deployment artifacts and installed PowerShell module versions across the selected PCs.

Feature overrides selected in the app are passed to Python for the current deployment and do not rewrite `config.py`.

Repository and Python runtime fields, the BGInfo folder, and deployment concurrency are available from the **Settings** button. Saving from the `targets.txt` editor is the only action that rewrites `targets.txt`; direct-entry deployments leave it unchanged.

Normal runs show compact visual progress for each PC. The raw output is available on demand in the **Open detailed log** modal, where its PC list can be searched and its entries can be filtered by PC. Selecting a PC also shows a separate script-status panel with success, warning, error, or fatal icons above the raw log. The detailed-log history remains available for the lifetime of the app, including logs from retries. Once a PC finishes, its progress tile includes icon-only actions to open that PC's filtered log or retry it.

After a run finishes, PCs with only informational output are shown as complete. PCs with a failed ping test are shown as blue **Offline** tiles. Script-level outcomes are appended to the detailed log, and each completed PC tile provides icon-only log and retry actions.

The **Monitoring** tab runs read-only checks through `Monitoring_Orchestrator.py` and `utility_scripts\MonitorTarget.ps1`. It reports whether `C:\ProgramData\CTS` exists, whether audio and display profiles have been saved, whether both public-desktop shortcuts exist, whether all BGInfo assets and either its standalone or consolidated startup launcher are deployed, and every installed AudioDeviceCmdlets and DisplayConfig module version. Monitoring reports can be filtered by any failed check and copied as newline-separated PC lists. The progress bar advances as each target finishes inspection; raw monitoring logs are not shown in the GUI.

## Run in development

Install Flutter with Windows desktop support and Visual Studio's **Desktop development with C++** workload. From this directory, run:

```powershell
flutter pub get
flutter run -d windows
```

The app searches parent directories for `Deployment_Orchestrator.py`. If it cannot find the repository automatically, enter the repository root in the **Repository root** field. The Python command defaults to `python`; it can also be an absolute path to `python.exe`.

## Build

```powershell
flutter build windows
```

The release bundle is created under `build\windows\x64\runner\Release`. Keep the bundle inside the repository, or use the **Repository root** field to point it to a checkout containing the Python orchestrator and installer assets.

The resulting app remains a front end to Python and PowerShell, so the deployment workstation still needs Python and the permissions/WinRM access described in the repository README.
