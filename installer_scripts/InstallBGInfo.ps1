Param(
    [Parameter(Mandatory = $true)]
    [string]$PC,
    [Parameter(Mandatory = $false)]
    [string]$standalone = "true"
)

# Workaround for passing in booleans from Python
if ($standalone -eq "true") {
    $standalone = $true
}
else {
    $standalone = $false
}

$year = "26_27" #TODO: Change every year.

Import-Module (Join-Path $PSScriptRoot 'shared\SharedHelpers.psm1') -Force

# Check if running locally
$isLocal = Test-IsLocalComputer -ComputerName $PC

Write-Output "$PC Installing BGInfo and scripts"

try {

    if ($standalone -and -not $isLocal) {
        if (-not (Test-HostReachable -ComputerName $PC -TimeoutMilliseconds 1000)) {
            throw "$PC is not reachable"
        }
    }


    # Create C:\ProgramData\CTS folder if it doesn't exist
    if ($standalone) {
        Invoke-LocalOrRemote -ComputerName $PC -IsLocal $isLocal -ScriptBlock {
            if (!(Test-Path -Path 'C:\ProgramData\CTS')) {
                New-Item -ItemType Directory -Path 'C:\ProgramData\CTS' -Force | Out-Null
            }
        } | Out-Null
    }


    if ($isLocal) { $prefix = "C:\" }
    else { $prefix = "\\$PC\C$\" }

    # --- Repo source paths ---
    $repo_bginfo_dir = Join-Path $PSScriptRoot "..\BGInfo\$year"

    $repo_exe_path = Join-Path $repo_bginfo_dir "Bginfo64.exe"
    $repo_config_path = Join-Path $repo_bginfo_dir "bg_$year.bgi"
    $repo_bat_path = Join-Path $repo_bginfo_dir "run_bgi_$year.bat"
    $repo_image_path = Join-Path $repo_bginfo_dir "desktop_$year.jpg"

    # Fail fast if repo assets are missing
    $repoRequired = @(
        @{ Path = $repo_exe_path; Name = "Bginfo64.exe (repo)" },
        @{ Path = $repo_config_path; Name = "bg_$year.bgi (repo)" },
        @{ Path = $repo_bat_path; Name = "run_bgi_$year.bat (repo)" },
        @{ Path = $repo_image_path; Name = "desktop_$year.jpg (repo)" }
    )
    foreach ($f in $repoRequired) {
        if (-not (Test-Path -LiteralPath $f.Path)) {
            throw "$PC Missing required repo file: $($f.Name) at $($f.Path)"
        }
    }

    # --- Destination paths on target ---
    $local_exe_path_suffix = "ProgramData\CTS\Bginfo64.exe"
    $local_config_path_suffix = "ProgramData\CTS\bg_$year.bgi"
    $local_bat_path_suffix = "ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\cts_bginfo_startup.bat"
    $local_image_path_suffix = "ProgramData\CTS\desktop_$year.jpg"

    $local_exe_path = Join-Path $prefix $local_exe_path_suffix
    $local_config_path = Join-Path $prefix $local_config_path_suffix
    $local_bat_path = Join-Path $prefix $local_bat_path_suffix
    $local_image_path = Join-Path $prefix $local_image_path_suffix

    # Copy image (force update to ensure correct year)
    Copy-Item -Path $repo_image_path -Destination $local_image_path -Force

    # Copy exe (only if missing)
    if (!(Test-Path $local_exe_path)) {
        Copy-Item -Path $repo_exe_path -Destination $local_exe_path -Force
    }

    # Replace config every time
    if (Test-Path $local_config_path) { Remove-Item $local_config_path -Force }
    Copy-Item -Path $repo_config_path -Destination $local_config_path -Force


    # This script owns its own Startup bat. Consolidation/ordering is handled by Cleanup.ps1.
    Copy-Item -Path $repo_bat_path -Destination $local_bat_path -Force

    # Verify Files
    $requiredFiles = @(
        @{ Path = $local_exe_path; Name = "Bginfo64.exe" },
        @{ Path = $local_config_path; Name = "bg_$year.bgi" },
        @{ Path = $local_bat_path; Name = "cts_bginfo_startup.bat" },
        @{ Path = $local_image_path; Name = "desktop_$year.jpg" }
    )

    foreach ($file in $requiredFiles) {
        if (Test-Path $file.Path) {
            continue
        }
        else {
            throw "$PC $($file.Name) not found at $($file.Path)"
        }
    }

    Write-Output "$PC BGInfo installation complete"



} # end try
catch {
    Write-Output "$PC InstallBGInfo failed: $_"
    Exit 1
}

Start-Sleep 1
Exit 0