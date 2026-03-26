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

Import-Module (Join-Path $PSScriptRoot 'shared\SharedHelpers.psm1') -Force

$isLocal = Test-IsLocalComputer -ComputerName $PC

Write-Output "INFO: Installing AudioDeviceCmdlets"

$AudioDeviceCmdletsVersion = '3.3'

$AudioDeviceCmdletsModuleName = 'AudioDeviceCmdlets'
$AudioDeviceCmdletsZipUrl = "https://github.com/mefranklin6/AudioDeviceCmdlets/releases/download/v$AudioDeviceCmdletsVersion/AudioDeviceCmdlets-$AudioDeviceCmdletsVersion.zip"

try {

    if ($standalone -and -not $isLocal) {
        if (-not (Test-HostReachable -ComputerName $PC -TimeoutMilliseconds 1000)) {
            throw "$PC is not reachable"
        }
    }

    # Install AudioDeviceCmdlets from a pinned GitHub release ZIP
    Invoke-LocalOrRemote -ComputerName $PC -IsLocal $isLocal -ArgumentList @(
        $AudioDeviceCmdletsZipUrl,
        $AudioDeviceCmdletsModuleName,
        $AudioDeviceCmdletsVersion
    ) -ScriptBlock {
        param(
            [Parameter(Mandatory = $true)]
            [string]$ZipUrl,
            [Parameter(Mandatory = $true)]
            [string]$ModuleName,
            [Parameter(Mandatory = $true)]
            [string]$ModuleVersion
        )

        $ErrorActionPreference = 'Stop'
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        function Get-WritableModuleRoot {
            param(
                [Parameter(Mandatory = $true)]
                [string[]]$Candidates
            )

            foreach ($candidate in $Candidates) {
                if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
                $p = $candidate.Trim()
                try {
                    $null = New-Item -ItemType Directory -Path $p -Force -ErrorAction Stop
                    $probe = Join-Path $p ("_probe_" + [Guid]::NewGuid().ToString('N'))
                    $null = New-Item -ItemType Directory -Path $probe -Force -ErrorAction Stop
                    Remove-Item -LiteralPath $probe -Recurse -Force -ErrorAction SilentlyContinue
                    return $p
                }
                catch {
                    continue
                }
            }

            return $null
        }

        function Remove-OldModuleVersions {
            param(
                [Parameter(Mandatory = $true)]
                [string[]]$ModulePaths,
                [Parameter(Mandatory = $true)]
                [string]$ModuleName,
                [Parameter(Mandatory = $true)]
                [string]$TargetVersion
            )

            foreach ($modulePath in ($ModulePaths | Select-Object -Unique)) {
                if ([string]::IsNullOrWhiteSpace($modulePath)) {
                    continue
                }

                $moduleRoot = Join-Path $modulePath $ModuleName
                if (-not (Test-Path -LiteralPath $moduleRoot)) {
                    continue
                }

                $versionDirectories = Get-ChildItem -LiteralPath $moduleRoot -Directory -ErrorAction SilentlyContinue
                foreach ($versionDirectory in $versionDirectories) {
                    if ($versionDirectory.Name -eq $TargetVersion) {
                        continue
                    }

                    Remove-Item -LiteralPath $versionDirectory.FullName -Recurse -Force -ErrorAction Stop
                    Write-Output "INFO: Removed old $ModuleName version $($versionDirectory.Name) from $moduleRoot"
                }

                $legacyManifestPath = Join-Path $moduleRoot "${ModuleName}.psd1"
                if (-not (Test-Path -LiteralPath $legacyManifestPath)) {
                    continue
                }

                $legacyVersion = $null
                try {
                    $legacyVersion = (Test-ModuleManifest -Path $legacyManifestPath -ErrorAction Stop).Version.ToString()
                }
                catch {
                    Write-Output "WARNING: Could not read legacy manifest at $legacyManifestPath; removing unversioned module files to avoid stale module resolution"
                }

                if ($legacyVersion -eq $TargetVersion) {
                    continue
                }

                Get-ChildItem -LiteralPath $moduleRoot -File -Force -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction Stop
                Write-Output "INFO: Removed unversioned $ModuleName files from $moduleRoot"
            }
        }

        $modulePaths = @()
        if (-not [string]::IsNullOrWhiteSpace($env:PSModulePath)) {
            $modulePaths = $env:PSModulePath -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
        }

        $machineCandidates = @()
        if (-not [string]::IsNullOrWhiteSpace($env:ProgramW6432)) {
            $machineCandidates += (Join-Path $env:ProgramW6432 'WindowsPowerShell\Modules')
        }
        if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
            $machineCandidates += (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules')
        }

        # Also consider any machine-wide paths already present in PSModulePath.
        $machineCandidates += $modulePaths | Where-Object { $_ -match '\\Program Files( \(x86\))?\\WindowsPowerShell\\Modules$' }

        $installRoot = Get-WritableModuleRoot -Candidates ($machineCandidates | Select-Object -Unique)
        if ([string]::IsNullOrWhiteSpace($installRoot)) {
            throw "No writable machine-wide module path found (need write access to Program Files\\WindowsPowerShell\\Modules). Run elevated or as SYSTEM."
        }

        $destVersionRoot = Join-Path (Join-Path $installRoot $ModuleName) $ModuleVersion
        $destManifestPath = Join-Path $destVersionRoot "${ModuleName}.psd1"

        $alreadyInstalled = $false

        if (Test-Path -LiteralPath $destManifestPath) {
            try {
                Import-Module -Name $destManifestPath -Force -ErrorAction Stop
                $alreadyInstalled = $true
            }
            catch {
                # fall through to reinstall
            }
        }

        Remove-OldModuleVersions -ModulePaths $modulePaths -ModuleName $ModuleName -TargetVersion $ModuleVersion

        if ($alreadyInstalled) {
            Write-Output "INFO: $ModuleName $ModuleVersion already installed at $destVersionRoot, skipping reinstall."
            return
        }

        $tempRoot = Join-Path $env:TEMP ("${ModuleName}_" + [Guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $tempRoot -Force
        $zipPath = Join-Path $tempRoot "${ModuleName}-${ModuleVersion}.zip"
        $extractPath = Join-Path $tempRoot 'extracted'

        try {
            Invoke-WebRequest -Uri $ZipUrl -OutFile $zipPath -UseBasicParsing
        }
        catch {
            $wc = New-Object System.Net.WebClient
            $wc.DownloadFile($ZipUrl, $zipPath)
        }

        if (-not (Test-Path -LiteralPath $zipPath)) {
            throw "Download failed: $zipPath not found"
        }

        try { Unblock-File -LiteralPath $zipPath -ErrorAction SilentlyContinue } catch { }

        Expand-Archive -LiteralPath $zipPath -DestinationPath $extractPath -Force

        $manifest = Get-ChildItem -Path $extractPath -Recurse -File -Filter "${ModuleName}.psd1" | Select-Object -First 1
        if ($null -eq $manifest) {
            throw "Module manifest ${ModuleName}.psd1 not found in extracted archive"
        }
        $moduleSourceRoot = Split-Path -Parent $manifest.FullName

        Remove-Item -LiteralPath $destVersionRoot -Recurse -Force -ErrorAction SilentlyContinue
        $null = New-Item -ItemType Directory -Path $destVersionRoot -Force

        Copy-Item -Path (Join-Path $moduleSourceRoot '*') -Destination $destVersionRoot -Recurse -Force

        if (-not (Test-Path -LiteralPath $destManifestPath)) {
            throw "$ModuleName manifest not found at $destManifestPath after copy"
        }

        try {
            Get-ChildItem -LiteralPath $destVersionRoot -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
        }
        catch { }

        Import-Module -Name $destManifestPath -Force -ErrorAction Stop
        Write-Output "INFO: Installed $ModuleName $ModuleVersion from ZIP to $destVersionRoot"
    }

    Start-Sleep -Seconds 1

    # Verify AudioDeviceCmdlets is installed/available
    $installed = Invoke-LocalOrRemote -ComputerName $PC -IsLocal $isLocal -ArgumentList @(
        $AudioDeviceCmdletsModuleName,
        $AudioDeviceCmdletsVersion
    ) -ScriptBlock {
        param(
            [string]$ModuleName,
            [string]$ModuleVersion
        )

        if ([string]::IsNullOrWhiteSpace($env:PSModulePath)) {
            return $null
        }

        foreach ($root in ($env:PSModulePath -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
            $candidate = Join-Path (Join-Path (Join-Path $root $ModuleName) $ModuleVersion) "${ModuleName}.psd1"
            if (Test-Path -LiteralPath $candidate) {
                try {
                    Import-Module -Name $candidate -Force -ErrorAction Stop
                }
                catch { }

                return (Get-Module -ListAvailable -Name $ModuleName |
                    Where-Object { $_.Version -eq ([version]$ModuleVersion) } |
                    Sort-Object Version -Descending |
                    Select-Object -First 1)
            }
        }

        return $null
    }

    if ($null -eq $installed) {
        throw "$PC AudioDeviceCmdlets not installed"
    }
    else {
        Write-Output "INFO: AudioDeviceCmdlets installed ($($installed.Version))"
    }


    # This script owns its own Startup bat + recall script.
    # Consolidation/ordering is handled by Cleanup.ps1.
    if ($isLocal) { $prefix = "C:\" }
    else { $prefix = "\\$PC\C$\" }

    $ctsFolder = Join-Path $prefix 'ProgramData\CTS'
    if (-not (Test-Path -LiteralPath $ctsFolder)) {
        New-Item -ItemType Directory -Path $ctsFolder -Force | Out-Null
    }

    $startupFolder = Join-Path $prefix 'ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp'
    if (-not (Test-Path -LiteralPath $startupFolder)) {
        New-Item -ItemType Directory -Path $startupFolder -Force | Out-Null
    }

    $sourceScripts = Join-Path $PSScriptRoot 'local_scripts'

    # Recall scripts (startup)
    $audioRecallScriptSourcePath = Join-Path $sourceScripts 'reacall_audio_config.ps1'
    $audioRecallScriptDestPath = Join-Path $ctsFolder 'recall_audio_config.ps1'

    $audioStartupBatSourcePath = Join-Path $sourceScripts 'run_recall_audio_config.bat'
    $audioStartupBatDestPath = Join-Path $startupFolder 'cts_audio_startup.bat'

    Copy-Item -Path $audioRecallScriptSourcePath -Destination $audioRecallScriptDestPath -Force -ErrorAction Stop
    Copy-Item -Path $audioStartupBatSourcePath -Destination $audioStartupBatDestPath -Force -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $audioRecallScriptDestPath)) {
        throw "$PC recall_audio_config.ps1 not found at $audioRecallScriptDestPath"
    }
    if (-not (Test-Path -LiteralPath $audioStartupBatDestPath)) {
        throw "$PC cts_audio_startup.bat not found at $audioStartupBatDestPath"
    }

    # Save scripts
    $audioSaveScriptSourcePath = Join-Path $sourceScripts 'save_audio_config.ps1'
    $audioSaveScriptDestPath = Join-Path $ctsFolder 'save_audio_config.ps1'

    $audioSaveBatSourcePath = Join-Path $sourceScripts 'run_save_audio_config.bat'
    $audioSaveBatDestPath = Join-Path $ctsFolder 'run_save_audio_config.bat'

    Copy-Item -Path $audioSaveScriptSourcePath -Destination $audioSaveScriptDestPath -Force -ErrorAction Stop
    Copy-Item -Path $audioSaveBatSourcePath -Destination $audioSaveBatDestPath -Force -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $audioSaveScriptDestPath)) {
        throw "$PC save_audio_config.ps1 not found at $audioSaveScriptDestPath"
    }

    Write-Output "INFO: Installed audio recall and save scripts"

    # Standalone: drop SAVE_AUDIO_SETTINGS.bat on Public Desktop
    if ($standalone) {
        $saveAudioBatSourcePath = Join-Path $sourceScripts 'SAVE_AUDIO_SETTINGS.bat'
        $publicDesktopPath = Join-Path $prefix 'Users\Public\Desktop'
        $saveAudioDesktopPath = Join-Path $publicDesktopPath 'SAVE_AUDIO_SETTINGS.bat'

        Copy-Item -Path $saveAudioBatSourcePath -Destination $saveAudioDesktopPath -Force -ErrorAction Stop

        # Add self-destruct to desktop file
        Add-Content -LiteralPath $saveAudioDesktopPath -Encoding ASCII -Value @(
            ''
            '(goto) 2>nul & del "%~f0"'
        )

        Write-Output "INFO: Dropped SAVE_AUDIO_SETTINGS.bat on Public Desktop"
    }


} # end try
catch {
    Write-Output "ERROR: InstallAudioDeviceCmdlets failed: $_"
    Exit 1
}

Exit 0
