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

$Version = '5.2.2'

Import-Module (Join-Path $PSScriptRoot 'shared\SharedHelpers.psm1') -Force

$isLocal = Test-IsLocalComputer -ComputerName $PC

Write-Output "$PC Installing DisplayConfig"


try {

    # Ping first if standalone
    if ($standalone -and -not $isLocal) {
        if (-not (Test-HostReachable -ComputerName $PC -TimeoutMilliseconds 1000)) {
            throw "$PC is not reachable"
        }
    }

    # Create C:\ProgramData\CTS folder if it doesn't exist, if standalone
    if ($standalone) {
        Invoke-LocalOrRemote -ComputerName $PC -IsLocal $isLocal -ScriptBlock {
            if (!(Test-Path -Path 'C:\ProgramData\CTS')) {
                New-Item -ItemType Directory -Path 'C:\ProgramData\CTS' -Force | Out-Null
            }
        } | Out-Null
    }

    $Version = ConvertTo-NormalizedVersion -Version $Version
    $DisplayConfigModuleName = 'DisplayConfig'
    $DisplayConfigZipUrl = "https://github.com/mefranklin6/$DisplayConfigModuleName/releases/download/v$Version/$DisplayConfigModuleName-$Version.zip"


    # Install DisplayConfig from a pinned GitHub release ZIP
    Invoke-LocalOrRemote -ComputerName $PC -IsLocal $isLocal -ArgumentList @(
        $DisplayConfigZipUrl,
        $DisplayConfigModuleName,
        $Version
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

        if (Test-Path -LiteralPath $destManifestPath) {
            try {
                Import-Module -Name $destManifestPath -Force -ErrorAction Stop
                Write-Output "$env:COMPUTERNAME $ModuleName $ModuleVersion already installed at $destVersionRoot, skipping."
                return
            }
            catch {
                # fall through to reinstall
            }
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
        Write-Output "$env:COMPUTERNAME Installed $ModuleName $ModuleVersion from ZIP to $destVersionRoot"
    }

    Start-Sleep -Seconds 1

    # Verify DisplayConfig is installed/available
    $installed = Invoke-LocalOrRemote -ComputerName $PC -IsLocal $isLocal -ArgumentList @(
        $DisplayConfigModuleName,
        $Version
    ) -ScriptBlock {
        param(
            [string]$ModuleName,
            [string]$ModuleVersion
        )

        $m = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
        if ($m) { return $m }

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
                return (Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1)
            }
        }

        return $null
    }

    if ($null -eq $installed) {
        throw "$PC DisplayConfig (mefranklin6 fork) not installed"
    }
    else {
        Write-Output "$PC DisplayConfig (mefranklin6 fork) installed ($($installed.Version))"
    }

    #### Copy execution scripts to destination ####

    $sourceScripts = Join-Path $PSScriptRoot 'local_scripts'

    $localRoot = 'C:\ProgramData\CTS'
    $remoteRoot = "\\$PC\C$\ProgramData\CTS"
    $root = if ($isLocal) { $localRoot } else { $remoteRoot }

    $saveDisplayScriptFileName = 'display_config_save_profile.ps1'
    $saveDisplayExecutionScriptFileName = 'run_display_config_save_profile.bat'
    $recallDisplayScriptFileName = 'recall_display_config.ps1'
    $recallDisplayExecutionSourceFileName = 'run_recall_display_config.bat'
    $recallDisplayExecutionDestFileName = 'cts_display_startup.bat'

    # Destination paths (local or UNC)
    $saveDisplayScriptPath = Join-Path $root $saveDisplayScriptFileName
    $saveDisplayExecutionScriptPath = Join-Path $root $saveDisplayExecutionScriptFileName
    $recallDisplayScriptPath = Join-Path $root $recallDisplayScriptFileName

    if ($isLocal) {
        $recallDisplayExecutionScriptPath = Join-Path 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup' $recallDisplayExecutionDestFileName
    }
    else {
        $recallDisplayExecutionScriptPath = Join-Path "\\$PC\C$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup" $recallDisplayExecutionDestFileName
    }

    # Local paths - used in save/recall logic below
    $localSaveDisplayScriptPath = Join-Path $localRoot $saveDisplayScriptFileName

    # Map source file paths to their destinations
    $fileMappings = [ordered]@{
        (Join-Path $sourceScripts $saveDisplayScriptFileName)            = $saveDisplayScriptPath
        (Join-Path $sourceScripts $saveDisplayExecutionScriptFileName)   = $saveDisplayExecutionScriptPath
        (Join-Path $sourceScripts $recallDisplayScriptFileName)          = $recallDisplayScriptPath
        (Join-Path $sourceScripts $recallDisplayExecutionSourceFileName) = $recallDisplayExecutionScriptPath
    }

    try {
        $startupDir = Split-Path -Parent $recallDisplayExecutionScriptPath
        if (!(Test-Path -Path $startupDir)) {
            New-Item -ItemType Directory -Path $startupDir -Force | Out-Null
        }
    }
    catch {
        throw "Could not ensure Startup folder exists on $PC : $_"
    }

    foreach ($source in $fileMappings.Keys) {
        $dest = $fileMappings[$source]
        try {
            Copy-Item -Path $source -Destination $dest -Force -ErrorAction Stop
        }
        catch {
            throw "Failed to copy $source to $dest : $_"
        }
    }

    foreach ($dest in $fileMappings.Values) {
        if (!(Test-Path $dest)) {
            throw "Could not verify $dest is on $PC"
        }
    }

    # Untrust PSGallery
    try {
        Invoke-LocalOrRemote -ComputerName $PC -IsLocal $isLocal -ScriptBlock {
            Set-PSRepository -Name PSGallery -InstallationPolicy Untrusted -ErrorAction Stop
        }
    }
    catch {
        Write-Output "$PC Failed to un-register PSGallery repository: $_"
    }


    Write-Output "$PC Please run $localSaveDisplayScriptPath on machine to save display configurations"
    Write-Output "$PC Installed DisplayConfig and Scripts"

} # end try
catch {
    Write-Output "$PC InstallDisplayConfig failed: $_"
    Exit 1
}

Exit 0
