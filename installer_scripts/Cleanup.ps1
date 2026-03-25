Param(
    [Parameter(Mandatory = $true)]
    [string]$PC,
    [Parameter(Mandatory = $false)]
    [string]$standalone = "true"
)


# Optomizes a non-standalone deployment.
# Run this code as the last step in a non-standalone deployment

# Workaround for passing in booleans from Python
if ($standalone -eq "true") {
    $standalone = $true
}
else {
    $standalone = $false
}

Import-Module (Join-Path $PSScriptRoot 'shared\SharedHelpers.psm1') -Force

$isLocal = Test-IsLocalComputer -ComputerName $PC

try {
    if ($standalone -and -not $isLocal) {
        if (-not (Test-HostReachable -ComputerName $PC -TimeoutMilliseconds 1000)) {
            throw "$PC is not reachable"
        }
    }

    if ($isLocal) { $prefix = "C:\" }
    else { $prefix = "\\$PC\C$\" }

    $startupFolderSuffix = "ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"

    $startupFolder = Join-Path $prefix $startupFolderSuffix

    if (!(Test-Path $startupFolder)) {
        Write-Output "ERROR: Startup path not found"
        throw
    }

    # Remove legacy startup bats (best-effort)
    $legacyStartupBats = @('dc.bat', 'run_bgi.bat')
    foreach ($legacyBat in $legacyStartupBats) {
        $legacyBatPath = Join-Path $startupFolder $legacyBat
        if (Test-Path -LiteralPath $legacyBatPath) {
            try {
                Remove-Item -LiteralPath $legacyBatPath -Force -ErrorAction Stop
                Write-Output "INFO: Removed legacy Startup bat: $legacyBatPath"
            }
            catch {
                Write-Output "WARNING: Could not remove legacy Startup bat: $legacyBatPath ($($_.Exception.Message))"
            }
        }
    }


    # Consolidate per-script Startup bats into one ordered bat.
    # Ensures display settings recall runs before BGInfo.
    $displayStartupBatPath = Join-Path $startupFolder 'cts_display_startup.bat'
    $audioStartupBatPath = Join-Path $startupFolder 'cts_audio_startup.bat'
    $bginfoStartupBatPath = Join-Path $startupFolder 'cts_bginfo_startup.bat'
    $consolidatedStartupBatPath = Join-Path $startupFolder 'av_config_startup.bat'

    $displayLines = @()
    $audioLines = @()
    $bginfoLines = @()

    if (Test-Path -LiteralPath $displayStartupBatPath) {
        $displayLines = Get-Content -LiteralPath $displayStartupBatPath |
        Where-Object { $_ -notmatch '^\s*@echo\s+off\s*$' -and $_ -notmatch '^\s*$' }
    }

    if (Test-Path -LiteralPath $bginfoStartupBatPath) {
        $bginfoLines = Get-Content -LiteralPath $bginfoStartupBatPath |
        Where-Object { $_ -notmatch '^\s*@echo\s+off\s*$' -and $_ -notmatch '^\s*$' }
    }

    if (Test-Path -LiteralPath $audioStartupBatPath) {
        $audioLines = Get-Content -LiteralPath $audioStartupBatPath |
        Where-Object { $_ -notmatch '^\s*@echo\s+off\s*$' -and $_ -notmatch '^\s*$' }
    }

    # Back-compat: if dedicated files don't exist but a consolidated file does, extract the lines.
    if ((($displayLines.Count -eq 0) -or ($audioLines.Count -eq 0) -or ($bginfoLines.Count -eq 0)) -and (Test-Path -LiteralPath $consolidatedStartupBatPath)) {
        $existingLines = Get-Content -LiteralPath $consolidatedStartupBatPath |
        Where-Object { $_ -notmatch '^\s*@echo\s+off\s*$' -and $_ -notmatch '^\s*$' }

        if ($displayLines.Count -eq 0) {
            $displayLines = $existingLines | Where-Object { $_ -match '(?i)recall_display_config\.ps1' }
        }
        if ($bginfoLines.Count -eq 0) {
            $bginfoLines = $existingLines | Where-Object { $_ -match '(?i)bginfo|\.bgi\b' }
        }
        if ($audioLines.Count -eq 0) {
            $audioLines = $existingLines | Where-Object { $_ -match '(?i)recall_audio_config\.ps1' }
        }
    }

    if (($displayLines.Count -gt 0) -or ($audioLines.Count -gt 0) -or ($bginfoLines.Count -gt 0)) {
        $newContent = @('@echo off')
        if ($displayLines.Count -gt 0) {
            $newContent += $displayLines
        }
        if (($displayLines.Count -gt 0) -and ($audioLines.Count -gt 0)) {
            $newContent += ''
        }
        if ($audioLines.Count -gt 0) {
            $newContent += $audioLines
        }
        if ((($displayLines.Count -gt 0) -or ($audioLines.Count -gt 0)) -and ($bginfoLines.Count -gt 0)) {
            $newContent += ''
        }
        if ($bginfoLines.Count -gt 0) {
            $newContent += $bginfoLines
        }

        $newContent | Set-Content -LiteralPath $consolidatedStartupBatPath -Force -Encoding ASCII
        Write-Output "INFO: Wrote consolidated Startup bat to $consolidatedStartupBatPath"

        # Remove per-script bats so they don't also run at login.
        Remove-Item -LiteralPath $displayStartupBatPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $audioStartupBatPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $bginfoStartupBatPath -Force -ErrorAction SilentlyContinue
    }
    else {
        Write-Output "INFO: No CTS startup bats found to consolidate"
    }

    $saveFile = "SAVE_AV_SETTINGS.bat"

    $sourceScripts = Join-Path $PSScriptRoot 'local_scripts'
    $saveScriptSource = Join-Path $sourceScripts $saveFile
    
    $publicDesktopPath = Join-Path $prefix "Users\Public\Desktop"
    $saveAvSettingsDesktopPath = Join-Path $publicDesktopPath $saveFile
    $saveAvSettingsBackupPath = Join-Path $prefix "ProgramData\CTS\SAVE_AV_SETTINGS.bat"

    Copy-Item $saveScriptSource $saveAvSettingsBackupPath
    Copy-Item $saveScriptSource $saveAvSettingsDesktopPath

    # Add self-destruct to desktop file
    Add-Content -LiteralPath $saveAvSettingsDesktopPath -Encoding ASCII -Value @(
        ''
        '(goto) 2>nul & del "%~f0"'
    )

    # Remove standalone SAVE_AUDIO_SETTINGS.bat left by AudioDeviceCmdlets installer
    $standaloneAudioSavePath = Join-Path $publicDesktopPath 'SAVE_AUDIO_SETTINGS.bat'
    if (Test-Path -LiteralPath $standaloneAudioSavePath) {
        Remove-Item -LiteralPath $standaloneAudioSavePath -Force -ErrorAction SilentlyContinue
        Write-Output "INFO: Removed standalone SAVE_AUDIO_SETTINGS.bat (consolidated into SAVE_AV_SETTINGS.bat)"
    }
}

catch {
    Write-Output "ERROR: Cleanup failed: $_"
    Exit 1
}

Start-Sleep 1
Exit 0