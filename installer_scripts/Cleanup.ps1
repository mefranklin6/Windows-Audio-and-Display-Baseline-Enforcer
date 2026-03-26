Param(
    [Parameter(Mandatory = $true)]
    [string]$PC,
    [Parameter(Mandatory = $false)]
    [string]$standalone = "true"
)


# Optomizes a non-standalone deployment.
# Run this code as the last step in a non-standalone deployment

# Feature Flag: Add reboot and Log Out shortcuts to the public desktop
# These shortcuts recall settings first.
$addShortcuts = $true

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
    $consolidatedStartupBatPath = Join-Path $startupFolder 'av_config_recall.bat'

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

    $ctsFolder = Join-Path $prefix 'ProgramData\CTS'
    if (-not (Test-Path -LiteralPath $ctsFolder)) {
        New-Item -ItemType Directory -Path $ctsFolder -Force | Out-Null
    }

    $saveFile = "SAVE_AV_SETTINGS.bat"

    $sourceScripts = Join-Path $PSScriptRoot 'local_scripts'
    $saveScriptSource = Join-Path $sourceScripts $saveFile
    
    $publicDesktopPath = Join-Path $prefix "Users\Public\Desktop"
    $saveAvSettingsDesktopPath = Join-Path $publicDesktopPath $saveFile
    $saveAvSettingsBackupPath = Join-Path $ctsFolder $saveFile

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

    if ($addShortcuts) {
        $localCtsFolder = 'C:\ProgramData\CTS'
        $localPublicDesktopPath = 'C:\Users\Public\Desktop'
        $programDataStartupBatPath = Join-Path $ctsFolder 'av_config_recall.bat'
        $localProgramDataStartupBatPath = Join-Path $localCtsFolder 'av_config_recall.bat'
        $logoutActionBatPath = Join-Path $ctsFolder 'cts_log_out_and_recall_av.bat'
        $localLogoutActionBatPath = Join-Path $localCtsFolder 'cts_log_out_and_recall_av.bat'
        $rebootActionBatPath = Join-Path $ctsFolder 'cts_reboot_and_recall_av.bat'
        $localRebootActionBatPath = Join-Path $localCtsFolder 'cts_reboot_and_recall_av.bat'
        $localLogoutIconLocation = 'C:\Windows\System32\shell32.dll,44'
        $localRebootIconLocation = 'C:\Windows\System32\shell32.dll,238'

        $startupBatCopyContent = if (Test-Path -LiteralPath $consolidatedStartupBatPath) {
            Get-Content -LiteralPath $consolidatedStartupBatPath
        }
        else {
            @('@echo off')
        }

        if ($startupBatCopyContent.Count -eq 0) {
            $startupBatCopyContent = @('@echo off')
        }

        $shutdownRecallBatContent = @(
            $startupBatCopyContent | Where-Object { $_ -notmatch '(?i)bginfo|\.bgi\b' }
        )

        if ($shutdownRecallBatContent.Count -eq 0) {
            $shutdownRecallBatContent = @('@echo off')
        }

        $shutdownRecallBatContent | Set-Content -LiteralPath $programDataStartupBatPath -Force -Encoding ASCII
        @(
            '@echo off'
            ('call "{0}"' -f $localProgramDataStartupBatPath)
            'shutdown.exe /l'
        ) | Set-Content -LiteralPath $logoutActionBatPath -Force -Encoding ASCII

        @(
            '@echo off'
            ('call "{0}"' -f $localProgramDataStartupBatPath)
            'shutdown.exe /r /t 1'
        ) | Set-Content -LiteralPath $rebootActionBatPath -Force -Encoding ASCII

        Invoke-LocalOrRemote -ComputerName $PC -IsLocal $isLocal -ArgumentList @(
            $localPublicDesktopPath,
            $localLogoutActionBatPath,
            $localRebootActionBatPath,
            $localLogoutIconLocation,
            $localRebootIconLocation,
            $localCtsFolder
        ) -ScriptBlock {
            param(
                [Parameter(Mandatory = $true)]
                [string]$PublicDesktopPath,
                [Parameter(Mandatory = $true)]
                [string]$LogoutActionBatPath,
                [Parameter(Mandatory = $true)]
                [string]$RebootActionBatPath,
                [Parameter(Mandatory = $true)]
                [string]$LogoutIconLocation,
                [Parameter(Mandatory = $true)]
                [string]$RebootIconLocation,
                [Parameter(Mandatory = $true)]
                [string]$CtsFolder
            )

            $ErrorActionPreference = 'Stop'

            function Set-DesktopShortcut {
                param(
                    [Parameter(Mandatory = $true)]
                    [string]$ShortcutPath,
                    [Parameter(Mandatory = $true)]
                    [string]$TargetPath,
                    [Parameter(Mandatory = $true)]
                    [string]$Description,
                    [Parameter(Mandatory = $true)]
                    [string]$IconLocation,
                    [Parameter(Mandatory = $true)]
                    [string]$WorkingDirectory
                )

                $shell = New-Object -ComObject WScript.Shell
                try {
                    $shortcut = $shell.CreateShortcut($ShortcutPath)
                    $shortcut.TargetPath = $TargetPath
                    $shortcut.WorkingDirectory = $WorkingDirectory
                    $shortcut.Description = $Description
                    $shortcut.IconLocation = $IconLocation
                    $shortcut.WindowStyle = 7
                    $shortcut.Save()
                }
                finally {
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
                }
            }

            $logoutShortcutPath = Join-Path $PublicDesktopPath 'Log Out.lnk'
            Set-DesktopShortcut -ShortcutPath $logoutShortcutPath -TargetPath $LogoutActionBatPath -Description 'Recall saved AV settings, then log out.' -IconLocation $LogoutIconLocation -WorkingDirectory $CtsFolder
            Write-Output "INFO: Configured Log Out shortcut at $logoutShortcutPath"

            $rebootShortcutPath = Join-Path $PublicDesktopPath 'Reboot.lnk'
            Set-DesktopShortcut -ShortcutPath $rebootShortcutPath -TargetPath $RebootActionBatPath -Description 'Recall saved AV settings, then reboot.' -IconLocation $RebootIconLocation -WorkingDirectory $CtsFolder
            Write-Output "INFO: Configured Reboot shortcut at $rebootShortcutPath"
        }

        Write-Output "INFO: Created ProgramData recall, log out, and reboot launchers in $ctsFolder"
    }
    else {
        Write-Output 'INFO: Shortcut creation disabled by $addShortcuts'
    }

}

catch {
    Write-Output "ERROR: Cleanup failed: $_"
    Exit 1
}

Start-Sleep 1
Exit 0