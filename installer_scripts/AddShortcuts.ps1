Param(
    [Parameter(Mandatory = $true)]
    [string]$PC,
    [Parameter(Mandatory = $false)]
    [string]$standalone = "true"
)


# Run after Cleanup.ps1 so Startup recall ordering is already consolidated.

# Delay before calling shutdown.exe to allow recall to happen first
$shortcutRecallGraceSeconds = 5 # Default 5

# Workaround for passing in booleans from Python
if ($standalone -eq "true") {
    $standalone = $true
}
else {
    $standalone = $false
}

Import-Module (Join-Path $PSScriptRoot 'shared\SharedHelpers.psm1') -Force

$isLocal = Test-IsLocalComputer -ComputerName $PC

function Set-ShortcutActionLauncher {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LauncherPath,
        [Parameter(Mandatory = $true)]
        [string]$RecallBatPath,
        [Parameter(Mandatory = $true)]
        [string]$ShutdownCommand,
        [Parameter(Mandatory = $true)]
        [int]$GraceSeconds,
        [Parameter(Mandatory = $false)]
        [string]$NotificationBody = ''
    )

    $launcherContent = @(
        '@echo off'
        'setlocal'
    )

    if ($NotificationBody -ne '') {
        $launcherContent += ('start "" /b powershell.exe -WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -File "C:\ProgramData\CTS\cts_notify.ps1" -Body "{0}"' -f $NotificationBody)
    }

    $launcherContent += @(
        ('set "RECALL_BAT={0}"' -f $RecallBatPath)
        'if exist "%RECALL_BAT%" ('
        '    start "" /min cmd.exe /c call "%RECALL_BAT%"'
    )

    if ($GraceSeconds -gt 0) {
        $launcherContent += ('    timeout /t {0} /nobreak >nul' -f $GraceSeconds)
    }

    $launcherContent += @(
        ')'
        $ShutdownCommand
        'endlocal'
    )

    if ($PSCmdlet.ShouldProcess($LauncherPath, 'Write shortcut action launcher')) {
        $launcherContent | Set-Content -LiteralPath $LauncherPath -Force -Encoding ASCII
    }
}

try {
    if ($standalone -and -not $isLocal) {
        if (-not (Test-HostReachable -ComputerName $PC -TimeoutMilliseconds 1000)) {
            throw "$PC is not reachable"
        }
    }

    if ($isLocal) { $prefix = "C:\" }
    else { $prefix = "\\$PC\C$\" }

    $startupFolderSuffix = 'ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp'
    $startupFolder = Join-Path $prefix $startupFolderSuffix

    if (!(Test-Path -LiteralPath $startupFolder)) {
        Write-Output 'ERROR: Startup path not found'
        throw
    }

    $ctsFolder = Join-Path $prefix 'ProgramData\CTS'
    if (-not (Test-Path -LiteralPath $ctsFolder)) {
        New-Item -ItemType Directory -Path $ctsFolder -Force | Out-Null
    }

    $sourceScripts = Join-Path $PSScriptRoot 'local_scripts'
    $notifyScriptSource = Join-Path $sourceScripts 'cts_notify.ps1'

    $consolidatedStartupBatPath = Join-Path $startupFolder 'av_config_recall.bat'
    $publicDesktopPath = Join-Path $prefix 'Users\Public\Desktop'

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

    $notifyScriptPath = Join-Path $ctsFolder 'cts_notify.ps1'
    Copy-Item -LiteralPath $notifyScriptSource -Destination $notifyScriptPath -Force
    Write-Output "INFO: Copied notification helper to $notifyScriptPath"

    Set-ShortcutActionLauncher -LauncherPath $logoutActionBatPath -RecallBatPath $localProgramDataStartupBatPath -ShutdownCommand 'shutdown.exe /l' -GraceSeconds $shortcutRecallGraceSeconds -NotificationBody 'You will be signed out shortly.'
    Set-ShortcutActionLauncher -LauncherPath $rebootActionBatPath -RecallBatPath $localProgramDataStartupBatPath -ShutdownCommand 'shutdown.exe /r /t 1' -GraceSeconds $shortcutRecallGraceSeconds -NotificationBody 'You will be restarted shortly.'

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
            [CmdletBinding(SupportsShouldProcess = $true)]
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
                if ($PSCmdlet.ShouldProcess($ShortcutPath, 'Create or update desktop shortcut')) {
                    $shortcut = $shell.CreateShortcut($ShortcutPath)
                    $shortcut.TargetPath = $TargetPath
                    $shortcut.WorkingDirectory = $WorkingDirectory
                    $shortcut.Description = $Description
                    $shortcut.IconLocation = $IconLocation
                    $shortcut.WindowStyle = 7
                    $shortcut.Save()
                }
            }
            finally {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
            }
        }

        function Remove-ExistingDesktopShortcut {
            [CmdletBinding(SupportsShouldProcess = $true)]
            param(
                [Parameter(Mandatory = $true)]
                [string]$DesktopPath
            )

            $shortcutBaseNames = @('Log Out', 'Logout', 'Reboot', 'Restart')
            $existingShortcuts = Get-ChildItem -LiteralPath $DesktopPath -File -Filter '*.lnk' -ErrorAction SilentlyContinue |
            Where-Object { $shortcutBaseNames -contains $_.BaseName }

            foreach ($shortcut in $existingShortcuts) {
                if ($PSCmdlet.ShouldProcess($shortcut.FullName, 'Remove existing desktop shortcut')) {
                    Remove-Item -LiteralPath $shortcut.FullName -Force -ErrorAction SilentlyContinue
                    Write-Output "INFO: Removed existing desktop shortcut at $($shortcut.FullName)"
                }
            }
        }

        Remove-ExistingDesktopShortcut -DesktopPath $PublicDesktopPath

        $logoutShortcutPath = Join-Path $PublicDesktopPath 'Log Out.lnk'
        Set-DesktopShortcut -ShortcutPath $logoutShortcutPath -TargetPath $LogoutActionBatPath -Description 'Recall saved AV settings, then log out.' -IconLocation $LogoutIconLocation -WorkingDirectory $CtsFolder
        Write-Output "INFO: Configured Log Out shortcut at $logoutShortcutPath"

        $rebootShortcutPath = Join-Path $PublicDesktopPath 'Reboot.lnk'
        Set-DesktopShortcut -ShortcutPath $rebootShortcutPath -TargetPath $RebootActionBatPath -Description 'Recall saved AV settings, then reboot.' -IconLocation $RebootIconLocation -WorkingDirectory $CtsFolder
        Write-Output "INFO: Configured Reboot shortcut at $rebootShortcutPath"
    }

    Write-Output "INFO: Created ProgramData recall, log out, and reboot launchers in $ctsFolder"
}
catch {
    Write-Output "ERROR: Shortcut setup failed: $_"
    Exit 1
}

Start-Sleep 1
Exit 0