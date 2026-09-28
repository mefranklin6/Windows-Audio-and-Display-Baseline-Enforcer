[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ComputerName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$BackupFolder
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$sharedHelpersPath = Join-Path $PSScriptRoot '..\installer_scripts\shared\SharedHelpers.psm1'

function ConvertTo-StableAudioDeviceName {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Name
    )

    # Windows adds an instance number when an endpoint is re-enumerated, for
    # example "ExtronScalerD (2- HD Audio Driver for Display Audio)". The
    # number is not part of the hardware identity and can change after driver
    # updates or reconnects. Restrict normalization to the start of a
    # parenthesized hardware name so legitimate numbers elsewhere are kept.
    return (($Name.Trim()) -replace '\(\s*\d+\s*-\s+', '(')
}

function Get-SavedAudioDeviceIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath
    )

    if (-not (Test-Path -LiteralPath $LiteralPath -PathType Leaf)) {
        return @()
    }

    $json = Get-Content -LiteralPath $LiteralPath -Raw -ErrorAction Stop
    $parsed = ConvertFrom-Json -InputObject $json -ErrorAction Stop
    $devices = if ($parsed -is [System.Array]) {
        $parsed
    }
    else {
        @($parsed)
    }

    $identities = @(
        foreach ($device in $devices) {
            if ([string]::IsNullOrWhiteSpace([string]$device.Name)) {
                continue
            }

            $name = ([string]$device.Name).Trim()
            $stableName = ConvertTo-StableAudioDeviceName -Name $name
            $type = ([string]$device.Type).Trim()
            [pscustomobject]@{
                RawKey    = ('{0}|{1}' -f $type, $name).ToLowerInvariant()
                StableKey = ('{0}|{1}' -f $type, $stableName).ToLowerInvariant()
                Label     = if ([string]::IsNullOrWhiteSpace($type)) {
                    $name
                }
                else {
                    '{0} ({1})' -f $name, $type
                }
            }
        }
    )
    return @($identities | Sort-Object -Property RawKey -Unique)
}

try {
    Import-Module $sharedHelpersPath -Force -ErrorAction Stop

    $backupDirectory = Get-Item -LiteralPath $BackupFolder -ErrorAction Stop
    if (-not $backupDirectory.PSIsContainer) {
        throw "BackupFolder is not a directory: $BackupFolder"
    }

    $savedAudio = @(Get-SavedAudioDeviceIdentity -LiteralPath (
            Join-Path $backupDirectory.FullName 'audio_device_list.json'
        ))

    $scanAudioDevices = {
        param([bool]$IncludeAudio)

        $audioRaw = @()
        $audioStable = @()
        $scanErrors = @()

        if ($IncludeAudio) {
            try {
                Import-Module AudioDeviceCmdlets -ErrorAction Stop
                $devices = @(Get-AudioDevice -List -ErrorAction Stop |
                        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Name) })
                $audioRaw = @($devices | ForEach-Object {
                        $name = ([string]$_.Name).Trim()
                        $type = ([string]$_.Type).Trim()
                        ('{0}|{1}' -f $type, $name).ToLowerInvariant()
                    })
                $audioStable = @($devices | ForEach-Object {
                        # Keep this normalization inside the script block so it
                        # also works when the scan runs in a remote session.
                        $name = (([string]$_.Name).Trim()) -replace '\(\s*\d+\s*-\s+', '('
                        $type = ([string]$_.Type).Trim()
                        ('{0}|{1}' -f $type, $name).ToLowerInvariant()
                    })
            }
            catch {
                $scanErrors += 'Audio device scan failed: {0}' -f $_.Exception.Message
            }
        }

        [pscustomobject]@{
            AudioRaw    = $audioRaw
            AudioStable = $audioStable
            ScanErrors  = $scanErrors
        }
    }

    $name = $ComputerName.Trim()
    $isLocal = Test-IsLocalComputer -ComputerName $name
    $current = Invoke-LocalOrRemote -ComputerName $name -IsLocal $isLocal -ScriptBlock $scanAudioDevices -ArgumentList @(
        $savedAudio.Count -gt 0
    )

    $currentAudioRaw = @($current.AudioRaw)
    $currentAudioStable = @($current.AudioStable)
    $scanErrors = @($current.ScanErrors | Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_)
        })
    $scanReturnedNoDevices = ($savedAudio.Count -gt 0) -and
        ($currentAudioRaw.Count -eq 0) -and
        ($currentAudioStable.Count -eq 0) -and
        ($scanErrors.Count -eq 0)

    # A sleeping display or USB device can make AudioDeviceCmdlets return an
    # empty list while the PC is in a low-power state. In that specific case,
    # allow the restore without presenting a hardware-change warning. A scan
    # error or a partial device list is still evaluated normally.
    $missing = @()
    if (-not $scanReturnedNoDevices) {
        $missing = @(
            $savedAudio |
            Where-Object {
                $savedDevice = $_
                $currentAudioRaw -notcontains $savedDevice.RawKey -and
                @($currentAudioStable | Where-Object { $_ -eq $savedDevice.StableKey }).Count -ne 1
            } |
            ForEach-Object { $_.Label }
        )
    }

    [ordered]@{
        compatible      = ($missing.Count -eq 0) -and ($scanErrors.Count -eq 0)
        missing_devices = $missing
        error           = $scanErrors -join '; '
    } | ConvertTo-Json -Compress
}
catch {
    [ordered]@{
        compatible      = $false
        missing_devices = @()
        error           = $_.Exception.Message
    } | ConvertTo-Json -Compress
}
