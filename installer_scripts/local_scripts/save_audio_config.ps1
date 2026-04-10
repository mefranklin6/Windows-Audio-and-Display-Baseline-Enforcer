$ErrorActionPreference = 'Stop'

$CtsFolder = 'C:\ProgramData\CTS'
$DeviceListPath = Join-Path $CtsFolder 'audio_device_list.json'
$LevelsPath = Join-Path $CtsFolder 'audio_levels.json'

function Get-AudioDeviceValueOrNull {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Expression
    )

    try {
        $value = & $Expression
        if ($null -ne $value) {
            return $value
        }
    }
    catch {
        Write-Output "No value found for $expression"
    }

    return $null
}

$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
$isAdministrator = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdministrator) {
    Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList @(
        '-NoProfile'
        '-ExecutionPolicy'
        'Bypass'
        '-File'
        ('"{0}"' -f $PSCommandPath)
    ) | Out-Null
    exit 0
}

if (-not (Test-Path -LiteralPath $CtsFolder)) {
    New-Item -ItemType Directory -Path $CtsFolder -Force | Out-Null
}

Get-AudioDevice -List | ConvertTo-Json | Out-File -FilePath $DeviceListPath -Force

$levels = [PSCustomObject]@{
    PlaybackCommunicationVolume  = Get-AudioDeviceValueOrNull { Get-AudioDevice -PlaybackCommunicationVolume }
    PlaybackVolume               = Get-AudioDeviceValueOrNull { Get-AudioDevice -PlaybackVolume }
    RecordingCommunicationVolume = Get-AudioDeviceValueOrNull { Get-AudioDevice -RecordingCommunicationVolume }
    RecordingVolume              = Get-AudioDeviceValueOrNull { Get-AudioDevice -RecordingVolume }
}

$levels | ConvertTo-Json | Out-File -FilePath $LevelsPath -Force

Write-Output "Audio settings saved to $CtsFolder"
