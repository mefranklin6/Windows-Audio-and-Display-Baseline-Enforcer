$ErrorActionPreference = 'Stop'

$CtsFolder = 'C:\ProgramData\CTS'
$DeviceListPath = Join-Path $CtsFolder 'audio_device_list.json'
$LevelsPath = Join-Path $CtsFolder 'audio_levels.json'

if (-not (Test-Path -LiteralPath $CtsFolder)) {
    New-Item -ItemType Directory -Path $CtsFolder -Force | Out-Null
}

Get-AudioDevice -List | ConvertTo-Json | Out-File -FilePath $DeviceListPath -Force

$levels = [PSCustomObject]@{
    PlaybackCommunicationVolume  = (Get-AudioDevice -PlaybackCommunicationVolume)
    PlaybackVolume               = (Get-AudioDevice -PlaybackVolume)
    RecordingCommunicationVolume = (Get-AudioDevice -RecordingCommunicationVolume)
    RecordingVolume              = (Get-AudioDevice -RecordingVolume)
}

$levels | ConvertTo-Json | Out-File -FilePath $LevelsPath -Force

Write-Output "Audio settings saved to $CtsFolder"
