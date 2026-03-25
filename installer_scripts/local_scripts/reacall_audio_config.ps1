$ErrorActionPreference = 'Stop'

# --- Config ---
$JsonPath = 'C:\ProgramData\CTS\audio_device_list.json'
$LevelsJsonPath = 'C:\ProgramData\CTS\audio_levels.json'

# Log lives next to this script
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$LogPath = Join-Path $ScriptDir 'AudioDeviceStartup.log'

function Write-Log {
    param(
        [Parameter(Mandatory)][string] $Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string] $Level = 'INFO'
    )

    try {
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        Add-Content -LiteralPath $LogPath -Value "$ts [$Level] $Message" -Encoding UTF8
    }
    catch {
        # Absolute last resort: never show anything to user
    }
}

function ConvertTo-NormalizedAudioDeviceListJson {
    param([Parameter(Mandatory)][object[]] $Devices)

    $norm = foreach ($d in $Devices) {
        [pscustomobject]@{
            Index                = [int]($d.Index)
            Type                 = [string]($d.Type)
            Name                 = [string]($d.Name)
            ID                   = [string]($d.ID)
            Default              = [bool]($d.Default)
            DefaultCommunication = [bool]($d.DefaultCommunication)
        }
    }

    ($norm | Sort-Object Type, Name, ID | ConvertTo-Json -Depth 5)
}

function Get-TargetDefaultsFromFile {
    param([Parameter(Mandatory)][object[]] $FileDevices)

    foreach ($d in $FileDevices) {
        $isDef = [bool]($d.Default)
        $isComm = [bool]($d.DefaultCommunication)

        if ($isDef -or $isComm) {
            [pscustomobject]@{
                Type                 = [string]$d.Type   # "Playback" or "Recording"
                Name                 = [string]$d.Name
                ID                   = [string]$d.ID
                Default              = $isDef
                DefaultCommunication = $isComm
            }
        }
    }
}

function Resolve-DeviceStrict {
    param(
        [Parameter(Mandatory)][object[]] $CurrentDevices,
        [Parameter(Mandatory)][string] $Type,
        [Parameter(Mandatory)][string] $SavedId,
        [Parameter(Mandatory)][string] $SavedName
    )

    # 1) ID exact match only
    $idMatch = $CurrentDevices | Where-Object {
        $_.Type -eq $Type -and $_.ID -eq $SavedId
    } | Select-Object -First 1

    if ($idMatch) { return $idMatch }

    # 2) ONLY if no ID match exists: Name exact match (case-insensitive exact)
    $nameMatch = $CurrentDevices | Where-Object {
        $_.Type -eq $Type -and $_.Name -eq $SavedName
        # For case-sensitive exact match, use:
        # $_.Type -eq $Type -and $_.Name -ceq $SavedName
    } | Select-Object -First 1

    return $nameMatch
}

function Restore-AudioLevels {
    param([Parameter(Mandatory)][string] $LevelsPath)

    if (-not (Test-Path -LiteralPath $LevelsPath)) {
        Write-Log "Audio levels JSON not found: $LevelsPath" "WARN"
        return
    }

    try {
        $levelsRaw = Get-Content -LiteralPath $LevelsPath -Raw
        $levels = $levelsRaw | ConvertFrom-Json
    }
    catch {
        Write-Log "Failed to parse audio levels JSON: $($_.Exception.Message)" "ERROR"
        return
    }

    # Map JSON keys to Set-AudioDevice parameter names
    $paramMap = @{
        'PlaybackVolume'               = 'PlaybackVolume'
        'PlaybackCommunicationVolume'  = 'PlaybackCommunicationVolume'
        'RecordingVolume'              = 'RecordingVolume'
        'RecordingCommunicationVolume' = 'RecordingCommunicationVolume'
    }

    foreach ($key in $paramMap.Keys) {
        $rawValue = $levels.$key
        if ($null -eq $rawValue) { continue }

        # Strip '%' and any whitespace, then convert to a rounded integer
        $cleaned = ($rawValue -replace '[%\s]', '').Trim()
        $parsed = [double]$cleaned
        $intValue = [int][Math]::Round($parsed)

        # Clamp to 0-100
        if ($intValue -lt 0) { $intValue = 0 }
        if ($intValue -gt 100) { $intValue = 100 }

        $paramName = $paramMap[$key]
        Write-Log "Setting $paramName -> $intValue (raw='$rawValue')"

        try {
            $setArgs = @{ $paramName = $intValue }
            Set-AudioDevice @setArgs | Out-Null
        }
        catch {
            Write-Log "Failed to set ${paramName}: $($_.Exception.Message)" "ERROR"
        }
    }

    Write-Log "Done applying audio levels."
}

function Restore-DefaultDevices {
    if (-not (Test-Path -LiteralPath $JsonPath)) {
        Write-Log "JSON file not found: $JsonPath" "ERROR"
        return
    }

    $savedRaw = Get-Content -LiteralPath $JsonPath -Raw
    $savedDevices = $savedRaw | ConvertFrom-Json

    if (-not $savedDevices) {
        Write-Log "No devices found in JSON: $JsonPath" "WARN"
        return
    }

    $currentDevices = Get-AudioDevice -List

    if (-not $currentDevices) {
        Write-Log "No current audio devices found." "WARN"
        return
    }

    $savedNorm = ConvertTo-NormalizedAudioDeviceListJson -Devices @($savedDevices)
    $currentNorm = ConvertTo-NormalizedAudioDeviceListJson -Devices @($currentDevices)

    if ($savedNorm -eq $currentNorm) {
        Write-Log "No I/O device change needed (current list matches saved list)."
        return
    }

    Write-Log "Device list differs; applying saved default selections..."

    $targets = Get-TargetDefaultsFromFile -FileDevices @($savedDevices)

    foreach ($t in $targets) {
        $resolved = Resolve-DeviceStrict -CurrentDevices @($currentDevices) `
            -Type $t.Type -SavedId $t.ID -SavedName $t.Name

        if (-not $resolved) {
            Write-Log "No match for [$($t.Type)] Name='$($t.Name)' ID='$($t.ID)'. Skipping." "WARN"
            continue
        }

        if ($t.Default) {
            Write-Log "Setting DEFAULT [$($t.Type)] -> '$($resolved.Name)' (ID=$($resolved.ID))"
            Set-AudioDevice -ID $resolved.ID -DefaultOnly | Out-Null
        }

        if ($t.DefaultCommunication) {
            Write-Log "Setting COMM DEFAULT [$($t.Type)] -> '$($resolved.Name)' (ID=$($resolved.ID))"
            Set-AudioDevice -ID $resolved.ID -CommunicationOnly | Out-Null
        }
    }

    Write-Log "Done applying device I/O configuration."
}

# --- Main (silent) ---
try {
    Restore-DefaultDevices
    Restore-AudioLevels -LevelsPath $LevelsJsonPath
}
catch {
    Write-Log "Unhandled error: $($_.Exception.Message)" "ERROR"
    exit 1
}