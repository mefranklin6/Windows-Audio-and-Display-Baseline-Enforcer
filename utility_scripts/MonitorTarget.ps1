[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function New-BaseResult {
    param(
        [string]$Name,
        [bool]$Online,
        [bool]$WinRM,
        [string]$ErrorMessage = ''
    )

    [ordered]@{
        pc                       = $Name
        online                   = $Online
        winrm                    = $WinRM
        error                    = $ErrorMessage
        cts_deployed             = $false
        audio_configured         = $false
        display_configured       = $false
        logout_shortcut          = $false
        reboot_shortcut          = $false
        bginfo_deployed          = $false
        bginfo_executable        = $false
        bginfo_profile           = $false
        bginfo_background        = $false
        bginfo_startup           = $false
        bginfo_startup_method    = ''
        audio_device_cmdlets_versions = @()
        display_config_versions  = @()
        audio_configuration      = $null
        audio_configuration_error = ''
        display_configuration    = $null
        display_configuration_error = ''
    }
}

$name = $ComputerName.Trim()
$isLocal = ($name -ieq $env:COMPUTERNAME) -or
    ($name -ieq 'localhost') -or
    ($name -ieq '.') -or
    ($name -ieq '127.0.0.1')

$online = $isLocal
if (-not $isLocal) {
    try {
        $ping = [System.Net.NetworkInformation.Ping]::new()
        $reply = $ping.Send($name, 1500)
        $online = $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success
    }
    catch {
        $online = $false
    }
}

if (-not $online) {
    New-BaseResult -Name $name -Online $false -WinRM $false -ErrorMessage 'Ping test failed' |
        ConvertTo-Json -Compress -Depth 5
    exit 0
}

$inspection = {
    function Get-AdapterKey {
        param($Adapter)

        if ($null -eq $Adapter) {
            return ''
        }
        return '{0}:{1}' -f $Adapter.HighPart, $Adapter.LowPart
    }

    function Get-DisplayConfigurationSummary {
        param([string]$ProfilePath)

        $profile = Import-Clixml -LiteralPath $ProfilePath -ErrorAction Stop
        $paths = @($profile.PathArray)
        $modes = @($profile.ModeArray)
        $pathNames = @($profile.AvailablePathNames)
        $namesByTarget = @{}

        for ($index = 0; $index -lt $pathNames.Count; $index++) {
            $pathName = $pathNames[$index]
            $targetId = [int]$pathName.header.id
            $adapterKey = Get-AdapterKey $pathName.header.adapterId
            $key = '{0}:{1}' -f $adapterKey, $targetId
            $friendlyName = [string]$pathName.monitorFriendlyDeviceName
            $namesByTarget[$key] = [pscustomobject]@{
                number = $index + 1
                name   = if ([string]::IsNullOrWhiteSpace($friendlyName)) { 'Display {0}' -f ($index + 1) } else { $friendlyName }
            }
        }

        $activePaths = @(
            $paths | Where-Object { ([uint32]$_.flags -band 1) -eq 1 }
        )
        $monitors = @()
        $desktopKeys = @()

        for ($index = 0; $index -lt $activePaths.Count; $index++) {
            $path = $activePaths[$index]
            $sourceModeIndex = [int]$path.sourceInfo.SourceModeInfoIdx
            if (($sourceModeIndex -lt 0) -or ($sourceModeIndex -ge $modes.Count)) {
                $sourceModeIndex = [int]$path.sourceInfo.modeInfoIdx
            }
            if (($sourceModeIndex -lt 0) -or ($sourceModeIndex -ge $modes.Count)) {
                continue
            }

            $sourceMode = $modes[$sourceModeIndex].modeInfo.sourceMode
            $width = [int]$sourceMode.width
            $height = [int]$sourceMode.height
            if (($width -le 0) -or ($height -le 0)) {
                continue
            }

            $targetId = [int]$path.targetInfo.id
            $targetAdapterKey = Get-AdapterKey $path.targetInfo.adapterId
            $targetKey = '{0}:{1}' -f $targetAdapterKey, $targetId
            $nameRecord = $namesByTarget[$targetKey]
            $displayNumber = if ($null -eq $nameRecord) { $index + 1 } else { [int]$nameRecord.number }
            $displayName = if ($null -eq $nameRecord) { 'Display {0}' -f $displayNumber } else { [string]$nameRecord.name }
            $refreshRate = $null
            $numerator = [double]$path.targetInfo.refreshRate.Numerator
            $denominator = [double]$path.targetInfo.refreshRate.Denominator
            if ($denominator -gt 0) {
                $refreshRate = [math]::Round($numerator / $denominator, 2)
            }
            $rotation = switch ([uint32]$path.targetInfo.rotation) {
                2 { 90 }
                3 { 180 }
                4 { 270 }
                default { 0 }
            }
            $x = [int]$sourceMode.position.x
            $y = [int]$sourceMode.position.y
            $desktopKeys += '{0}:{1}' -f $x, $y
            $monitors += [pscustomobject][ordered]@{
                number       = $displayNumber
                name         = $displayName
                width        = $width
                height       = $height
                x            = $x
                y            = $y
                rotation     = $rotation
                refresh_rate = $refreshRate
                primary      = ($x -eq 0) -and ($y -eq 0)
            }
        }

        $mode = if ($monitors.Count -eq 0) {
            'No active displays'
        }
        elseif ($monitors.Count -eq 1) {
            if ($pathNames.Count -gt 1) {
                'Only show on display {0}' -f $monitors[0].number
            }
            else {
                'Single display'
            }
        }
        else {
            $uniqueDesktopCount = @($desktopKeys | Sort-Object -Unique).Count
            if ($uniqueDesktopCount -eq 1) {
                'Duplicate displays'
            }
            elseif ($uniqueDesktopCount -lt $monitors.Count) {
                'Extended desktop (some displays duplicated)'
            }
            else {
                'Extended desktop'
            }
        }

        [pscustomobject][ordered]@{
            mode         = $mode
            profile_path = $ProfilePath
            monitors     = $monitors
        }
    }

    $ctsFolder = 'C:\ProgramData\CTS'
    $startupFolder = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp'
    $publicDesktop = 'C:\Users\Public\Desktop'
    $ctsDeployed = Test-Path -LiteralPath $ctsFolder -PathType Container
    $bgInfoExecutable = $ctsDeployed -and
        (@(Get-ChildItem -LiteralPath $ctsFolder -File -Filter 'BGInfo64.exe' -ErrorAction SilentlyContinue).Count -gt 0)
    $bgInfoProfile = $ctsDeployed -and
        (@(Get-ChildItem -LiteralPath $ctsFolder -File -Filter '*.bgi' -ErrorAction SilentlyContinue).Count -gt 0)
    $bgInfoBackground = $false
    if ($ctsDeployed) {
        $bgInfoBackground = @(
            Get-ChildItem -LiteralPath $ctsFolder -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match '^\.(jpg|jpeg|png|bmp|gif)$' }
        ).Count -gt 0
    }
    $standaloneBgInfoStartup = Test-Path -LiteralPath (Join-Path $startupFolder 'cts_bginfo_startup.bat') -PathType Leaf
    $consolidatedBgInfoStartup = $false
    foreach ($avRecallPath in @(
        (Join-Path $startupFolder 'av_config_recall.bat'),
        (Join-Path $ctsFolder 'av_config_recall.bat')
    )) {
        if (-not (Test-Path -LiteralPath $avRecallPath -PathType Leaf)) {
            continue
        }
        $avRecallContent = Get-Content -LiteralPath $avRecallPath -Raw -ErrorAction SilentlyContinue
        if (($null -ne $avRecallContent) -and
            ($avRecallContent.IndexOf('C:\ProgramData\CTS\Bginfo64.exe', [System.StringComparison]::OrdinalIgnoreCase) -ge 0)) {
            $consolidatedBgInfoStartup = $true
            break
        }
    }
    $bgInfoStartup = $standaloneBgInfoStartup -or $consolidatedBgInfoStartup
    $bgInfoStartupMethod = if ($standaloneBgInfoStartup) {
        'cts_bginfo_startup.bat'
    }
    elseif ($consolidatedBgInfoStartup) {
        'av_config_recall.bat'
    }
    else {
        ''
    }
    $displayProfilePath = Join-Path $ctsFolder 'display_config_profile.xml'
    $audioLevelsPath = Join-Path $ctsFolder 'audio_levels.json'
    $audioDevicesPath = Join-Path $ctsFolder 'audio_device_list.json'
    $audioConfigured = $ctsDeployed -and
        (Test-Path -LiteralPath $audioLevelsPath -PathType Leaf) -and
        (Test-Path -LiteralPath $audioDevicesPath -PathType Leaf)
    $audioConfiguration = $null
    $audioConfigurationError = ''
    if ($audioConfigured) {
        try {
            $savedLevels = Get-Content -LiteralPath $audioLevelsPath -Raw | ConvertFrom-Json
            $savedDevices = @(Get-Content -LiteralPath $audioDevicesPath -Raw | ConvertFrom-Json)
            $audioConfiguration = [pscustomobject][ordered]@{
                levels = $savedLevels
                devices = @(
                    $savedDevices | ForEach-Object {
                        [pscustomobject][ordered]@{
                            Name = [string]$_.Name
                            Type = [string]$_.Type
                            Default = [bool]$_.Default
                            DefaultCommunication = [bool]$_.DefaultCommunication
                        }
                    }
                )
            }
        }
        catch {
            $audioConfigurationError = $_.Exception.Message
        }
    }
    $displayConfigured = $ctsDeployed -and (Test-Path -LiteralPath $displayProfilePath -PathType Leaf)
    $displayConfiguration = $null
    $displayConfigurationError = ''
    if ($displayConfigured) {
        try {
            $displayConfiguration = Get-DisplayConfigurationSummary -ProfilePath $displayProfilePath
        }
        catch {
            $displayConfigurationError = $_.Exception.Message
        }
    }

    [pscustomobject][ordered]@{
        cts_deployed             = $ctsDeployed
        audio_configured         = $audioConfigured
        display_configured       = $displayConfigured
        logout_shortcut          = Test-Path -LiteralPath (Join-Path $publicDesktop 'Log Out.lnk') -PathType Leaf
        reboot_shortcut          = Test-Path -LiteralPath (Join-Path $publicDesktop 'Reboot.lnk') -PathType Leaf
        bginfo_deployed          = $bgInfoExecutable -and $bgInfoProfile -and $bgInfoBackground -and $bgInfoStartup
        bginfo_executable        = $bgInfoExecutable
        bginfo_profile           = $bgInfoProfile
        bginfo_background        = $bgInfoBackground
        bginfo_startup           = $bgInfoStartup
        bginfo_startup_method    = $bgInfoStartupMethod
        audio_device_cmdlets_versions = @(
            Get-Module -ListAvailable -Name 'AudioDeviceCmdlets' -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty Version |
                Sort-Object -Descending -Unique |
                ForEach-Object { $_.ToString() }
        )
        display_config_versions  = @(
            Get-Module -ListAvailable -Name 'DisplayConfig' -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty Version |
                Sort-Object -Descending -Unique |
                ForEach-Object { $_.ToString() }
        )
        audio_configuration      = $audioConfiguration
        audio_configuration_error = $audioConfigurationError
        display_configuration    = $displayConfiguration
        display_configuration_error = $displayConfigurationError
    }
}

try {
    if ($isLocal) {
        $details = & $inspection
    }
    else {
        $details = Invoke-Command -ComputerName $name -ScriptBlock $inspection -ErrorAction Stop
    }

    $result = New-BaseResult -Name $name -Online $true -WinRM $true
    foreach ($property in $details.PSObject.Properties) {
        if ($result.Contains($property.Name)) {
            $result[$property.Name] = $property.Value
        }
    }
    $result | ConvertTo-Json -Compress -Depth 8
}
catch {
    New-BaseResult -Name $name -Online $true -WinRM $false -ErrorMessage $_.Exception.Message |
        ConvertTo-Json -Compress -Depth 8
}
