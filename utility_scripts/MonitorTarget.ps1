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

    [pscustomobject][ordered]@{
        cts_deployed             = $ctsDeployed
        audio_configured         = $ctsDeployed -and (Test-Path -LiteralPath (Join-Path $ctsFolder 'audio_levels.json') -PathType Leaf)
        display_configured       = $ctsDeployed -and (Test-Path -LiteralPath (Join-Path $ctsFolder 'display_config_profile.xml') -PathType Leaf)
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
    $result | ConvertTo-Json -Compress -Depth 5
}
catch {
    New-BaseResult -Name $name -Online $true -WinRM $false -ErrorMessage $_.Exception.Message |
        ConvertTo-Json -Compress -Depth 5
}
