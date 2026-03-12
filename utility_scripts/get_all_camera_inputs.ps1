$ErrorActionPreference = 'Stop'
$root = Join-Path $PSScriptRoot ..
$targets = Get-Content (Join-Path $root 'targets.txt') | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique

$throttleLimit = 50
$pingTimeoutMs = 1000

# ── Step 1: Concurrent ping sweep using .NET async ──────────────────────────
Write-Host "Pinging $($targets.Count) targets concurrently..."

$pingTasks = [ordered]@{}
foreach ($t in $targets) {
    $ping = New-Object System.Net.NetworkInformation.Ping
    $pingTasks[$t] = @{ Ping = $ping; Task = $ping.SendPingAsync($t, $pingTimeoutMs) }
}

try { [System.Threading.Tasks.Task]::WaitAll(@($pingTasks.Values.Task)) }
catch { <# AggregateException if any task faulted; handled per-target below #> }

$reachable = @()
foreach ($t in $targets) {
    $entry = $pingTasks[$t]
    try { $entry.Ping.Dispose() } catch {}

    if ($entry.Task.Status -eq 'RanToCompletion' -and
        $entry.Task.Result.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
        $reachable += $t
    }
    else {
        Write-Warning "Cannot ping $t; skipping entirely"
    }
}

Write-Host "$($reachable.Count)/$($targets.Count) targets reachable"

if ($reachable.Count -eq 0) {
    Write-Warning 'No reachable targets; exiting.'
    return
}

# ── Step 2: Query reachable targets via PowerShell Remoting ──────────────────
$invokeErrors = @()

$results = @(
    Invoke-Command -ComputerName $reachable -ThrottleLimit $throttleLimit -ErrorAction Continue -ErrorVariable +invokeErrors -ScriptBlock {
        $devices = @()

        # Prefer Get-PnpDevice when available
        if (Get-Command -Name Get-PnpDevice -ErrorAction SilentlyContinue) {
            try { $devices = Get-PnpDevice -Class Camera -ErrorAction Stop }
            catch { $devices = @() }
        }

        # Fallback inside the remote session (works even if PnpDevice module isn't present)
        if (-not $devices -or $devices.Count -eq 0) {
            try { $devices = Get-CimInstance -ClassName Win32_PnPEntity -Filter "PNPClass='Camera'" -ErrorAction Stop }
            catch { $devices = @() }
        }

        foreach ($d in $devices) {
            $name = $null
            if ($d.PSObject.Properties.Match('FriendlyName').Count -gt 0) {
                $name = $d.FriendlyName
            }
            if (-not $name) { $name = $d.Name }

            if ($name) {
                [pscustomobject]@{
                    ComputerName = $env:COMPUTERNAME
                    DeviceName   = $name
                }
            }
        }
    }
)

# ── Step 3: CIM/DCOM fallback for remoting failures (already known reachable) ─
function Get-InvokeCommandComputerName {
    param(
        [Parameter(Mandatory)]
        $ErrorRecord
    )

    if ($ErrorRecord.PSObject.Properties.Match('PSComputerName').Count -gt 0 -and $ErrorRecord.PSComputerName) {
        return [string]$ErrorRecord.PSComputerName
    }

    if ($ErrorRecord.PSObject.Properties.Match('OriginInfo').Count -gt 0 -and $ErrorRecord.OriginInfo -and $ErrorRecord.OriginInfo.PSComputerName) {
        return [string]$ErrorRecord.OriginInfo.PSComputerName
    }

    if ($ErrorRecord.TargetObject -is [string] -and $ErrorRecord.TargetObject) {
        return [string]$ErrorRecord.TargetObject
    }

    return $null
}

$failedTargets = @(
    $invokeErrors |
    ForEach-Object { Get-InvokeCommandComputerName -ErrorRecord $_ } |
    Where-Object { $_ } |
    Sort-Object -Unique
)

# No re-ping needed — these targets already passed the upfront ping sweep
if ($failedTargets.Count -gt 0) {
    foreach ($pc in $failedTargets) {
        Write-Warning "Invoke-Command failed on $pc; trying CIM/DCOM fallback"
    }

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $cimResults = $failedTargets | ForEach-Object -Parallel {
            $pc = $_
            try {
                $opt = New-CimSessionOption -Protocol Dcom
                $session = New-CimSession -ComputerName $pc -SessionOption $opt -ErrorAction Stop
                try {
                    Get-CimInstance -CimSession $session -ClassName Win32_PnPEntity -Filter "PNPClass='Camera'" |
                    ForEach-Object {
                        [pscustomobject]@{
                            ComputerName = $pc
                            DeviceName   = $_.Name
                        }
                    }
                }
                finally {
                    Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue
                }
            }
            catch {
                Write-Warning "CIM fallback failed on $pc : $($_.Exception.Message)"
            }
        } -ThrottleLimit $throttleLimit
        $results += @($cimResults)
    }
    else {
        foreach ($pc in $failedTargets) {
            try {
                $opt = New-CimSessionOption -Protocol Dcom
                $session = New-CimSession -ComputerName $pc -SessionOption $opt -ErrorAction Stop
                try {
                    $cim = Get-CimInstance -CimSession $session -ClassName Win32_PnPEntity -Filter "PNPClass='Camera'"
                    $results += $cim | ForEach-Object {
                        [pscustomobject]@{
                            ComputerName = $pc
                            DeviceName   = $_.Name
                        }
                    }
                }
                finally {
                    Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue
                }
            }
            catch {
                Write-Warning "CIM fallback failed on $pc : $($_.Exception.Message)"
            }
        }
    }
}

# ── Output ───────────────────────────────────────────────────────────────────
$unique = $results | Where-Object { $_.DeviceName } | Select-Object -ExpandProperty DeviceName | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique

[pscustomobject]@{
    Targets       = $targets -join ', '
    Reachable     = $reachable.Count
    Unreachable   = ($targets.Count - $reachable.Count)
    ThrottleLimit = $throttleLimit
    ResultCount   = $results.Count
    UniqueDevices = $unique.Count
} | Format-List

'--- Unique camera device names ---'
$unique

# Build per-PC JSON output
$perPc = $results |
Where-Object { $_.ComputerName -and $_.DeviceName } |
Group-Object -Property ComputerName |
ForEach-Object {
    [pscustomobject]@{
        ComputerName = $_.Name
        CameraInputs = @($_.Group | ForEach-Object { $_.DeviceName.Trim() } | Sort-Object -Unique)
    }
} | Sort-Object -Property ComputerName

$jsonPath = Join-Path $PSScriptRoot 'camera_inputs.json'
$perPc | ConvertTo-Json -Depth 3 | Set-Content -Path $jsonPath -Encoding UTF8
Write-Host "JSON written to $jsonPath"