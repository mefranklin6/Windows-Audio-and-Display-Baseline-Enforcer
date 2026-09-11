function Test-IsLocalComputer {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    $name = $ComputerName.Trim()

    return ($name -ieq $env:COMPUTERNAME) -or
    ($name -ieq 'localhost') -or
    ($name -ieq '.') -or
    ($name -ieq '127.0.0.1')
}

function Invoke-LocalOrRemote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory = $false)]
        [object[]]$ArgumentList,

        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [bool]$IsLocal
    )

    if ($IsLocal) {
        if ($null -ne $ArgumentList -and $ArgumentList.Count -gt 0) {
            return & $ScriptBlock @ArgumentList
        }
        return & $ScriptBlock
    }

    $params = @{
        ComputerName = $ComputerName
        ScriptBlock  = $ScriptBlock
        ErrorAction  = 'Stop'
    }
    if ($null -ne $ArgumentList -and $ArgumentList.Count -gt 0) {
        $params['ArgumentList'] = $ArgumentList
    }
    return Invoke-Command @params
}

function Test-HostReachable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutMilliseconds = 1000
    )

    try {
        $ping = [System.Net.NetworkInformation.Ping]::new()
        $reply = $ping.Send($ComputerName, $TimeoutMilliseconds)
        return $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success
    }
    catch {
        return $false
    }
}


function ConvertTo-NormalizedVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Version
    )

    $normalized = $Version.Trim()
    if ($normalized.StartsWith('v', [System.StringComparison]::OrdinalIgnoreCase)) {
        $normalized = $normalized.Substring(1)
    }

    if ([string]::IsNullOrWhiteSpace($normalized)) {
        throw "Invalid Version '$Version'. Expected numeric dotted version like '3.2' (optionally prefixed with 'v')."
    }

    # Require at least one dot to avoid accepting just '3'
    if ($normalized -notmatch '^\d+(\.\d+)+$') {
        throw "Invalid Version '$Version'. Expected numeric dotted version like '3.2' (optionally prefixed with 'v')."
    }

    return $normalized
}

function Set-ContentWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [object]$Value,

        [Parameter(Mandatory = $false)]
        [int]$MaxAttempts = 4,

        [Parameter(Mandatory = $false)]
        [int]$InitialDelayMilliseconds = 250
    )

    if ($MaxAttempts -lt 1) {
        throw 'MaxAttempts must be at least 1.'
    }

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Set-Content -LiteralPath $LiteralPath -Value $Value -Encoding ASCII -Force -ErrorAction Stop
            return
        }
        catch {
            if ($attempt -eq $MaxAttempts) {
                throw
            }

            $delay = $InitialDelayMilliseconds * [math]::Pow(2, $attempt - 1)
            Write-Output "WARNING: Write attempt $attempt of $MaxAttempts failed for $LiteralPath; retrying in $delay ms: $($_.Exception.Message)"
            Start-Sleep -Milliseconds $delay
        }
    }
}

function Copy-ItemWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$Destination,

        [Parameter(Mandatory = $false)]
        [int]$MaxAttempts = 4,

        [Parameter(Mandatory = $false)]
        [int]$InitialDelayMilliseconds = 250
    )

    if ($MaxAttempts -lt 1) {
        throw 'MaxAttempts must be at least 1.'
    }

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Copy-Item -LiteralPath $LiteralPath -Destination $Destination -Force -ErrorAction Stop
            return
        }
        catch {
            if ($attempt -eq $MaxAttempts) {
                throw
            }

            $delay = $InitialDelayMilliseconds * [math]::Pow(2, $attempt - 1)
            Write-Output "WARNING: Copy attempt $attempt of $MaxAttempts failed for $Destination; retrying in $delay ms: $($_.Exception.Message)"
            Start-Sleep -Milliseconds $delay
        }
    }
}


Export-ModuleMember -Function Test-IsLocalComputer, Invoke-LocalOrRemote, Test-HostReachable, ConvertTo-NormalizedVersion, Set-ContentWithRetry, Copy-ItemWithRetry
