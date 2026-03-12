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


Export-ModuleMember -Function Test-IsLocalComputer, Invoke-LocalOrRemote, Test-HostReachable, ConvertTo-NormalizedVersion
