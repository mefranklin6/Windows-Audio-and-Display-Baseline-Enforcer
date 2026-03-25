Param(
    [Parameter(Mandatory = $true)]
    [string]$PC,
    [Parameter(Mandatory = $false)]
    [string]$standalone = "true"
)

function Get-SingleMatchingFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,
        [Parameter(Mandatory = $true)]
        [string[]]$Patterns,
        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $matches_ = foreach ($pattern in $Patterns) {
        Get-ChildItem -LiteralPath $Directory -File -Filter $pattern -ErrorAction Stop
    }

    $matches_ = $matches_ | Sort-Object FullName -Unique

    if ($matches_.Count -eq 0) {
        $patternList = $Patterns -join ', '
        throw "No $Description found in $Directory matching: $patternList"
    }

    if ($matches_.Count -gt 1) {
        $matchList = ($matches_ | Select-Object -ExpandProperty Name) -join ', '
        throw "Expected exactly one $Description in $Directory, but found $($matches_.Count): $matchList"
    }

    return $matches_[0]
}

# Workaround for passing in booleans from Python
if ($standalone -eq "true") {
    $standalone = $true
}
else {
    $standalone = $false
}

$bginfoFolder = "25_26" # Within \BGInfo

Import-Module (Join-Path $PSScriptRoot 'shared\SharedHelpers.psm1') -Force

# Check if running locally
$isLocal = Test-IsLocalComputer -ComputerName $PC

Write-Output "INFO: Installing BGInfo and scripts"

try {

    if ($standalone -and -not $isLocal) {
        if (-not (Test-HostReachable -ComputerName $PC -TimeoutMilliseconds 1000)) {
            throw "$PC is not reachable"
        }
    }


    # Create C:\ProgramData\CTS folder if it doesn't exist
    if ($standalone) {
        Invoke-LocalOrRemote -ComputerName $PC -IsLocal $isLocal -ScriptBlock {
            if (!(Test-Path -Path 'C:\ProgramData\CTS')) {
                New-Item -ItemType Directory -Path 'C:\ProgramData\CTS' -Force | Out-Null
            }
        } | Out-Null
    }


    if ($isLocal) { $prefix = "C:\" }
    else { $prefix = "\\$PC\C$\" }

    # --- Repo source paths ---
    $repo_bginfo_root = Join-Path $PSScriptRoot '..\BGInfo'
    $repo_bginfo_dir = Join-Path $repo_bginfo_root $bginfoFolder

    if (-not (Test-Path -LiteralPath $repo_bginfo_dir)) {
        throw "$PC BGInfo folder not found: $repo_bginfo_dir"
    }

    $repo_exe = Get-SingleMatchingFile -Directory $repo_bginfo_dir -Patterns @('BGInfo64.exe') -Description 'BGInfo executable'
    $repo_config = Get-SingleMatchingFile -Directory $repo_bginfo_dir -Patterns @('*.bgi') -Description 'BGInfo config'
    $repo_image = Get-SingleMatchingFile -Directory $repo_bginfo_dir -Patterns @('*.jpg', '*.jpeg', '*.png', '*.bmp', '*.gif') -Description 'BGInfo background image'

    $repo_exe_path = $repo_exe.FullName
    $repo_config_path = $repo_config.FullName
    $repo_image_path = $repo_image.FullName

    # Fail fast if repo assets are missing
    $repoRequired = @(
        @{ Path = $repo_exe_path; Name = "$($repo_exe.Name) (repo)" },
        @{ Path = $repo_config_path; Name = "$($repo_config.Name) (repo)" },
        @{ Path = $repo_image_path; Name = "$($repo_image.Name) (repo)" }
    )
    foreach ($f in $repoRequired) {
        if (-not (Test-Path -LiteralPath $f.Path)) {
            throw "$PC Missing required repo file: $($f.Name) at $($f.Path)"
        }
    }

    # --- Destination paths on target ---
    $local_exe_path_suffix = Join-Path 'ProgramData\CTS' $repo_exe.Name
    $local_config_path_suffix = Join-Path 'ProgramData\CTS' $repo_config.Name
    $local_bat_path_suffix = "ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\cts_bginfo_startup.bat"
    $local_image_path_suffix = Join-Path 'ProgramData\CTS' $repo_image.Name

    $local_exe_path = Join-Path $prefix $local_exe_path_suffix
    $local_config_path = Join-Path $prefix $local_config_path_suffix
    $local_bat_path = Join-Path $prefix $local_bat_path_suffix
    $local_image_path = Join-Path $prefix $local_image_path_suffix

    $bginfoArgs = "'C:\$local_config_path_suffix' /timer:0 /nolicprompt /silent" 
    $startupBatContent = @(
        '@echo off',
        ('start "" "{0}" {1}' -f "C:\$local_exe_path_suffix", $bginfoArgs)
    )

    # Copy all files

    Copy-Item -Path $repo_image_path -Destination $local_image_path -Force

    Copy-Item -Path $repo_exe_path -Destination $local_exe_path -Force

    if (Test-Path $local_config_path) { Remove-Item $local_config_path -Force }
    Copy-Item -Path $repo_config_path -Destination $local_config_path -Force


    # This script owns its own Startup bat. Consolidation/ordering is handled by Cleanup.ps1.
    Set-Content -LiteralPath $local_bat_path -Value $startupBatContent -Encoding ASCII -Force

    # Verify Files
    $requiredFiles = @(
        @{ Path = $local_exe_path; Name = $repo_exe.Name },
        @{ Path = $local_config_path; Name = $repo_config.Name },
        @{ Path = $local_bat_path; Name = "cts_bginfo_startup.bat" },
        @{ Path = $local_image_path; Name = $repo_image.Name }
    )

    foreach ($file in $requiredFiles) {
        if (Test-Path $file.Path) {
            continue
        }
        else {
            throw "$PC $($file.Name) not found at $($file.Path)"
        }
    }

    Write-Output "INFO: BGInfo installation complete"



} # end try
catch {
    Write-Output "ERROR: InstallBGInfo failed: $_"
    Exit 1
}

Start-Sleep 1
Exit 0