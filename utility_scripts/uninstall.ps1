[CmdletBinding()]
Param(
	[Parameter(Mandatory = $true)]
	[string]$PC,
	[Parameter(Mandatory = $false)]
	[string]$standalone = "true"
)


# Uninstalls all cmdlets and removes all files created

# Workaround for passing in booleans from Python
if ($standalone -eq "true") {
	$standalone = $true
}
else {
	$standalone = $false
}

Import-Module (Join-Path $PSScriptRoot '..\installer_scripts\shared\SharedHelpers.psm1') -Force

$isLocal = Test-IsLocalComputer -ComputerName $PC

try {
	if ($standalone -and -not $isLocal) {
		if (-not (Test-HostReachable -ComputerName $PC -TimeoutMilliseconds 1000)) {
			throw "$PC is not reachable"
		}
	}

	Invoke-LocalOrRemote -ComputerName $PC -IsLocal $isLocal -ArgumentList @(
		'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp',
		'C:\ProgramData\CTS',
		'C:\Users\Public\Desktop'
	) -ScriptBlock {
		param(
			[Parameter(Mandatory = $true)]
			[string]$StartupFolder,
			[Parameter(Mandatory = $true)]
			[string]$CtsFolder,
			[Parameter(Mandatory = $true)]
			[string]$PublicDesktopPath
		)

		$ErrorActionPreference = 'Stop'

		function Get-MachineModuleRoots {
			$modulePaths = @()
			if (-not [string]::IsNullOrWhiteSpace($env:PSModulePath)) {
				$modulePaths = $env:PSModulePath -split ';' |
				ForEach-Object { $_.Trim() } |
				Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
				Select-Object -Unique
			}

			$roots = @()
			if (-not [string]::IsNullOrWhiteSpace($env:ProgramW6432)) {
				$roots += (Join-Path $env:ProgramW6432 'WindowsPowerShell\Modules')
			}
			if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
				$roots += (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules')
			}
			if (-not [string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) {
				$roots += (Join-Path ${env:ProgramFiles(x86)} 'WindowsPowerShell\Modules')
			}

			$roots += $modulePaths | Where-Object { $_ -match '\\Program Files( \(x86\))?\\WindowsPowerShell\\Modules$' }

			return ($roots | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
		}

		$moduleNames = @('AudioDeviceCmdlets', 'DisplayConfig')
		$machineModuleRoots = Get-MachineModuleRoots

		foreach ($moduleName in $moduleNames) {
			Get-Module -Name $moduleName -All -ErrorAction SilentlyContinue | Remove-Module -Force -ErrorAction SilentlyContinue

			$removedModule = $false
			foreach ($machineModuleRoot in $machineModuleRoots) {
				$moduleRoot = Join-Path $machineModuleRoot $moduleName
				if (-not (Test-Path -LiteralPath $moduleRoot)) {
					continue
				}

				Remove-Item -LiteralPath $moduleRoot -Recurse -Force -ErrorAction Stop
				Write-Output "INFO: Removed PowerShell module at $moduleRoot"
				$removedModule = $true
			}

			if (-not $removedModule) {
				Write-Output "INFO: PowerShell module already absent: $moduleName"
			}
		}

		$startupBatNames = @(
			'av_config_recall.bat',
			'cts_display_startup.bat',
			'cts_audio_startup.bat',
			'cts_bginfo_startup.bat',
			'dc.bat',
			'run_bgi.bat'
		)

		foreach ($startupBatName in $startupBatNames) {
			$startupBatPath = Join-Path $StartupFolder $startupBatName
			if (Test-Path -LiteralPath $startupBatPath) {
				Remove-Item -LiteralPath $startupBatPath -Force -ErrorAction Stop
				Write-Output "INFO: Removed Startup bat at $startupBatPath"
			}
			else {
				Write-Output "INFO: Startup bat already absent at $startupBatPath"
			}
		}

		if (Test-Path -LiteralPath $CtsFolder) {
			Remove-Item -LiteralPath $CtsFolder -Recurse -Force -ErrorAction Stop
			Write-Output "INFO: Removed CTS folder at $CtsFolder"
		}
		else {
			Write-Output "INFO: CTS folder already absent at $CtsFolder"
		}

		if (-not (Test-Path -LiteralPath $PublicDesktopPath)) {
			Write-Output "INFO: Public Desktop path not found at $PublicDesktopPath"
			return
		}

		$shortcutBaseNames = @('Log Out', 'Logout', 'Reboot')
		$existingShortcuts = Get-ChildItem -LiteralPath $PublicDesktopPath -File -Filter '*.lnk' -ErrorAction SilentlyContinue |
		Where-Object { $shortcutBaseNames -contains $_.BaseName }

		if (($null -eq $existingShortcuts) -or ($existingShortcuts.Count -eq 0)) {
			Write-Output "INFO: No CTS desktop shortcuts found in $PublicDesktopPath"
			return
		}

		foreach ($shortcut in $existingShortcuts) {
			Remove-Item -LiteralPath $shortcut.FullName -Force -ErrorAction Stop
			Write-Output "INFO: Removed desktop shortcut at $($shortcut.FullName)"
		}

		$desktopBatchNames = @(
			'SAVE_AV_SETTINGS.bat',
			'SAVE_AUDIO_SETTINGS.bat'
		)

		foreach ($desktopBatchName in $desktopBatchNames) {
			$desktopBatchPath = Join-Path $PublicDesktopPath $desktopBatchName
			if (Test-Path -LiteralPath $desktopBatchPath) {
				Remove-Item -LiteralPath $desktopBatchPath -Force -ErrorAction Stop
				Write-Output "INFO: Removed desktop batch file at $desktopBatchPath"
			}
			else {
				Write-Output "INFO: Desktop batch file already absent at $desktopBatchPath"
			}
		}
	}
}
catch {
	Write-Output "ERROR: Uninstall failed: $_"
	Exit 1
}

Exit 0
