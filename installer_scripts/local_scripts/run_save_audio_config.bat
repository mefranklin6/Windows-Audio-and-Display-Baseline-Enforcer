@echo off
setlocal EnableExtensions

net session >nul 2>&1
if %errorlevel% neq 0 (
	Powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
	exit /b
)

Powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File C:\ProgramData\CTS\save_audio_config.ps1
