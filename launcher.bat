@echo off
REM This command launches the PowerShell GUI script in a hidden window.
REM The -ExecutionPolicy Bypass allows the script to run, and -WindowStyle Hidden keeps the console from appearing.
powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0\installer.ps1"
