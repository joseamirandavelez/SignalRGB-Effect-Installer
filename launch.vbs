Set WShell = CreateObject("WScript.Shell")
ScriptPath = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
Cmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & ScriptPath & "installer.ps1"""
WShell.Run Cmd, 0, false
Set WShell = Nothing