@echo off
set "SCRIPT=%~dp0sync-onenote.ps1"
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -STA -ExecutionPolicy Bypass -File "%SCRIPT%" %*
