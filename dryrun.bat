@echo off
cd /d "%~dp0"
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File "%~dp0lastprice.ps1"
pause
