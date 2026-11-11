@echo off
cd /d "%~dp0"
start "" C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0prepare.ps1" -AllowFile "C:\Users\Lenovo\Desktop\petragate sadeen.qbw"
