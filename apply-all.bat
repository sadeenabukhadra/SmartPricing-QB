@echo off
cd /d "%~dp0"
REM TEST COPY ONLY: writes price levels for ALL active customers
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File "%~dp0lastprice.ps1" -Apply -AllowFile "C:\Users\Lenovo\Desktop\petragate sadeen.qbw"
pause
