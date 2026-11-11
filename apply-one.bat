@echo off
cd /d "%~dp0"
REM TEST COPY ONLY: writes the price level for ONE customer (Hareer Hotel)
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File "%~dp0lastprice.ps1" -Apply -AllowFile "C:\Users\Lenovo\Desktop\petragate sadeen.qbw" -OnlyCustomer "Hareer Hotel"
pause
