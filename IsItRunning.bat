@echo off
setlocal

set "EXEName=XLMonitor.exe"

tasklist /fi "imagename eq %EXEName%" /nh | find /i "%EXEName%" >nul

if %errorlevel% EQU 0 (
    echo It's running
) else (
    echo It's not running
)

pause