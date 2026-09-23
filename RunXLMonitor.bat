@echo off
set "EXEName=XLMonitor.exe"
set "EXEFullPath=C:\XLMonitor\XLMonitor.exe"

tasklist | find /i "%EXEName%" >nul
if errorlevel 1 (
    start "" /b "%EXEFullPath%" >output.txt 2>error.txt
)