@echo off
set "EXEName=XLMonitor.exe"
set "EXEFullPath=C:\XLMonitor\XLMonitor.exe"

tasklist | find /i "%EXEName%" >nul
if errorlevel 1 (
    powershell -NoProfile -Command "Start-Process -FilePath '%EXEFullPath%' -WindowStyle Hidden -RedirectStandardOutput 'output.txt' -RedirectStandardError 'error.txt'"
)