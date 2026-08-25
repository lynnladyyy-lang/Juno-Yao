@echo off
setlocal
rem ============================================================
rem  word2md.bat  -  Word to Markdown converter (web UI)
rem  Double-click to start the local server and open the browser.
rem  (Kept ASCII-only to avoid Windows codepage issues.)
rem ============================================================

rem Kill any lingering old server processes running word2md_server.py
powershell -NoProfile -Command "Get-WmiObject Win32_Process | Where-Object { $_.CommandLine -like '*word2md_server.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" 2>nul

timeout /t 1 /nobreak >nul 2>nul

set "PY=C:\Users\Lynn\.workbuddy\binaries\python\envs\default\Scripts\pythonw.exe"
if not exist "%PY%" set "PY=C:\Users\Lynn\.workbuddy\binaries\python\versions\3.13.12\pythonw.exe"
if not exist "%PY%" set "PY=C:\Users\Lynn\.workbuddy\versions\3.13.12\python.exe"

start "" "%PY%" "%~dp0word2md_server.py"

endlocal
