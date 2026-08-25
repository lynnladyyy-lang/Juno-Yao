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

rem Locate Python: prefer the bundled interpreter (if present), otherwise PATH
set "PY="
set "BUNDLED=%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\pythonw.exe"
if exist "%BUNDLED%" set "PY=%BUNDLED%"

if not defined PY (
    where pythonw >nul 2>nul
    if not errorlevel 1 set "PY=pythonw"
)
if not defined PY (
    where python >nul 2>nul
    if not errorlevel 1 set "PY=python"
)

if not defined PY (
    echo [ERROR] Python 3 not found.
    echo Install Python 3 from https://www.python.org/downloads/ and check "Add python.exe to PATH".
    echo Then run:  pip install -r requirements.txt
    pause
    exit /b 1
)

start "" "%PY%" "%~dp0word2md_server.py"

endlocal
