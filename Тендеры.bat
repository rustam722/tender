@echo off
rem ASCII only: cmd reads .bat in the OEM codepage, Cyrillic would be garbled.
cd /d "%~dp0"
set URL=http://127.0.0.1:8765

if not exist venv\Scripts\python.exe (
    echo [ERROR] venv not found. Create it: python -m venv venv
    pause
    exit /b 1
)

venv\Scripts\python -c "import importlib.util,sys; sys.exit(importlib.util.find_spec('app.web') is None)"
if errorlevel 1 (
    echo [ERROR] Module app.web not found - the web server is not built yet.
    pause
    exit /b 1
)

echo Server address: %URL%
echo Close this window to stop the program.
rem open the browser after a short pause, while the server starts
start "" /min cmd /c "ping -n 3 127.0.0.1 >nul && start %URL%"
venv\Scripts\python -m app.web
if errorlevel 1 pause
