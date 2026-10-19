@echo off
setlocal
cd /d "%~dp0"
set "APP=%~dp0app.py"
set "LAUNCHER=%~dp0launch_gui.pyw"

rem Prefer the windowed Python launcher. launch_gui.pyw writes startup errors to logs\gui_error.log.
where pyw.exe >nul 2>nul
if not errorlevel 1 (
    pyw -3 "%LAUNCHER%"
    goto :eof
)

where pythonw.exe >nul 2>nul
if not errorlevel 1 (
    pythonw.exe "%LAUNCHER%"
    goto :eof
)

rem Last resort: use a console Python so startup errors remain visible.
where py.exe >nul 2>nul
if not errorlevel 1 (
    py -3 "%LAUNCHER%"
    goto :eof
)

where python.exe >nul 2>nul
if not errorlevel 1 (
    python "%LAUNCHER%"
    goto :eof
)

echo [ERROR] Python 3 was not found.
echo Install Python 3.10 or newer from https://www.python.org/downloads/windows/
echo During installation, enable the Python Launcher option.
pause
