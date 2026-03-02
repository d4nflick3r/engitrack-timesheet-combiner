@echo off
cd /d "%~dp0.."
python windows_launcher.py
if errorlevel 1 (
    echo.
    echo Failed to start. Run install.bat first if you haven't already.
    pause
)
