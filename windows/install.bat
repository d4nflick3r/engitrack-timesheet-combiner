@echo off
echo Installing EngiTrack Timesheet Combiner dependencies...
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python 3.11+ from https://python.org and try again.
    pause
    exit /b 1
)

python -m pip install --upgrade pip
python -m pip install streamlit openpyxl pandas pillow streamlit-desktop-app pywebview

echo.
echo Installation complete. Run run.bat to start the app.
pause
