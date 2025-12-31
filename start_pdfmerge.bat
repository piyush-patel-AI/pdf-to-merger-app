@echo off
title PDF Merge Server
echo =====================================
echo     🚀 Starting PDF Merge App...
echo =====================================
echo.

REM Activate your Python environment if needed (uncomment below)
REM call venv\Scripts\activate

REM Start Flask server
start "" cmd /k "py -3.11 app.py"

REM Wait a few seconds for server to start
timeout /t 3 >nul

REM Open in browser (localhost version)
start http://127.0.0.1:5000

echo.
echo ✅ PDF Merge Server launched successfully!
echo Close this window to stop the server.
pause
