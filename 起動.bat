@echo off
title Kintai Checker
cd /d "%~dp0"

echo ============================================================
echo   Kintai Checker - Starting server...
echo ============================================================
echo.
echo   Browser will open automatically in a few seconds.
echo   Close this window to stop the server.
echo.
echo ============================================================
echo.

set "EXE=%~dp0dist\KintaiChecker\KintaiChecker.exe"
if exist "%EXE%" (
    "%EXE%"
    echo.
    echo ============================================================
    echo   Server stopped.
    echo ============================================================
    pause
    exit /b 0
)

where python >nul 2>nul
if errorlevel 1 goto :no_python

start "" /B cmd /c "timeout /t 4 /nobreak > nul && start http://localhost:5000"

python app.py

echo.
echo ============================================================
echo   Server stopped.
echo ============================================================
pause
exit /b 0

:no_python
echo [ERROR] KintaiChecker.exe was not found and Python is not installed.
echo.
echo Please ask the administrator to run build_exe.bat first.
echo.
pause
exit /b 1
