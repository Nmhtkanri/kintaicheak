@echo off
title Kintai Checker
pushd "%~dp0"
if errorlevel 1 (
    echo [ERROR] Failed to open the application folder.
    echo.
    pause
    exit /b 1
)

echo ============================================================
echo   Kintai Checker - Starting server...
echo ============================================================
echo.
echo   Browser will open automatically in a few seconds.
echo   Close this window to stop the server.
echo.
echo ============================================================
echo.

REM Use the bundled EXE first so users do not depend on local Python setup.
set "EXE=%~dp0dist\KintaiChecker\KintaiChecker.exe"
if exist "%EXE%" (
    "%EXE%"
    echo.
    echo ============================================================
    echo   Server stopped.
    echo ============================================================
    pause
    popd
    exit /b 0
)

REM Fallback for development/admin PCs when the EXE has not been built yet.
where python >nul 2>nul
if not errorlevel 1 (
    start "" /B cmd /c "timeout /t 4 /nobreak > nul && start http://localhost:5000"
    python app.py
    echo.
    echo ============================================================
    echo   Server stopped.
    echo ============================================================
    pause
    popd
    exit /b 0
)

echo [ERROR] KintaiChecker.exe was not found and Python is not installed.
echo.
echo Please ask the administrator to run build_exe.bat first.
echo.
pause
popd
exit /b 1
