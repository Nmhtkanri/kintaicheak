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

REM Python が利用可能なら python app.py を優先する。
REM 開発中の変更（quick_compare / quick_export など）も即反映される。
where python >nul 2>nul
if not errorlevel 1 (
    start "" /B cmd /c "timeout /t 4 /nobreak > nul && start http://localhost:5000"
    python app.py
    echo.
    echo ============================================================
    echo   Server stopped.
    echo ============================================================
    pause
    exit /b 0
)

REM Python が無ければ PyInstaller でビルド済みの EXE をフォールバックで使用。
set "EXE=%~dp0dist\KintaiChecker\KintaiChecker.exe"
if exist "%EXE%" (
    echo [info] Python is not installed. Falling back to bundled EXE.
    echo        Note: the EXE may be stale -- run build_exe.bat to rebuild if needed.
    echo.
    "%EXE%"
    echo.
    echo ============================================================
    echo   Server stopped.
    echo ============================================================
    pause
    exit /b 0
)

echo [ERROR] Python is not installed and KintaiChecker.exe was not found.
echo.
echo Please install Python, or ask the administrator to run build_exe.bat first.
echo.
pause
exit /b 1
