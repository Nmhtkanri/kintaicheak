@echo off
title Build Kintai Checker EXE
cd /d "%~dp0"

echo ============================================================
echo   Kintai Checker - Build portable EXE
echo ============================================================
echo.

where python >nul 2>nul
if errorlevel 1 goto :no_python

python -m pip show pyinstaller >nul 2>nul
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller
    if errorlevel 1 goto :install_failed
)

echo Building EXE...
REM 経理モードの services/keiri_* は app.py の route 内で import しているため
REM PyInstaller の静的解析では拾えない。hidden-import で明示しないと exe から欠落する。
python -m PyInstaller ^
  --noconfirm ^
  --onedir ^
  --name KintaiChecker ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --add-data ".env;." ^
  --collect-all pdfplumber ^
  --hidden-import services.keiri_engine ^
  --hidden-import services.keiri_keihi_tenki ^
  --hidden-import services.keiri_api ^
  --hidden-import services.keiri_diff ^
  --hidden-import services.mail_draft ^
  --hidden-import services.mail_ledger_sync ^
  launcher.py

if errorlevel 1 goto :build_failed

echo.
echo ============================================================
echo   Build complete.
echo   Users can now run 起勁Ebat without installing Python.
echo ============================================================
pause
exit /b 0

:no_python
echo [ERROR] Python is required only on the build/admin PC.
echo Please run this bat on a PC where Python is installed.
pause
exit /b 1

:install_failed
echo [ERROR] Failed to install PyInstaller.
echo Check the network connection or install PyInstaller manually:
echo   python -m pip install pyinstaller
pause
exit /b 1

:build_failed
echo [ERROR] Failed to build EXE.
pause
exit /b 1
