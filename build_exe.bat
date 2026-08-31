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
REM 関数内 import しかされない services/* を hidden-import で明示する（保険）。
REM PyInstaller の解析は関数内 import も追えるので通常は無くても同梱されるが、
REM 呼び出し元をリファクタしたときに黙って exe から落ちるのを防ぐため列挙しておく。
REM ここに足し忘れても即欠落するわけではない一方、足しておけば確実に入る。
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
  --hidden-import services.invoice_mode ^
  --hidden-import services.mail_draft ^
  --hidden-import services.mail_ledger_sync ^
  --hidden-import services.paid_leave_mail ^
  --hidden-import services.higashi_shift_parser ^
  --hidden-import services.kdx_shift_parser ^
  --hidden-import services.ual_shift_parser ^
  --hidden-import services.bbs_shift_parser ^
  --hidden-import services.employee_alias ^
  --hidden-import services.kotsuhi_seisa ^
  --hidden-import services.sap_import_ledger ^
  --hidden-import services.keihi_import_ledger ^
  --hidden-import services.shiwake_teiki_append ^
  --hidden-import services.health_hpm_excel ^
  --hidden-import services.health_hpm_master ^
  --hidden-import services.health_hpm_match ^
  --hidden-import services.health_hpm_csv ^
  --hidden-import services.health_hpm_pdf ^
  --hidden-import services.sharoushi_export ^
  --hidden-import services.shaho_master ^
  --hidden-import services.shaho_engine ^
  --hidden-import services.shaho_check ^
  --hidden-import services.shaho_report ^
  --hidden-import services.shaho_pdf ^
  --hidden-import services.shaho_writer ^
  --hidden-import services.shaho_its ^
  --hidden-import services.invoice_pdf ^
  --hidden-import services.invoice_folders ^
  --hidden-import services.expense_check ^
  --hidden-import services.jinjer_api_client ^
  --hidden-import services.keihi_payroll_import ^
  --hidden-import services.keihi_summary ^
  --collect-submodules services.daicho ^
  --hidden-import win32com.client.dynamic ^
  --hidden-import pythoncom ^
  --hidden-import win32timezone ^
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
