@echo off
title Kintai Checker
setlocal

REM ============================================================
REM  Kintai Checker launcher (local-sync edition)
REM    1. If the server is already running -> just open browser
REM    2. Sync the app from NAS to this PC (only when updated)
REM    3. Start the LOCAL copy (fast startup, per-user isolation)
REM  Old version (direct NAS launch): docs\ (backup .bak file)
REM ============================================================

set "MASTER=%~dp0dist\KintaiChecker"
set "MASTER_ENV=%~dp0.env"
set "LOCALROOT=%LOCALAPPDATA%\KintaiChecker"
set "LOCALAPP=%LOCALROOT%\dist\KintaiChecker"
set "EXE=%LOCALAPP%\KintaiChecker.exe"

echo ============================================================
echo   Kintai Checker - Starting...
echo ============================================================
echo.

REM --- Already running on this PC? Then just open the browser. ---
netstat -ano | findstr /C:":5000 " | findstr "LISTENING" >nul 2>nul
if not errorlevel 1 goto :already_running

if not exist "%MASTER%\KintaiChecker.exe" goto :no_master

REM --- Check whether the local copy is up to date (2 sentinel files) ---
set "NEED_SYNC=0"
if not exist "%EXE%" set "NEED_SYNC=1"
if not exist "%LOCALROOT%\.env" set "NEED_SYNC=1"
if "%NEED_SYNC%"=="1" goto :do_sync

powershell -NoProfile -Command "$pairs=@(,@('%MASTER%\KintaiChecker.exe','%EXE%'))+@(,@('%MASTER_ENV%','%LOCALROOT%\.env')); foreach($p in $pairs){$a=Get-Item -LiteralPath $p[0] -ErrorAction SilentlyContinue; $b=Get-Item -LiteralPath $p[1] -ErrorAction SilentlyContinue; if(-not $a){continue}; if(-not $b -or $a.LastWriteTimeUtc -ne $b.LastWriteTimeUtc -or $a.Length -ne $b.Length){exit 1}}; exit 0"
if errorlevel 1 set "NEED_SYNC=1"

if "%NEED_SYNC%"=="0" goto :run_local

:do_sync
echo Updating local copy from the shared drive...
echo (The first run copies about 200 MB and may take a few minutes.
echo  After that, launch is instant unless the app was updated.)
echo.
robocopy "%MASTER%" "%LOCALAPP%" /MIR /R:2 /W:2 /NP /NFL /NDL /NJH /NJS
if errorlevel 8 goto :sync_failed
if exist "%MASTER_ENV%" copy /Y "%MASTER_ENV%" "%LOCALROOT%\.env" >nul

:run_local
echo Starting Kintai Checker (local copy)...
echo Browser will open automatically in a moment.
echo Close this window to stop the server.
echo.
"%EXE%"
echo.
echo ============================================================
echo   Server stopped.
echo ============================================================
pause
exit /b 0

:already_running
echo Kintai Checker is already running on this PC.
echo Opening the browser...
start "" http://localhost:5000
timeout /t 2 >nul
exit /b 0

:sync_failed
echo.
echo [WARN] Could not update the local copy.
echo        Starting directly from the shared drive instead (slower)...
echo.
"%MASTER%\KintaiChecker.exe"
echo.
echo Server stopped.
pause
exit /b 0

:no_master
echo [ERROR] KintaiChecker.exe was not found on the shared drive:
echo         %MASTER%
echo.
echo Please contact the administrator (build_exe.bat has not been run).
pause
exit /b 1
