@echo off
title オペレーションハブ
setlocal

REM ============================================================
REM  オペレーションハブ 起動用バッチ（ローカル同期版）
REM    1. すでにサーバーが起動していれば → ブラウザを開くだけ
REM    2. NAS から このPC へアプリを同期（更新があったときだけ）
REM    3. ローカルコピーを起動（起動が速く、ユーザーごとに独立）
REM  旧版（NAS から直接起動）: docs\ の .bak ファイル
REM ============================================================

set "MASTER=%~dp0dist\KintaiChecker"
set "MASTER_ENV=%~dp0.env"
set "LOCALROOT=%LOCALAPPDATA%\KintaiChecker"
set "LOCALAPP=%LOCALROOT%\dist\KintaiChecker"
set "EXE=%LOCALAPP%\KintaiChecker.exe"

echo ============================================================
echo   オペレーションハブ を起動しています...
echo ============================================================
echo.

REM --- このPCで既に起動中なら、ブラウザを開くだけで終了 ---
netstat -ano | findstr /C:":5000 " | findstr "LISTENING" >nul 2>nul
if not errorlevel 1 goto :already_running

if not exist "%MASTER%\KintaiChecker.exe" goto :no_master

REM --- ローカルコピーが最新かどうか判定（目印の2ファイルで比較） ---
set "NEED_SYNC=0"
if not exist "%EXE%" set "NEED_SYNC=1"
if not exist "%LOCALROOT%\.env" set "NEED_SYNC=1"
if "%NEED_SYNC%"=="1" goto :do_sync

powershell -NoProfile -Command "$pairs=@(,@('%MASTER%\KintaiChecker.exe','%EXE%'))+@(,@('%MASTER_ENV%','%LOCALROOT%\.env')); foreach($p in $pairs){$a=Get-Item -LiteralPath $p[0] -ErrorAction SilentlyContinue; $b=Get-Item -LiteralPath $p[1] -ErrorAction SilentlyContinue; if(-not $a){continue}; if(-not $b -or $a.LastWriteTimeUtc -ne $b.LastWriteTimeUtc -or $a.Length -ne $b.Length){exit 1}}; exit 0"
if errorlevel 1 set "NEED_SYNC=1"

if "%NEED_SYNC%"=="0" goto :run_local

:do_sync
echo 共有ドライブから最新版をこのPCにコピーしています...
echo （初回は約200MBのコピーで数分かかります。
echo   2回目以降は、アプリが更新されたときだけコピーが走ります。）
echo.
robocopy "%MASTER%" "%LOCALAPP%" /MIR /R:2 /W:2 /NP /NFL /NDL /NJH /NJS
if errorlevel 8 goto :sync_failed
if exist "%MASTER_ENV%" copy /Y "%MASTER_ENV%" "%LOCALROOT%\.env" >nul

:run_local
echo オペレーションハブ を起動します（ローカルコピー）...
echo まもなくブラウザが自動で開きます。
echo 終了するときは、このウィンドウを閉じてください。
echo.
"%EXE%"
echo.
echo ============================================================
echo   サーバーを停止しました。
echo ============================================================
pause
exit /b 0

:already_running
echo オペレーションハブ は、このPCで既に起動しています。
echo ブラウザを開きます...
start "" http://localhost:5000
timeout /t 2 >nul
exit /b 0

:sync_failed
echo.
echo [警告] ローカルコピーの更新に失敗しました。
echo        共有ドライブから直接起動します（少し遅くなります）...
echo.
"%MASTER%\KintaiChecker.exe"
echo.
echo サーバーを停止しました。
pause
exit /b 0

:no_master
echo [エラー] 共有ドライブに KintaiChecker.exe が見つかりません:
echo         %MASTER%
echo.
echo 管理者に連絡してください（build_exe.bat がまだ実行されていません）。
pause
exit /b 1
