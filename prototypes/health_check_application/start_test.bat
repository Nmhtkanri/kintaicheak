@echo off
chcp 65001 > nul
cd /d "%~dp0"
title 健康診断申込テスト版
python -X utf8 preview_server.py
if errorlevel 1 (
  echo.
  echo 起動できませんでした。PythonとFlaskが利用できるか確認してください。
  pause
)

