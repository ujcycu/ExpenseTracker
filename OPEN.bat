@echo off
title 啟動工作環境
echo 正在幫您開啟工作環境，請稍候...

:: 先開啟Swagger UI網頁
start "" "http://127.0.0.1:8000/docs"

:: 切換到batch檔所在資料夾
cd /d "%~dp0"

:: powershell執行後保持開啟、確保可執行.ps1
:: 指令1. 解除當前視窗的腳本執行限制  
:: 指令2. 啟用虛擬環境
:: 指令3. 啟動 API
powershell -NoExit -ExecutionPolicy Bypass -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process; .\.venv\Scripts\Activate.ps1; python main.py"