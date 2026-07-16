# 跨錢包批次記帳

## 啟用虛擬環境
- 開啟 PowerShell 進入資料夾  
- 解除當前視窗的腳本執行限制  
`Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process`  
- 啟用虛擬環境  
`.\.venv\Scripts\Activate.ps1`  

## 安裝套件  
- `pip install fastapi`  
- `pip install uvicorn`  
- `pip install pydantic`  

## 啟動 API
- 啟動 API  `python main.py`  
- [測試 API (Swagger UI)](http://127.0.0.1:8000/docs)  
