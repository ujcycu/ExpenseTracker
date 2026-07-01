import os
import sqlite3
from typing import List, Optional
from fastapi import FastAPI, HTTPException, status
from pydantic import BaseModel, Field
from fastapi.middleware.cors import CORSMiddleware
from datetime import date, timedelta

# 定義資料庫檔案名稱
DB_FILE = "accounting.db"

app = FastAPI(
    title="高效防呆記帳系統 API",
    description="支援類 Excel 批次編輯與三階段金融級對帳機制的後端核心",
    version="1.0.0"
)

# 避免觸發 CORS 跨網域封鎖
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # 允許所有前端網頁造訪
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==========================================
# 1. Pydantic 資料驗證模型 (前端傳入規格與防呆)
# ==========================================

class BulkRecordItem(BaseModel):
    id: Optional[int] = Field(None, description="紀錄 ID。新列填 None，舊資料修改填原本的 ID")
    date: str = Field(default_factory=lambda: date.today().isoformat(), description="交易日期 YYYY-MM-DD")
    type: str = Field(..., pattern="^(income|expense|transfer)$", description="交易類型")
    amount: int = Field(..., gt=0, description="防呆：金額必須大於 0")
    category: str = Field(..., min_length=1, description="分類不能為空字串")
    note: Optional[str] = None
    from_wallet_id: Optional[int] = None
    to_wallet_id: Optional[int] = None
    recon_status: str = Field("unreconciled", pattern="^(unreconciled|matched|billed)$")


# ==========================================
# 2. 資料庫核心工具與初始化
# ==========================================

def get_db_connection():
    """建立資料庫連線，並強制啟用 SQLite 的外鍵約束"""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row  # 讓查詢結果可以用欄位名稱讀取
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

@app.on_event("startup")
def init_db():
    """系統啟動時，自動建立結構健全的資料表"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 建立錢包表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS wallets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        type TEXT NOT NULL,
        balance INTEGER NOT NULL DEFAULT 0,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        CONSTRAINT CHK_WalletType CHECK (type IN ('cash', 'bank', 'credit_card', 'loan'))
    );
    """)
    
    # 建立帳單表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS statements (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        wallet_id INTEGER NOT NULL,
        statement_period TEXT NOT NULL,
        bank_total_amount INTEGER NOT NULL,
        is_paid INTEGER NOT NULL DEFAULT 1,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (wallet_id) REFERENCES wallets(id)
    );
    """)
    
    # 建立紀錄表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,
        type TEXT NOT NULL,
        amount INTEGER NOT NULL,
        category TEXT NOT NULL,
        note TEXT,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        from_wallet_id INTEGER,
        to_wallet_id INTEGER,
        recon_status TEXT NOT NULL DEFAULT 'unreconciled',
        statement_id INTEGER,
        FOREIGN KEY (from_wallet_id) REFERENCES wallets(id),
        FOREIGN KEY (to_wallet_id) REFERENCES wallets(id),
        FOREIGN KEY (statement_id) REFERENCES statements(id),
        CONSTRAINT CHK_AmountPositive CHECK (amount > 0),
        CONSTRAINT CHK_ValidType CHECK (type IN ('income', 'expense', 'transfer')),
        CONSTRAINT CHK_ReconStatus CHECK (recon_status IN ('unreconciled', 'matched', 'billed'))
    );
    """)
    
    # 💡 塞入單人初始測試資料 (Seed Data) 如果表是空的
    cursor.execute("SELECT COUNT(*) FROM wallets;")
    if cursor.fetchone()[0] == 0:
        cursor.executemany("""
            INSERT INTO wallets (name, type, balance) VALUES (?, ?, ?);
        """, [
            ("個人皮夾", "cash", 3000),
            ("富邦銀行", "bank", 50000),
            ("國泰信用卡", "credit_card", 0),  # 信用卡初始未欠款
            ("保險分期", "loan", -40000),      # 貸款初始化為負數負債
            ("車貸", "loan", -24000),      # 貸款初始化為負數負債
            ("iPass Monoey", "cash", 600)
        ])
    
    conn.commit()
    conn.close()


# ==========================================
# 3. API 路由端點 (Endpoints)
# ==========================================

@app.get("/api/wallets", summary="獲取所有錢包餘額與看板資訊")
def get_wallets():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name, type, balance FROM wallets;")
    wallets = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return wallets


@app.get("/api/records", summary="篩選日常編輯主表格明細（排除已歸檔封存資料）")
def get_records(
    days: Optional[int] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    wallet_id: Optional[int] = None
):
# def get_active_records():
    conn = get_db_connection()
    cursor = conn.cursor()

    # 基本條件：排除已月底歸檔的
    query = "SELECT id, date, type, amount, category, note, from_wallet_id, to_wallet_id, recon_status "
    query += " FROM records WHERE recon_status != 'billed'"
    params = []
    # 篩選條件 A：N天之內
    if days is not None:
        target_date = (date.today() - timedelta(days=days)).isoformat()
        query += " AND date >= ?"
        params.append(target_date)
    
    # 篩選條件 B：特定日期區間
    if start_date:
        query += " AND date >= ?"
        params.append(start_date)
    if end_date:
        query += " AND date <= ?"
        params.append(end_date)
        
    # 篩選條件 C：檢視某個特定錢包的交易資料
    if wallet_id is not None:
        query += " AND (from_wallet_id = ? OR to_wallet_id = ?)"
        params.append(wallet_id)
        params.append(wallet_id)
        
    # 排序
    query += " ORDER BY date ASC, id ASC;"
    
    cursor.execute(query, params)
    records = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return records


@app.post("/api/records/bulk-save", summary="【核心防呆】類 Excel 批次原子化儲存端點")
@app.post("/api/records/bulk-save")
def bulk_save_records(items: List[BulkRecordItem]):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        conn.execute("BEGIN TRANSACTION;")

        # 先撈出所有錢包的型態對照表，方便在迴圈中快速判斷
        cursor.execute("SELECT id, type FROM wallets;")
        wallet_types = {row["id"]: row["type"] for row in cursor.fetchall()}
        
        for index, item in enumerate(items):
            # === 【核心防禦：防重複扣款】 ===
            if item.id is not None:
                # 1. 撈出該 ID 存在資料庫的「舊狀態」
                cursor.execute("SELECT type, amount, from_wallet_id, to_wallet_id, recon_status FROM records WHERE id = ?;", (item.id,))
                old = cursor.fetchone()
                
                if old:
                    old_from_type = wallet_types.get(old["from_wallet_id"])
                    old_to_type = wallet_types.get(old["to_wallet_id"])
                    
                    # 💡 【舊 From 錢包回滾】
                    if old["from_wallet_id"]:
                        # 現金/銀行是在日常(不論狀態)就扣款，所以一律加回來
                        if old_from_type in ('cash', 'bank', 'loan'):
                            cursor.execute("UPDATE wallets SET balance = balance + ? WHERE id = ?;", (old["amount"], old["from_wallet_id"]))
                        # 信用卡只有在已對帳(matched)時才算扣款，所以舊狀態是 matched 才需要加回來
                        elif old_from_type == 'credit_card' and old["recon_status"] == 'matched':
                            cursor.execute("UPDATE wallets SET balance = balance + ? WHERE id = ?;", (old["amount"], old["from_wallet_id"]))

                    # 💡 【舊 To 錢包回滾】
                    if old["to_wallet_id"]:
                        if old_to_type in ('cash', 'bank', 'loan'):
                            cursor.execute("UPDATE wallets SET balance = balance - ? WHERE id = ?;", (old["amount"], old["to_wallet_id"]))
                        elif old_to_type == 'credit_card' and old["recon_status"] == 'matched':
                            cursor.execute("UPDATE wallets SET balance = balance - ? WHERE id = ?;", (old["amount"], old["to_wallet_id"]))
            
            # === 2. 驗證新欄位規則 ===
            if item.type == "expense" and item.from_wallet_id is None:
                raise HTTPException(status_code=400, detail=f"第 {index + 1} 行：支出缺少 From 錢包")
            if item.type == "income" and item.to_wallet_id is None:
                raise HTTPException(status_code=400, detail=f"第 {index + 1} 行：收入缺少 To 錢包")
            if item.type == "transfer" and (item.from_wallet_id is None or item.to_wallet_id is None):
                raise HTTPException(status_code=400, detail=f"第 {index + 1} 行：轉帳缺少錢包")

            # === 3. 寫入新資料 / 更新舊資料 ===
            if item.id is None:
                cursor.execute("""
                    INSERT INTO records (date, type, amount, category, note, from_wallet_id, to_wallet_id, recon_status)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?);
                """, (item.date, item.type, item.amount, item.category, item.note, item.from_wallet_id, item.to_wallet_id, item.recon_status))
            else:
                cursor.execute("""
                    UPDATE records SET date=?, type=?, amount=?, category=?, note=?, from_wallet_id=?, to_wallet_id=?, recon_status=? WHERE id=?;
                """, (item.date, item.type, item.amount, item.category, item.note, item.from_wallet_id, item.to_wallet_id, item.recon_status, item.id))
            
            # === 4. 套用新餘額扣減 ===
            new_from_type = wallet_types.get(item.from_wallet_id)
            new_to_type = wallet_types.get(item.to_wallet_id)
            
            # 💡 【新 From 錢包扣款】
            if item.from_wallet_id:
                if new_from_type in ('cash', 'bank', 'loan'):
                    # 現金/銀行/貸款：日常即扣款
                    cursor.execute("UPDATE wallets SET balance = balance - ? WHERE id = ?;", (item.amount, item.from_wallet_id))
                elif new_from_type == 'credit_card' and item.recon_status == 'matched':
                    # 信用卡：只有在勾選「已核對」時才扣款（增加負債）
                    cursor.execute("UPDATE wallets SET balance = balance - ? WHERE id = ?;", (item.amount, item.from_wallet_id))

            # 💡 【新 To 錢包流入】
            if item.to_wallet_id:
                if new_to_type in ('cash', 'bank', 'loan'):
                    # 現金/銀行/貸款：日常即增加入帳
                    cursor.execute("UPDATE wallets SET balance = balance + ? WHERE id = ?;", (item.amount, item.to_wallet_id))
                elif new_to_type == 'credit_card' and item.recon_status == 'matched':
                    # 信用卡：只有在勾選「已核對」時才影響金額（例如退刷或還款）
                    cursor.execute("UPDATE wallets SET balance = balance + ? WHERE id = ?;", (item.amount, item.to_wallet_id))

        conn.commit()
        return {"status": "success", "message": "批次處理成功（已自動計算回滾防重複扣款）"}
    except Exception as e:
        conn.rollback()
        raise e if isinstance(e, HTTPException) else HTTPException(status_code=500, detail=str(e))
    finally:
        conn.close()

if __name__ == "__main__":
    import uvicorn
    # 自動啟動本地伺服器
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)