import sqlite3

def check_database():
    # 💡 請將下面的 'accounting.db' 換成你專案中實際的 SQLite 檔案名稱
    db_name = 'accounting.db' 
    
    try:
        conn = sqlite3.connect(db_name)
        # 讓查詢結果可以用欄位名稱讀取
        conn.row_factory = sqlite3.Row 
        cursor = conn.cursor()

        # with conn:
            # cursor.execute("UPDATE wallets SET NAME = '交割A' WHERE id = 3;")
            # cursor.execute("UPDATE wallets SET parent_id = 1 WHERE id = 3;")
            # cursor.execute("UPDATE wallets SET is_archived = 1 WHERE id = 2;")

            # cursor.execute("UPDATE wallets SET balance = 2474 WHERE id = 3;")
            # cursor.execute("UPDATE wallets SET balance = 3000 WHERE id = 3;")
            # print("✅ 已將交割錢包餘額校正")
            # cursor.execute("UPDATE wallets SET balance = 9995 WHERE id = 4;")
            # print("✅ 已將STOCK錢包餘額校正")
            # cursor.execute("DELETE FROM records WHERE ID>3;")
            # print(f"✅ 已成功刪除 records 表中的髒資料（影響行數: {cursor.rowcount}）")
        
        print("=" * 60)
        print(" 🏦 1. 當前 WALLETS 資料表快照 (未封存)")
        print("=" * 60)
        cursor.execute("SELECT id, name, type, balance, parent_id, is_archived FROM wallets WHERE is_archived = 0;")
        wallets = cursor.fetchall()
        
        # 建立一個 ID 到名稱的對照表，方便等等看明細
        wallet_names = {}
        print(f"{'ID':<4} | {'錢包名稱':<12} | {'類型':<10} | {'目前餘額 (balance)':<15} | {'母錢包ID':<6}")
        print("-" * 60)
        for w in wallets:
            wallet_names[w['id']] = w['name']
            print(f"{w['id']:<4} | {w['name']:<12} | {w['type']:<10} | {w['balance']:<18,} | {str(w['parent_id']):<6}")
            
        print("\n" + "=" * 80)
        print(" 📝 2. 最新寫入的 10 筆 RECORDS 明細")
        print("=" * 80)
        cursor.execute("""
            SELECT id, date, type, amount, category, note, from_wallet_id, to_wallet_id, recon_status 
            FROM records 
            ORDER BY id DESC 
            LIMIT 10;
        """)
        records = cursor.fetchall()
        
        print(f"{'ID':<4} | {'日期':<10} | {'類型':<8} | {'金額':<8} | {'分類':<10} | {'從錢包':<10} | {'到錢包':<10} | {'狀態':<10}")
        print("-" * 80)
        for r in records:
            from_w = wallet_names.get(r['from_wallet_id'], f"ID:{r['from_wallet_id']}") if r['from_wallet_id'] else "(空)"
            to_w = wallet_names.get(r['to_wallet_id'], f"ID:{r['to_wallet_id']}") if r['to_wallet_id'] else "(空)"
            print(f"{r['id']:<4} | {r['date']:<10} | {r['type']:<8} | {r['amount']:<8,} | {r['category']:<10} | {from_w:<10} | {to_w:<10} | {r['recon_status']:<10}")
            
        conn.close()
    except Exception as e:
        print(f"❌ 讀取資料庫失敗: {e}")
        print("請檢查 db_name 是否正確，或腳本是否放在正確的專案路徑下。")

if __name__ == "__main__":
    check_database()