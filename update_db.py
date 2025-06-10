import sqlite3

def update_db():
    with sqlite3.connect('checkin.db') as conn:
        c = conn.cursor()
        
        # 檢查 type 欄位是否存在
        c.execute("PRAGMA table_info(classes)")
        columns = [column[1] for column in c.fetchall()]
        
        if 'type' not in columns:
            # 備份現有的課程資料
            c.execute("CREATE TABLE classes_backup AS SELECT * FROM classes")
            
            # 刪除舊的課程表
            c.execute("DROP TABLE classes")
            
            # 創建新的課程表
            c.execute("""
            CREATE TABLE classes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                type TEXT NOT NULL DEFAULT 'multi_session'
            )
            """)
            
            # 恢復課程資料，並設置預設類型
            c.execute("""
            INSERT INTO classes (id, name, type)
            SELECT id, name, 'multi_session'
            FROM classes_backup
            """)
            
            # 刪除備份表
            c.execute("DROP TABLE classes_backup")
            
            print("資料庫結構已更新")
        else:
            print("資料庫結構已經是最新的")

if __name__ == "__main__":
    update_db() 