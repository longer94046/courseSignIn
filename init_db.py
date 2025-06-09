import sqlite3
import hashlib
import os

def init_db():
    with sqlite3.connect('checkin.db') as conn:
        c = conn.cursor()
        c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            is_admin INTEGER DEFAULT 0
        )
        """)
        c.execute("INSERT OR IGNORE INTO users (username, password, is_admin) VALUES (?, ?, ?)",
                 ('admin', hashlib.sha256('admin123'.encode()).hexdigest(), 1))
        c.execute("""
        CREATE TABLE IF NOT EXISTS org_info (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            org_name TEXT,
            manager TEXT,
            contact TEXT
        )
        """)
        c.execute("INSERT OR IGNORE INTO org_info (id, org_name, manager, contact) VALUES (1, '課堂簽到系統', '', '')")
        c.execute("""
        CREATE TABLE IF NOT EXISTS classes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_id INTEGER,
            week INTEGER,
            date TEXT,
            start_time TEXT,
            end_time TEXT,
            FOREIGN KEY (class_id) REFERENCES classes(id)
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            department TEXT,
            hash TEXT UNIQUE,
            gender TEXT,
            address TEXT,
            phone TEXT,
            id_number TEXT,
            dietary TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS custom_fields (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            field_name TEXT NOT NULL,
            field_type TEXT NOT NULL,
            is_required INTEGER DEFAULT 0,
            display_order INTEGER DEFAULT 0
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS student_custom_values (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER,
            field_id INTEGER,
            field_value TEXT,
            FOREIGN KEY (student_id) REFERENCES students(id),
            FOREIGN KEY (field_id) REFERENCES custom_fields(id)
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS class_students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_id INTEGER,
            student_id INTEGER,
            FOREIGN KEY (class_id) REFERENCES classes(id),
            FOREIGN KEY (student_id) REFERENCES students(id),
            UNIQUE(class_id, student_id)
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS checkins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER,
            student_id INTEGER,
            check_in_time TEXT,
            check_out_time TEXT,
            UNIQUE(session_id, student_id)
        )""")
        c.execute("""
        INSERT OR IGNORE INTO custom_fields (field_name, field_type, is_required, display_order) VALUES 
        ('性別', 'select', 1, 1),
        ('住址', 'text', 0, 2),
        ('連絡電話', 'text', 0, 3),
        ('身分證號', 'text', 0, 4),
        ('餐飲葷素', 'select', 0, 5)
        """)
        c.execute("DELETE FROM custom_fields WHERE field_name='飲食習慣'")
        c.execute("""
        CREATE TABLE IF NOT EXISTS field_options (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            field_id INTEGER,
            option_value TEXT,
            display_order INTEGER DEFAULT 0,
            FOREIGN KEY (field_id) REFERENCES custom_fields(id)
        )""")
        # 性別選項重建
        c.execute("DELETE FROM field_options WHERE field_id=1")
        c.execute("""
        INSERT INTO field_options (field_id, option_value, display_order) VALUES 
        (1, '男', 1),
        (1, '女', 2),
        (1, '其他', 3)
        """)
        c.execute("""
        INSERT OR IGNORE INTO field_options (field_id, option_value, display_order) VALUES 
        (5, '葷食', 1),
        (5, '素食', 2)
        """)
    conn.commit()

if __name__ == "__main__":
    init_db()
    print("資料庫初始化完成！") 