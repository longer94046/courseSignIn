import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import sqlite3
import csv
import hashlib
import qrcode
import os
import sys
from datetime import datetime
from PIL import Image, ImageTk
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import TableStyle, Table
from tkcalendar import DateEntry
import platform
import winsound
import json
import pyttsx3

DB_FILE = "checkin.db"
QR_FOLDER = "qrcodes"
QR_SEED = "secure_seed_2024"

if not os.path.exists(QR_FOLDER):
    os.makedirs(QR_FOLDER)

def hash_name(name):
    return hashlib.sha256((name + QR_SEED).encode()).hexdigest()

def init_db():
    with sqlite3.connect(DB_FILE) as conn:
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
        # 預設插入一筆資料（僅一筆）
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
        
        # 新增學員基本資料表
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
        
        # 新增自定義欄位表
        c.execute("""
        CREATE TABLE IF NOT EXISTS custom_fields (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            field_name TEXT NOT NULL,
            field_type TEXT NOT NULL,
            is_required INTEGER DEFAULT 0,
            display_order INTEGER DEFAULT 0
        )""")
        
        # 新增學員自定義欄位值表
        c.execute("""
        CREATE TABLE IF NOT EXISTS student_custom_values (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER,
            field_id INTEGER,
            field_value TEXT,
            FOREIGN KEY (student_id) REFERENCES students(id),
            FOREIGN KEY (field_id) REFERENCES custom_fields(id)
        )""")
        
        # 修改課程學員關聯表
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
        
        # 插入預設欄位
        c.execute("""
        INSERT OR IGNORE INTO custom_fields (field_name, field_type, is_required, display_order) VALUES 
        ('性別', 'select', 1, 1),
        ('住址', 'text', 0, 2),
        ('連絡電話', 'text', 0, 3),
        ('身分證號', 'text', 0, 4),
        ('餐飲葷素', 'select', 0, 5)
        """)
        # 刪除重複的「飲食習慣」欄位
        c.execute("DELETE FROM custom_fields WHERE field_name='飲食習慣'")
        
        # 插入性別選項
        c.execute("""
        CREATE TABLE IF NOT EXISTS field_options (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            field_id INTEGER,
            option_value TEXT,
            display_order INTEGER DEFAULT 0,
            FOREIGN KEY (field_id) REFERENCES custom_fields(id)
        )""")
        
        # 插入預設選項
        c.execute("""
        INSERT OR IGNORE INTO field_options (field_id, option_value, display_order) VALUES 
        (1, '男', 1),
        (1, '女', 2),
        (1, '其他', 3),
        (5, '葷食', 1),
        (5, '素食', 2)
        """)
        # 刪除性別欄位重複選項，只保留「男」「女」「其他」
        c.execute("DELETE FROM field_options WHERE field_id=1 AND option_value NOT IN ('男','女','其他')")
    conn.commit()

class ManageAttendeesDialog(tk.Toplevel):
    def __init__(self, parent, class_id, refresh_callback):
        super().__init__(parent)
        self.class_id = class_id
        self.refresh_callback = refresh_callback
        self.title("管理課程學員")
        self.geometry("600x500")
        self.resizable(False, False)

        # 建立搜尋框架
        search_frame = ttk.Frame(self)
        search_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(search_frame, text="搜尋：").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_students)
        ttk.Entry(search_frame, textvariable=self.search_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # 建立學員列表
        self.tree = ttk.Treeview(self, columns=("name", "dept", "status"), show="headings")
        self.tree.heading("name", text="姓名")
        self.tree.heading("dept", text="部門")
        self.tree.heading("status", text="狀態")
        self.tree.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)

        # 建立按鈕框架
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=5)
        
        ttk.Button(btn_frame, text="新增選取學員", command=self.add_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="移除選取學員", command=self.remove_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="關閉", command=self.destroy).pack(side=tk.RIGHT, padx=5)

        self.load_students()

    def load_students(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            # 取得所有學員
            c.execute("""
                SELECT s.id, s.name, s.department, 
                       CASE WHEN cs.id IS NOT NULL THEN '已加入' ELSE '未加入' END as status
                FROM students s
                LEFT JOIN class_students cs ON cs.student_id = s.id AND cs.class_id = ?
                ORDER BY s.name
            """, (self.class_id,))
            
            for sid, name, dept, status in c.fetchall():
                self.tree.insert("", tk.END, iid=sid, values=(name, dept, status))

    def filter_students(self, *args):
        search_text = self.search_var.get().lower()
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            if search_text in values[0].lower() or search_text in values[1].lower():
                self.tree.item(item, tags=())
            else:
                self.tree.item(item, tags=('hidden',))
        self.tree.tag_configure('hidden', foreground='gray')

    def add_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "請選擇要新增的學員")
            return
        
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            added = 0
            for sid in selected:
                try:
                    c.execute("INSERT INTO class_students (class_id, student_id) VALUES (?, ?)",
                            (self.class_id, sid))
                    added += 1
                except sqlite3.IntegrityError:
                    continue
            conn.commit()
        
        if added > 0:
            messagebox.showinfo("成功", f"已新增 {added} 位學員")
            self.load_students()
            self.refresh_callback()

    def remove_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "請選擇要移除的學員")
            return
        
        if messagebox.askyesno("確認", f"確定要移除選取的 {len(selected)} 位學員嗎？"):
            with sqlite3.connect(DB_FILE) as conn:
                c = conn.cursor()
                for sid in selected:
                    c.execute("DELETE FROM class_students WHERE class_id=? AND student_id=?",
                            (self.class_id, sid))
                conn.commit()
            self.load_students()
            self.refresh_callback()

class LoginWindow(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback = callback
        # 讀取機構資訊
        org_info = self.load_org_info()
        org_name = org_info.get('org_name', '活動(課程)管理系統')
        self.title(f"{org_name}-活動(課程)管理系統登入")
        self.geometry("300x300")  # 增加高度以容納所有元件
        self.resizable(False, False)
        self.grab_set()
        self.focus_force()
        # 置中顯示
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
        # 建立登入表單
        frame = ttk.Frame(self, padding="20")
        frame.pack(expand=True, fill=tk.BOTH)
        
        # 添加組織名稱標籤
        title_label = ttk.Label(frame, text=f"{org_name}", font=("Microsoft JhengHei", 14, "bold"))
        title_label.pack(pady=(0, 5))
        subtitle_label = ttk.Label(frame, text="活動(課程)管理系統登入", font=("Microsoft JhengHei", 12))
        subtitle_label.pack(pady=(0, 15))
        
        # 建立輸入框架
        input_frame = ttk.Frame(frame)
        input_frame.pack(fill=tk.X, pady=5)
        
        # 帳號 row
        ttk.Label(input_frame, text="👤", font=("Microsoft JhengHei", 12)).grid(row=0, column=0, padx=(0, 5), pady=5, sticky="w")
        ttk.Label(input_frame, text="帳號：", font=("Microsoft JhengHei", 10)).grid(row=0, column=1, sticky="e", pady=5)
        self.username_var = tk.StringVar()
        username_entry = ttk.Entry(input_frame, textvariable=self.username_var, font=("Microsoft JhengHei", 10))
        username_entry.grid(row=0, column=2, sticky="we", padx=(5, 0), pady=5)
        
        # 密碼 row
        ttk.Label(input_frame, text="🔒", font=("Microsoft JhengHei", 12)).grid(row=1, column=0, padx=(0, 5), pady=5, sticky="w")
        ttk.Label(input_frame, text="密碼：", font=("Microsoft JhengHei", 10)).grid(row=1, column=1, sticky="e", pady=5)
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(input_frame, textvariable=self.password_var, show="*", font=("Microsoft JhengHei", 10))
        password_entry.grid(row=1, column=2, sticky="we", padx=(5, 0), pady=5)
        
        input_frame.columnconfigure(2, weight=1)
        
        # 建立按鈕框架
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=10)
        login_button = ttk.Button(button_frame, text="登入", command=self.login, width=20)
        login_button.pack()
        
        # 設定按鈕字體
        style = ttk.Style()
        style.configure("TButton", font=("Microsoft JhengHei", 10))
        
        self.bind("<Return>", lambda e: self.login())

    def load_org_info(self):
        import json
        if os.path.exists("org_info.json"):
            try:
                with open("org_info.json", "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                pass
        return {"org_name": "活動(課程)管理系統", "manager": "", "contact": ""}

    def login(self):
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        if not username or not password:
            messagebox.showwarning("警告", "請輸入帳號和密碼")
            return
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("SELECT id, is_admin FROM users WHERE username=? AND password=?",
                     (username, hashlib.sha256(password.encode()).hexdigest()))
            user = c.fetchone()
            if user:
                self.grab_release()
                self.callback(True, user[0], bool(user[1]))
                self.destroy()
            else:
                messagebox.showerror("錯誤", "帳號或密碼錯誤")

class CheckInApp:
    def __init__(self, root):
        self.root = root
        self._after_ids = []
        self.org_info = self.load_org_info()
        self.root.title(f"{self.org_info.get('org_name', '活動(課程)簽到系統')}-活動(課程)簽到系統")
        self.root.geometry("1200x800")

        self.class_id = None
        self.session_id = None
        self.user_id = None
        self.is_admin = False

        self.main_widgets = []  # 新增：記錄所有主介面元件
        self.setup_ui()
        self.load_classes()
        self.tts_engine = pyttsx3.init()
        self.tts_engine.setProperty("rate", 160)

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')

        top_frame = ttk.Frame(self.root)
        top_frame.pack(pady=10, fill=tk.X)
        self.main_widgets.append(top_frame)  # 新增

        # 新增使用者管理按鈕（僅管理員可見）
        self.user_mgmt_btn = ttk.Button(top_frame, text="使用者管理", command=self.open_user_management)
        self.user_mgmt_btn.grid(row=0, column=8, padx=5)
        
        ttk.Label(top_frame, text="選擇活動(課程):").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.class_combo = ttk.Combobox(top_frame, state="readonly", width=40)
        self.class_combo.grid(row=0, column=1, sticky=tk.W)
        self.class_combo.bind("<<ComboboxSelected>>", lambda e: self.select_class())
        ttk.Button(top_frame, text="設定單位資訊", command=self.set_org_info).grid(row=1, column=5, padx=5)
        ttk.Button(top_frame, text="新增活動(課程)", command=self.add_class).grid(row=0, column=2, padx=5)
        ttk.Button(top_frame, text="新增週次", command=self.add_session).grid(row=0, column=3, padx=5)
        ttk.Button(top_frame, text="學員管理", command=self.open_student_management).grid(row=0, column=4, padx=5)
        ttk.Button(top_frame, text="管理課程學員", command=self.open_manage_dialog).grid(row=0, column=5, padx=5)
        ttk.Button(top_frame, text="產生QR Code", command=self.generate_qrcodes).grid(row=0, column=6, padx=5)
        ttk.Button(top_frame, text="刪除選取名單", command=self.delete_selected_attendees).grid(row=0, column=7, padx=5)

        ttk.Label(top_frame, text="選擇週次:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.session_combo = ttk.Combobox(top_frame, state="readonly", width=40)
        self.session_combo.grid(row=1, column=1, sticky=tk.W)
        self.session_combo.bind("<<ComboboxSelected>>", lambda e: self.select_session())

        ttk.Button(top_frame, text="匯入名單", command=self.import_attendees).grid(row=1, column=2, padx=5)
        ttk.Button(top_frame, text="匯出記錄", command=self.export_records).grid(row=1, column=3, padx=5)
        ttk.Button(top_frame, text="手動簽到/簽退", command=self.open_manual_check_window).grid(row=1, column=4, padx=5)

        ttk.Label(top_frame, text="掃描輸入：").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.scan_entry = ttk.Entry(top_frame, width=50)
        self.scan_entry.grid(row=2, column=1, columnspan=3, sticky=tk.W)
        self.scan_entry.bind("<Return>", self.process_scan)

        # 建立表格顯示
        tree_frame = ttk.Frame(self.root)
        tree_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        self.tree = ttk.Treeview(tree_frame, columns=("姓名", "部門", "簽到時間", "簽退時間"), show="headings")
        self.tree.heading("姓名", text="姓名")
        self.tree.heading("部門", text="部門")
        self.tree.heading("簽到時間", text="簽到時間")
        self.tree.heading("簽退時間", text="簽退時間")
        self.tree.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        # 加入橫向scrollbar
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        xscroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=xscroll.set)

        # 狀態區：目前時間 + 統計資訊
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)

        self.time_label = ttk.Label(bottom_frame, text="")
        self.time_label.pack(side=tk.LEFT)

        self.stats_label = ttk.Label(bottom_frame, text="", foreground="blue")
        self.stats_label.pack(side=tk.RIGHT)

        self.update_time()
        self.update_stats()

        # 登出鈕放在 top_frame 最右側
        self.logout_btn = ttk.Button(top_frame, text="登出", command=self.logout_callback)
        self.logout_btn.grid(row=0, column=99, padx=5, sticky="e")
        self.main_widgets.append(self.logout_btn)

        # 在 update_time, update_stats 等 after 任務都要記錄 after_id
        self._after_ids.append(self.root.after(1000, self.update_time))
        self._after_ids.append(self.root.after(1000, self.update_stats))

    def set_logout_callback(self, callback):
        self.logout_callback = callback

    def destroy(self):
        # 取消所有 after 任務
        for after_id in getattr(self, '_after_ids', []):
            try:
                self.root.after_cancel(after_id)
            except Exception:
                pass
        self._after_ids = []
        # 銷毀 root 下所有 widget（除了 LoginWindow）
        for widget in self.root.winfo_children():
            if not isinstance(widget, LoginWindow):
                try:
                    widget.destroy()
                except Exception:
                    pass

    def load_org_info(self):
        import json
        if os.path.exists("org_info.json"):
            try:
                with open("org_info.json", "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                pass
        return {"org_name": "活動(課程)簽到系統", "manager": "", "contact": ""}

    def set_org_info(self):
        top = tk.Toplevel(self.root)
        top.title("設定單位資訊")
        top.geometry("350x250")
        top.resizable(False, False)

        org_var = tk.StringVar(value=self.org_info.get("org_name", ""))
        mgr_var = tk.StringVar(value=self.org_info.get("manager", ""))
        contact_var = tk.StringVar(value=self.org_info.get("contact", ""))

        ttk.Label(top, text="機構名稱：").pack(pady=5)
        ttk.Entry(top, textvariable=org_var).pack(pady=5, fill=tk.X, padx=10)

        ttk.Label(top, text="管理人員：").pack(pady=5)
        ttk.Entry(top, textvariable=mgr_var).pack(pady=5, fill=tk.X, padx=10)

        ttk.Label(top, text="聯絡方式：").pack(pady=5)
        ttk.Entry(top, textvariable=contact_var).pack(pady=5, fill=tk.X, padx=10)

        def save():
            self.org_info = {
                "org_name": org_var.get().strip(),
                "manager": mgr_var.get().strip(),
                "contact": contact_var.get().strip()
            }
            with open("org_info.json", "w", encoding="utf-8") as f:
                json.dump(self.org_info, f, ensure_ascii=False, indent=2)
            self.root.title(self.org_info["org_name"] or "課堂簽到系統")
            top.destroy()
            messagebox.showinfo("完成", "已儲存單位資訊")

        ttk.Button(top, text="儲存", command=save).pack(pady=15)

    def update_time(self):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_label.config(text=f"目前時間: {now}")
        self.root.after(1000, self.update_time)

    def update_stats(self):
        if not self.class_id or not self.session_id:
            self.stats_label.config(text="請先選擇活動(課程)及週次")
        else:
            with sqlite3.connect(DB_FILE) as conn:
                c = conn.cursor()
                c.execute("SELECT COUNT(*) FROM students s INNER JOIN class_students cs ON cs.student_id = s.id WHERE cs.class_id=?", (self.class_id,))
                total = c.fetchone()[0]

                c.execute("SELECT COUNT(DISTINCT student_id) FROM checkins WHERE session_id=? AND check_in_time IS NOT NULL", (self.session_id,))
                checked_in = c.fetchone()[0]

                unchecked_in = total - checked_in

                c.execute("SELECT COUNT(DISTINCT student_id) FROM checkins WHERE session_id=? AND check_out_time IS NOT NULL", (self.session_id,))
                checked_out = c.fetchone()[0]

                unchecked_out = checked_in - checked_out

                stats_text = (f"應到: {total}  |  簽到: {checked_in}  |  未簽到: {unchecked_in}  |  "
                              f"簽退: {checked_out}  |  未簽退: {unchecked_out}")
                self.stats_label.config(text=stats_text)
        self.root.after(1000, self.update_stats)

    def load_classes(self):
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("SELECT id, name FROM classes")
            data = c.fetchall()
        self.class_combo['values'] = [f"{row[1]} ({row[0]})" for row in data]
        self.class_map = {f"{row[1]} ({row[0]})": row[0] for row in data}

    def select_class(self):
        selected = self.class_combo.get()
        if selected in self.class_map:
            self.class_id = self.class_map[selected]
            self.load_sessions()
            self.load_attendees()
            self.update_stats()

    def add_class(self):
        name = simpledialog.askstring("新增活動(課程)", "輸入活動(課程)名稱")
        if name:
            with sqlite3.connect(DB_FILE) as conn:
                c = conn.cursor()
                c.execute("INSERT INTO classes (name) VALUES (?)", (name,))
                conn.commit()
            self.load_classes()

    def load_sessions(self):
        if not self.class_id:
            return
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("SELECT id, week, date, start_time, end_time FROM sessions WHERE class_id=? ORDER BY week", (self.class_id,))
            data = c.fetchall()
        display = [f"第{row[1]}週 {row[2]} {row[3]}-{row[4]}" for row in data]
        self.session_combo['values'] = display
        self.session_map = {display[i]: data[i][0] for i in range(len(data))}

    def select_session(self):
        selected = self.session_combo.get()
        if selected in self.session_map:
            self.session_id = self.session_map[selected]
            self.load_attendees()
            self.update_stats()

    def add_session(self):
        if not self.class_id:
            messagebox.showwarning("警告", "請先選擇活動(課程)")
            return

        top = tk.Toplevel(self.root)
        top.title("新增週次")
        top.geometry("300x300")
        top.resizable(False, False)

        ttk.Label(top, text="週次:").pack(pady=5)
        week_var = tk.IntVar()
        ttk.Entry(top, textvariable=week_var).pack(pady=5)

        ttk.Label(top, text="日期 (YYYY-MM-DD):").pack(pady=5)
        date_entry = DateEntry(top, date_pattern='yyyy-MM-dd')
        date_entry.pack(pady=5)

        ttk.Label(top, text="開始時間 (HH:MM):").pack(pady=5)
        start_var = tk.StringVar()
        ttk.Entry(top, textvariable=start_var).pack(pady=5)

        ttk.Label(top, text="結束時間 (HH:MM):").pack(pady=5)
        end_var = tk.StringVar()
        ttk.Entry(top, textvariable=end_var).pack(pady=5)

        def save():
            try:
                week = week_var.get()
                date = date_entry.get_date().strftime("%Y-%m-%d")
                start = start_var.get().strip()
                end = end_var.get().strip()

                if not start or not end:
                    messagebox.showwarning("警告", "請輸入開始與結束時間")
                    return

                with sqlite3.connect(DB_FILE) as conn:
                    c = conn.cursor()
                    # 檢查是否已有相同週次
                    c.execute("SELECT id FROM sessions WHERE class_id=? AND week=?", (self.class_id, week))
                    if c.fetchone():
                        messagebox.showerror("錯誤", f"第 {week} 週已存在")
                        return

                    c.execute(
                        "INSERT INTO sessions (class_id, week, date, start_time, end_time) VALUES (?, ?, ?, ?, ?)",
                        (self.class_id, week, date, start, end)
                    )
                    conn.commit()

                messagebox.showinfo("成功", f"已新增第 {week} 週")
                top.destroy()
                self.load_sessions()

            except Exception as e:
                messagebox.showerror("錯誤", f"無法儲存週次：{e}")

        # ✅ 修正：確保這個按鈕在 `save` 完整定義後才出現
        ttk.Button(top, text="儲存", command=save).pack(pady=10)

    def load_attendees(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        if not self.class_id or not self.session_id:
            return
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("""
            SELECT s.id, s.name, s.department,
                ci.check_in_time, ci.check_out_time
            FROM students s
            INNER JOIN class_students cs ON cs.student_id = s.id
            LEFT JOIN checkins ci ON ci.student_id = s.id AND ci.session_id = ?
            WHERE cs.class_id = ?
            ORDER BY s.name
            """, (self.session_id, self.class_id))
            rows = c.fetchall()
            for sid, name, dept, cin, cout in rows:
                self.tree.insert("", tk.END, iid=sid, values=(
                    name, dept,
                    cin if cin else "",
                    cout if cout else ""
                ))

    def open_manual_check_window(self):
        if not self.session_id:
            self.show_timed_popup("請先選擇週次", popup_type="warning", duration=4)
            return

        win = tk.Toplevel(self.root)
        win.title("手動輸入備用碼簽到 / 簽退")
        win.geometry("400x160")
        win.resizable(False, False)

        ttk.Label(win, text="請輸入備用碼：").pack(pady=10)
        code_var = tk.StringVar()
        entry = ttk.Entry(win, textvariable=code_var, font=("Helvetica", 16), width=30)
        entry.pack(pady=5)
        entry.focus()

        def check():
            code = code_var.get().strip()
            if not code:
                return

            with sqlite3.connect(DB_FILE) as conn:
                c = conn.cursor()
                c.execute("SELECT id, name FROM students WHERE class_id=?", (self.class_id,))
                students = c.fetchall()

                matched_student = None
                for sid, name in students:
                    if hash_name(name) == code:
                        matched_student = (sid, name)
                        break

                if not matched_student:
                    self.show_timed_popup("查無此學員或備用碼錯誤", popup_type="error", duration=5)
                    return

                sid, name = matched_student
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                c.execute("SELECT check_in_time, check_out_time FROM checkins WHERE session_id=? AND student_id=?",
                          (self.session_id, sid))
                row = c.fetchone()

                if not row:
                    c.execute("INSERT INTO checkins (session_id, student_id, check_in_time) VALUES (?, ?, ?)",
                              (self.session_id, sid, now_str))
                    self.show_timed_popup(f"{name} 簽到成功", popup_type="success", duration=5)
                else:
                    cin, cout = row
                    if cin and not cout:
                        c.execute("UPDATE checkins SET check_out_time=? WHERE session_id=? AND student_id=?",
                                  (now_str, self.session_id, sid))
                        self.show_timed_popup(f"{name} 簽退成功", popup_type="success", duration=5)
                    elif cin and cout:
                        self.show_timed_popup(f"{name} 已簽退，無法重複簽到", popup_type="info", duration=5)
                    else:
                        c.execute("UPDATE checkins SET check_in_time=? WHERE session_id=? AND student_id=?",
                                  (now_str, self.session_id, sid))
                        self.show_timed_popup(f"{name} 簽到成功", popup_type="success", duration=5)

                conn.commit()

            self.load_attendees()
            self.update_stats()
            win.destroy()
            self.scan_entry.focus_set()  # ✅ 執行完自動回到掃描框

        entry.bind("<Return>", lambda e: check())
        ttk.Button(win, text="確認", command=check).pack(pady=10)

    def process_scan(self, event):
        code = self.scan_entry.get().strip()
        self.scan_entry.delete(0, tk.END)

        if not self.session_id:
            self.show_timed_popup("請先選擇週次", popup_type="warning", duration=4)
            return
        if not code:
            return

        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("""
                SELECT s.id, s.name 
                FROM students s
                INNER JOIN class_students cs ON cs.student_id = s.id
                WHERE cs.class_id = ? AND s.hash = ?
            """, (self.class_id, code))
            student = c.fetchone()

            if not student:
                self.show_timed_popup("查無此學員或QR碼錯誤", popup_type="error", duration=5)
                return

            sid, name = student
            c.execute("SELECT check_in_time, check_out_time FROM checkins WHERE session_id=? AND student_id=?",
                      (self.session_id, sid))
            row = c.fetchone()
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if not row:
                c.execute("INSERT INTO checkins (session_id, student_id, check_in_time) VALUES (?, ?, ?)",
                          (self.session_id, sid, now_str))
                self.show_timed_popup(f"{name} 簽到成功", popup_type="success", duration=5)
            else:
                cin, cout = row
                if cin and not cout:
                    c.execute("UPDATE checkins SET check_out_time=? WHERE session_id=? AND student_id=?",
                              (now_str, self.session_id, sid))
                    self.show_timed_popup(f"{name} 簽退成功", popup_type="success", duration=5)
                elif cin and cout:
                    self.show_timed_popup(f"{name} 已簽退，無法重複簽到", popup_type="info", duration=5)
                else:
                    c.execute("UPDATE checkins SET check_in_time=? WHERE session_id=? AND student_id=?",
                              (now_str, self.session_id, sid))
                    self.show_timed_popup(f"{name} 簽到成功", popup_type="success", duration=5)

            conn.commit()

        self.load_attendees()
        self.update_stats()

    def show_timed_popup(self, message, popup_type="info", duration=5):
        popup = tk.Toplevel(self.root)
        popup.withdraw()
        popup.title({
                        "info": "訊息",
                        "success": "成功",
                        "warning": "警告",
                        "error": "錯誤"
                    }.get(popup_type, "訊息"))
        popup.geometry("300x120")
        popup.resizable(False, False)
        popup.overrideredirect(True)

        frame = ttk.Frame(popup, padding=10, relief="ridge")
        frame.pack(expand=True, fill=tk.BOTH)

        color_map = {
            "info": "black",
            "success": "green",
            "warning": "orange",
            "error": "red"
        }
        text_color = color_map.get(popup_type, "black")

        ttk.Label(frame, text=message, font=("Helvetica", 12), foreground=text_color).pack(pady=(10, 5))
        countdown_var = tk.StringVar(value=f"{duration} 秒後自動關閉")
        ttk.Label(frame, textvariable=countdown_var, foreground="gray").pack()

        popup.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = popup.winfo_reqwidth()
        height = popup.winfo_reqheight()
        x = screen_width - width - 20
        y = screen_height - height - 60
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.deiconify()

        def update_countdown(sec):
            if sec > 0:
                countdown_var.set(f"{sec} 秒後自動關閉")
                self.root.after(1000, lambda: update_countdown(sec - 1))
            else:
                popup.destroy()
                self.scan_entry.focus_set()
        try:
            self.tts_engine.say(message)
            self.tts_engine.runAndWait()
        except Exception as e:
            print(f"TTS 撥放失敗：{e}")
        update_countdown(duration)

    def import_attendees(self):
        if not self.class_id:
            messagebox.showwarning("警告", "請先選擇活動(課程)")
            return
        file_path = filedialog.askopenfilename(filetypes=[("CSV檔案", "*.csv")])
        if not file_path:
            return
        try:
            with open(file_path, newline='', encoding="utf-8") as csvfile:
                reader = csv.DictReader(csvfile)
                rows = list(reader)
        except UnicodeDecodeError:
            with open(file_path, newline='', encoding="cp950") as csvfile:
                reader = csv.DictReader(csvfile)
                rows = list(reader)
        added = 0
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            for row in rows:
                name = row.get("姓名", "").strip()
                dept = row.get("部門", "").strip()
                if not name:
                    continue
                
                # 檢查學員是否已存在
                c.execute("SELECT id FROM students WHERE name=?", (name,))
                student = c.fetchone()
                
                if student:
                    student_id = student[0]
                else:
                    # 新增學員
                    h = hash_name(name)
                    c.execute("INSERT INTO students (name, department, hash) VALUES (?, ?, ?)",
                             (name, dept, h))
                    student_id = c.lastrowid
                
                # 檢查是否已經加入此課程
                c.execute("SELECT id FROM class_students WHERE class_id=? AND student_id=?",
                         (self.class_id, student_id))
                if not c.fetchone():
                    # 建立課程與學員的關聯
                    c.execute("INSERT INTO class_students (class_id, student_id) VALUES (?, ?)",
                             (self.class_id, student_id))
                    added += 1
            
            conn.commit()
        messagebox.showinfo("匯入完成", f"成功匯入 {added} 位學員")
        self.load_attendees()
        self.update_stats()

    def export_records(self):
        if not self.class_id or not self.session_id:
            messagebox.showwarning("警告", "請先選擇活動(課程)及週次")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                 filetypes=[("PDF檔案", "*.pdf")])
        if not file_path:
            return

        # 註冊微軟正黑體（Windows 系統）
        if platform.system() == "Windows":
            font_path = os.path.join(os.environ['WINDIR'], 'Fonts', 'msjh.ttc')
            if not os.path.exists(font_path):
                messagebox.showerror("錯誤", "找不到微軟正黑體字型（msjh.ttc）")
                return
            pdfmetrics.registerFont(TTFont('MicrosoftJhengHei', font_path))
            font_name = 'MicrosoftJhengHei'
        else:
            messagebox.showerror("錯誤", "目前僅支援 Windows 系統的中文字型顯示")
            return

        # 從資料庫讀取資料與課程/堂次資訊
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()

            # 活動(課程)名稱
            c.execute("SELECT name FROM classes WHERE id=?", (self.class_id,))
            row = c.fetchone()
            class_name = row[0] if row else "(未知活動(課程))"

            # 堂次資訊
            c.execute("SELECT week, date, start_time, end_time FROM sessions WHERE id=?", (self.session_id,))
            session_data = c.fetchone()
            if session_data:
                week, date, start, end = session_data
                session_info = f"第{week}週  {date}  {start}~{end}"
            else:
                session_info = "(未知堂次)"

            # 出席記錄
            c.execute("""
                SELECT a.name, a.department, ci.check_in_time, ci.check_out_time
                FROM students a
                INNER JOIN class_students cs ON cs.student_id = a.id
                LEFT JOIN checkins ci ON ci.student_id = a.id AND ci.session_id = ?
                WHERE cs.class_id = ?
                ORDER BY a.name
            """, (self.session_id, self.class_id))
            records = c.fetchall()

            # 統計資訊
            total = len(records)
            checked_in = sum(1 for r in records if r[2])
            checked_out = sum(1 for r in records if r[3])
            unchecked_in = total - checked_in
            unchecked_out = checked_in - checked_out
            stats_text = (f"應到: {total}  |  簽到: {checked_in}  |  未簽到: {unchecked_in}  |  "
                          f"簽退: {checked_out}  |  未簽退: {unchecked_out}")

        try:
            pdf = Canvas(file_path, pagesize=A4)
            width, height = A4

            # 畫列印時間
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            pdf.setFont(font_name, 10)
            pdf.drawRightString(width - 50, height - 30, f"列印日期：{now_str}")

            # 標題與課程資訊
            y = height - 60
            pdf.setFont(font_name, 14)
            pdf.drawString(50, y, "簽到記錄報表")
            y -= 20
            pdf.setFont(font_name, 12)
            # 單位資訊
            y -= 20
            pdf.setFont(font_name, 12)
            pdf.drawString(50, y, f"單位名稱：{self.org_info.get('org_name', '')}")
            y -= 20
            pdf.drawString(50, y, f"管理人員：{self.org_info.get('manager', '')}")
            y -= 20
            pdf.drawString(50, y, f"聯絡方式：{self.org_info.get('contact', '')}")
            y -= 20
            pdf.drawString(50, y, f"課程名稱：{class_name}")
            y -= 20
            pdf.drawString(50, y, f"堂次資訊：{session_info}")

            y -= 20
            pdf.setFont(font_name, 11)
            pdf.drawString(50, y, f"統計資訊：{stats_text}")

            # 建立表格資料
            table_data = [["姓名", "部門", "簽到時間", "簽退時間"]]
            table_data += records

            # 建立表格
            table = Table(table_data, colWidths=[100, 100, 150, 150])
            table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), font_name),
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ]))

            # 繪製表格
            table_width, table_height = table.wrap(0, 0)
            table.drawOn(pdf, 50, y - 40 - table_height)

            pdf.save()
            messagebox.showinfo("匯出完成", "PDF 匯出成功")
        except Exception as e:
            messagebox.showerror("錯誤", f"匯出 PDF 失敗：{e}")

    def generate_qrcodes(self):
        if not self.class_id:
            messagebox.showwarning("警告", "請先選擇課堂")
            return

        folder_path = filedialog.askdirectory(title="選擇 QR Code 儲存資料夾")
        if not folder_path:
            return

        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("""
                SELECT s.name, s.hash
                FROM students s
                INNER JOIN class_students cs ON cs.student_id = s.id
                WHERE cs.class_id = ?
            """, (self.class_id,))
            students = c.fetchall()
            
            if not students:
                messagebox.showwarning("警告", "此課堂尚無學員")
                return

            for name, h in students:
                backup_code = h[:10]

                qr = qrcode.QRCode(
                    version=1,
                    error_correction=qrcode.constants.ERROR_CORRECT_H,
                    box_size=10,
                    border=4,
                )
                qr.add_data(h)
                qr.make(fit=True)
                qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")

                from PIL import ImageDraw, ImageFont
                width, height = qr_img.size
                new_height = height + 80
                final_img = Image.new("RGB", (width, new_height), "white")
                final_img.paste(qr_img, (0, 0))

                draw = ImageDraw.Draw(final_img)
                try:
                    font = ImageFont.truetype("msjh.ttf", 20)
                except:
                    font = ImageFont.load_default()
                text = f"{name}｜備用碼：{backup_code}"

                bbox = draw.textbbox((0, 0), text, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]

                draw.text(((width - text_width) / 2, height + 10), text, fill="black", font=font)
                final_img.save(os.path.join(folder_path, f"{name}.png"))

        messagebox.showinfo("完成", f"QR Code（含備用碼）已儲存至 {folder_path}")

    def delete_selected_attendees(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "請先選擇欲刪除的學員")
            return
        confirm = messagebox.askyesno("確認刪除", f"確定刪除選取的 {len(selected)} 位學員嗎？")
        if not confirm:
            return
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            for sid in selected:
                c.execute("DELETE FROM class_students WHERE class_id=? AND student_id=?", (self.class_id, sid))
            conn.commit()
        self.load_attendees()
        self.update_stats()

    def open_manage_dialog(self):
        if not self.class_id:
            messagebox.showwarning("警告", "請先選擇活動(課程)")
            return
        ManageAttendeesDialog(self.root, self.class_id, self.update_stats)

    def open_user_management(self):
        if not self.is_admin:
            messagebox.showwarning("警告", "只有管理員可以使用此功能")
            return
            
        win = tk.Toplevel(self.root)
        win.title("使用者管理")
        win.geometry("500x400")
        win.resizable(False, False)

        # 建立使用者列表
        tree = ttk.Treeview(win, columns=("username", "is_admin"), show="headings")
        tree.heading("username", text="帳號")
        tree.heading("is_admin", text="管理員權限")
        tree.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        def load_users():
            for item in tree.get_children():
                tree.delete(item)
            with sqlite3.connect(DB_FILE) as conn:
                c = conn.cursor()
                c.execute("SELECT id, username, is_admin FROM users")
                for uid, username, is_admin in c.fetchall():
                    tree.insert("", tk.END, iid=uid, values=(username, "是" if is_admin else "否"))

        def add_user():
            dialog = tk.Toplevel(win)
            dialog.title("新增使用者")
            dialog.geometry("300x200")
            dialog.resizable(False, False)

            ttk.Label(dialog, text="帳號：").pack(pady=5)
            username_var = tk.StringVar()
            ttk.Entry(dialog, textvariable=username_var).pack(pady=5, fill=tk.X, padx=10)

            ttk.Label(dialog, text="密碼：").pack(pady=5)
            password_var = tk.StringVar()
            ttk.Entry(dialog, textvariable=password_var, show="*").pack(pady=5, fill=tk.X, padx=10)

            is_admin_var = tk.BooleanVar()
            ttk.Checkbutton(dialog, text="管理員權限", variable=is_admin_var).pack(pady=5)

            def save():
                username = username_var.get().strip()
                password = password_var.get().strip()
                if not username or not password:
                    messagebox.showwarning("警告", "請輸入帳號和密碼")
                    return
                try:
                    with sqlite3.connect(DB_FILE) as conn:
                        c = conn.cursor()
                        c.execute("INSERT INTO users (username, password, is_admin) VALUES (?, ?, ?)",
                                (username, hashlib.sha256(password.encode()).hexdigest(), int(is_admin_var.get())))
                        conn.commit()
                    dialog.destroy()
                    load_users()
                except sqlite3.IntegrityError:
                    messagebox.showerror("錯誤", "帳號已存在")

            ttk.Button(dialog, text="儲存", command=save).pack(pady=10)

        def delete_user():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("警告", "請選擇要刪除的使用者")
                return
            if messagebox.askyesno("確認", "確定要刪除選取的使用者嗎？"):
                with sqlite3.connect(DB_FILE) as conn:
                    c = conn.cursor()
                    for uid in selected:
                        c.execute("DELETE FROM users WHERE id=?", (uid,))
                    conn.commit()
                load_users()

        def reset_password():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("警告", "請選擇要重設密碼的使用者")
                return
            if len(selected) > 1:
                messagebox.showwarning("警告", "一次只能重設一個使用者的密碼")
                return
                
            new_password = simpledialog.askstring("重設密碼", "請輸入新密碼：", show="*")
            if new_password:
                with sqlite3.connect(DB_FILE) as conn:
                    c = conn.cursor()
                    c.execute("UPDATE users SET password=? WHERE id=?",
                            (hashlib.sha256(new_password.encode()).hexdigest(), selected[0]))
                    conn.commit()
                messagebox.showinfo("成功", "密碼已重設")

        # 按鈕框架
        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="新增使用者", command=add_user).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="刪除使用者", command=delete_user).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="重設密碼", command=reset_password).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="關閉", command=win.destroy).pack(side=tk.RIGHT, padx=5)

        load_users()

    def open_student_management(self):
        StudentManagementDialog(self.root)

    def logout_callback(self):
        python = sys.executable
        os.execl(python, python, *sys.argv)

class StudentManagementDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("學員管理")
        self.geometry("800x600")
        self.resizable(False, False)

        # 建立搜尋框架
        search_frame = ttk.Frame(self)
        search_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(search_frame, text="搜尋：").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_students)
        ttk.Entry(search_frame, textvariable=self.search_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # 建立學員列表（含橫向scrollbar）
        tree_frame = ttk.Frame(self)
        tree_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)
        self.tree = ttk.Treeview(tree_frame, columns=("name", "dept", "gender", "phone", "dietary"), show="headings")
        self.tree.heading("name", text="姓名")
        self.tree.heading("dept", text="部門")
        self.tree.heading("gender", text="性別")
        self.tree.heading("phone", text="連絡電話")
        self.tree.heading("dietary", text="餐飲葷素")
        self.tree.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        xscroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=xscroll.set)

        # 建立按鈕框架
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=5)
        
        ttk.Button(btn_frame, text="新增學員", command=self.add_student).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="編輯學員", command=self.edit_student).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="刪除學員", command=self.delete_student).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="匯入學員", command=self.import_students).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="匯出學員", command=self.export_students).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="管理欄位", command=self.manage_fields).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="關閉", command=self.destroy).pack(side=tk.RIGHT, padx=5)

        self.load_students()

    def load_students(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("""
                SELECT id, name, department, gender, phone, dietary 
                FROM students 
                ORDER BY name
            """)
            for sid, name, dept, gender, phone, dietary in c.fetchall():
                self.tree.insert("", tk.END, iid=sid, values=(name, dept, gender, phone, dietary))

    def filter_students(self, *args):
        search_text = self.search_var.get().lower()
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            if search_text in values[0].lower() or search_text in values[1].lower():
                self.tree.item(item, tags=())
            else:
                self.tree.item(item, tags=('hidden',))
        self.tree.tag_configure('hidden', foreground='gray')

    def add_student(self):
        dialog = tk.Toplevel(self)
        dialog.title("新增學員")
        dialog.geometry("400x600")
        dialog.resizable(False, False)

        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 基本資料
        basic_frame = ttk.LabelFrame(main_frame, text="基本資料", padding=10)
        basic_frame.pack(fill=tk.X, pady=5)
        row = 0
        ttk.Label(basic_frame, text="姓名：").grid(row=row, column=0, sticky="e", pady=5)
        name_var = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=name_var).grid(row=row, column=1, sticky="we", pady=5)
        row += 1
        ttk.Label(basic_frame, text="部門：").grid(row=row, column=0, sticky="e", pady=5)
        dept_var = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=dept_var).grid(row=row, column=1, sticky="we", pady=5)
        row += 1
        # 性別
        gender_var = tk.StringVar()
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("SELECT option_value FROM field_options WHERE field_id=1 ORDER BY display_order")
            gender_options = [row_[0] for row_ in c.fetchall()]
        ttk.Label(basic_frame, text="性別：").grid(row=row, column=0, sticky="e", pady=5)
        ttk.Combobox(basic_frame, textvariable=gender_var, values=gender_options, state="readonly").grid(row=row, column=1, sticky="we", pady=5)
        row += 1
        # 連絡電話
        ttk.Label(basic_frame, text="連絡電話：").grid(row=row, column=0, sticky="e", pady=5)
        phone_var = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=phone_var).grid(row=row, column=1, sticky="we", pady=5)
        row += 1
        # 餐飲葷素
        dietary_var = tk.StringVar()
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("SELECT option_value FROM field_options WHERE field_id=5 ORDER BY display_order")
            dietary_options = [row_[0] for row_ in c.fetchall()]
        ttk.Label(basic_frame, text="餐飲葷素：").grid(row=row, column=0, sticky="e", pady=5)
        ttk.Combobox(basic_frame, textvariable=dietary_var, values=dietary_options, state="readonly").grid(row=row, column=1, sticky="we", pady=5)
        row += 1
        basic_frame.columnconfigure(1, weight=1)

        # 其他資料
        custom_frame = ttk.LabelFrame(main_frame, text="其他資料", padding=10)
        custom_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        custom_vars = {}
        basic_names = {"姓名", "部門", "性別", "連絡電話", "餐飲葷素"}
        shown_names = set()
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("""
                SELECT id, field_name, field_type, is_required 
                FROM custom_fields 
                WHERE id NOT IN (1,3,5) 
                ORDER BY display_order
            """)
            row = 0
            for field_id, field_name, field_type, is_required in c.fetchall():
                if field_name in basic_names or field_name in shown_names:
                    continue
                shown_names.add(field_name)
                ttk.Label(custom_frame, text=f"{field_name}{'*' if is_required else ''}：").grid(row=row, column=0, sticky="e", pady=5)
                if field_type == 'select':
                    var = tk.StringVar()
                    c.execute("SELECT option_value FROM field_options WHERE field_id=? ORDER BY display_order", (field_id,))
                    options = [row_[0] for row_ in c.fetchall()]
                    ttk.Combobox(custom_frame, textvariable=var, values=options, state="readonly").grid(row=row, column=1, sticky="we", pady=5)
                else:
                    var = tk.StringVar()
                    ttk.Entry(custom_frame, textvariable=var).grid(row=row, column=1, sticky="we", pady=5)
                custom_vars[field_id] = var
                row += 1
        custom_frame.columnconfigure(1, weight=1)

        # 按鈕
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(side=tk.BOTTOM, pady=10)
        def save():
            name = name_var.get().strip()
            dept = dept_var.get().strip()
            if not name:
                messagebox.showwarning("警告", "姓名不能為空")
                return
            try:
                with sqlite3.connect(DB_FILE) as conn:
                    c = conn.cursor()
                    h = hash_name(name)
                    c.execute("""
                        INSERT INTO students (name, department, hash, gender, phone, dietary) 
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (name, dept, h, gender_var.get(), phone_var.get(), dietary_var.get()))
                    student_id = c.lastrowid
                    for field_id, var in custom_vars.items():
                        value = var.get().strip()
                        if value:
                            c.execute("""
                                INSERT INTO student_custom_values (student_id, field_id, field_value)
                                VALUES (?, ?, ?)
                            """, (student_id, field_id, value))
                    conn.commit()
                dialog.destroy()
                self.load_students()
            except sqlite3.IntegrityError:
                messagebox.showerror("錯誤", "該學員已存在")
        ttk.Button(btn_frame, text="儲存", command=save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        dialog.focus_set()

    def edit_student(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "請選擇要編輯的學員")
            return
        if len(selected) > 1:
            messagebox.showwarning("警告", "一次只能編輯一個學員")
            return

        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("""
                SELECT name, department, gender, phone, dietary 
                FROM students 
                WHERE id=?
            """, (selected[0],))
            name, dept, gender, phone, dietary = c.fetchone()

            # 取得自定義欄位值
            c.execute("""
                SELECT field_id, field_value 
                FROM student_custom_values 
                WHERE student_id=?
            """, (selected[0],))
            custom_values = dict(c.fetchall())

        dialog = tk.Toplevel(self)
        dialog.title("編輯學員")
        dialog.geometry("400x600")
        dialog.resizable(False, False)

        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 基本資料
        basic_frame = ttk.LabelFrame(main_frame, text="基本資料", padding=10)
        basic_frame.pack(fill=tk.X, pady=5)
        row = 0
        ttk.Label(basic_frame, text="姓名：").grid(row=row, column=0, sticky="e", pady=5)
        name_var = tk.StringVar(value=name)
        ttk.Entry(basic_frame, textvariable=name_var).grid(row=row, column=1, sticky="we", pady=5)
        row += 1
        ttk.Label(basic_frame, text="部門：").grid(row=row, column=0, sticky="e", pady=5)
        dept_var = tk.StringVar(value=dept)
        ttk.Entry(basic_frame, textvariable=dept_var).grid(row=row, column=1, sticky="we", pady=5)
        row += 1
        # 性別
        gender_var = tk.StringVar(value=gender)
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("SELECT option_value FROM field_options WHERE field_id=1 ORDER BY display_order")
            gender_options = [row_[0] for row_ in c.fetchall()]
        ttk.Label(basic_frame, text="性別：").grid(row=row, column=0, sticky="e", pady=5)
        ttk.Combobox(basic_frame, textvariable=gender_var, values=gender_options, state="readonly").grid(row=row, column=1, sticky="we", pady=5)
        row += 1
        # 連絡電話
        ttk.Label(basic_frame, text="連絡電話：").grid(row=row, column=0, sticky="e", pady=5)
        phone_var = tk.StringVar(value=phone)
        ttk.Entry(basic_frame, textvariable=phone_var).grid(row=row, column=1, sticky="we", pady=5)
        row += 1
        # 餐飲葷素
        dietary_var = tk.StringVar(value=dietary)
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("SELECT option_value FROM field_options WHERE field_id=5 ORDER BY display_order")
            dietary_options = [row_[0] for row_ in c.fetchall()]
        ttk.Label(basic_frame, text="餐飲葷素：").grid(row=row, column=0, sticky="e", pady=5)
        ttk.Combobox(basic_frame, textvariable=dietary_var, values=dietary_options, state="readonly").grid(row=row, column=1, sticky="we", pady=5)
        row += 1
        basic_frame.columnconfigure(1, weight=1)

        # 其他資料
        custom_frame = ttk.LabelFrame(main_frame, text="其他資料", padding=10)
        custom_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        custom_vars = {}
        basic_names = {"姓名", "部門", "性別", "連絡電話", "餐飲葷素"}
        shown_names = set()
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("""
                SELECT id, field_name, field_type, is_required 
                FROM custom_fields 
                WHERE id NOT IN (1,3,5) 
                ORDER BY display_order
            """)
            row = 0
            for field_id, field_name, field_type, is_required in c.fetchall():
                if field_name in basic_names or field_name in shown_names:
                    continue
                shown_names.add(field_name)
                ttk.Label(custom_frame, text=f"{field_name}{'*' if is_required else ''}：").grid(row=row, column=0, sticky="e", pady=5)
                if field_type == 'select':
                    var = tk.StringVar(value=custom_values.get(field_id, ""))
                    c.execute("SELECT option_value FROM field_options WHERE field_id=? ORDER BY display_order", (field_id,))
                    options = [row_[0] for row_ in c.fetchall()]
                    ttk.Combobox(custom_frame, textvariable=var, values=options, state="readonly").grid(row=row, column=1, sticky="we", pady=5)
                else:
                    var = tk.StringVar(value=custom_values.get(field_id, ""))
                    ttk.Entry(custom_frame, textvariable=var).grid(row=row, column=1, sticky="we", pady=5)
                custom_vars[field_id] = var
                row += 1
        custom_frame.columnconfigure(1, weight=1)

        # 按鈕
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(side=tk.BOTTOM, pady=10)
        def save():
            new_name = name_var.get().strip()
            new_dept = dept_var.get().strip()
            if not new_name:
                messagebox.showwarning("警告", "姓名不能為空")
                return
            try:
                with sqlite3.connect(DB_FILE) as conn:
                    c = conn.cursor()
                    h = hash_name(new_name)
                    c.execute("""
                        UPDATE students 
                        SET name=?, department=?, hash=?, gender=?, phone=?, dietary=?
                        WHERE id=?
                    """, (new_name, new_dept, h, gender_var.get(), phone_var.get(), dietary_var.get(), selected[0]))
                    c.execute("DELETE FROM student_custom_values WHERE student_id=?", (selected[0],))
                    for field_id, var in custom_vars.items():
                        value = var.get().strip()
                        if value:
                            c.execute("""
                                INSERT INTO student_custom_values (student_id, field_id, field_value)
                                VALUES (?, ?, ?)
                            """, (selected[0], field_id, value))
                    conn.commit()
                dialog.destroy()
                self.load_students()
            except sqlite3.IntegrityError:
                messagebox.showerror("錯誤", "該學員名稱已存在")
        ttk.Button(btn_frame, text="儲存", command=save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        dialog.focus_set()

    def manage_fields(self):
        dialog = tk.Toplevel(self)
        dialog.title("管理欄位")
        dialog.geometry("500x400")
        dialog.resizable(False, False)

        # 建立欄位列表
        tree = ttk.Treeview(dialog, columns=("name", "type", "required"), show="headings")
        tree.heading("name", text="欄位名稱")
        tree.heading("type", text="欄位類型")
        tree.heading("required", text="必填")
        tree.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)

        def load_fields():
            for row in tree.get_children():
                tree.delete(row)
            with sqlite3.connect(DB_FILE) as conn:
                c = conn.cursor()
                c.execute("""
                    SELECT id, field_name, field_type, is_required 
                    FROM custom_fields 
                    ORDER BY display_order
                """)
                for fid, name, type_, required in c.fetchall():
                    tree.insert("", tk.END, iid=fid, values=(name, type_, "是" if required else "否"))

        def add_field():
            field_dialog = tk.Toplevel(dialog)
            field_dialog.title("新增欄位")
            field_dialog.geometry("300x250")
            field_dialog.resizable(False, False)

            ttk.Label(field_dialog, text="欄位名稱：").pack(pady=5)
            name_var = tk.StringVar()
            ttk.Entry(field_dialog, textvariable=name_var).pack(pady=5, fill=tk.X, padx=10)

            ttk.Label(field_dialog, text="欄位類型：").pack(pady=5)
            type_var = tk.StringVar(value="text")
            ttk.Combobox(field_dialog, textvariable=type_var, values=["text", "select"], state="readonly").pack(pady=5, fill=tk.X, padx=10)

            required_var = tk.BooleanVar()
            ttk.Checkbutton(field_dialog, text="必填欄位", variable=required_var).pack(pady=5)

            def save_field():
                name = name_var.get().strip()
                if not name:
                    messagebox.showwarning("警告", "欄位名稱不能為空")
                    return

                with sqlite3.connect(DB_FILE) as conn:
                    c = conn.cursor()
                    c.execute("""
                        INSERT INTO custom_fields (field_name, field_type, is_required, display_order)
                        VALUES (?, ?, ?, (SELECT COALESCE(MAX(display_order), 0) + 1 FROM custom_fields))
                    """, (name, type_var.get(), int(required_var.get())))
                    field_id = c.lastrowid

                    if type_var.get() == "select":
                        options_dialog = tk.Toplevel(field_dialog)
                        options_dialog.title("設定選項")
                        options_dialog.geometry("300x400")
                        options_dialog.resizable(False, False)

                        options_list = []
                        options_frame = ttk.Frame(options_dialog)
                        options_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

                        def add_option():
                            option = simpledialog.askstring("新增選項", "請輸入選項值：")
                            if option:
                                options_list.append(option)
                                update_options_list()

                        def update_options_list():
                            for widget in options_frame.winfo_children():
                                widget.destroy()
                            for i, option in enumerate(options_list):
                                ttk.Label(options_frame, text=f"{i+1}. {option}").pack(pady=2)

                        ttk.Button(options_frame, text="新增選項", command=add_option).pack(pady=5)

                        def save_options():
                            with sqlite3.connect(DB_FILE) as conn:
                                c = conn.cursor()
                                for i, option in enumerate(options_list):
                                    c.execute("""
                                        INSERT INTO field_options (field_id, option_value, display_order)
                                        VALUES (?, ?, ?)
                                    """, (field_id, option, i+1))
                                conn.commit()
                            options_dialog.destroy()
                            field_dialog.destroy()
                            load_fields()

                        ttk.Button(options_dialog, text="儲存", command=save_options).pack(pady=10)
                    else:
                        conn.commit()
                        field_dialog.destroy()
                        load_fields()

            ttk.Button(field_dialog, text="儲存", command=save_field).pack(pady=10)

        def delete_field():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("警告", "請選擇要刪除的欄位")
                return
            if messagebox.askyesno("確認", "確定要刪除選取的欄位嗎？\n注意：刪除欄位將同時刪除所有相關的資料。"):
                with sqlite3.connect(DB_FILE) as conn:
                    c = conn.cursor()
                    for fid in selected:
                        c.execute("DELETE FROM field_options WHERE field_id=?", (fid,))
                        c.execute("DELETE FROM student_custom_values WHERE field_id=?", (fid,))
                        c.execute("DELETE FROM custom_fields WHERE id=?", (fid,))
                    conn.commit()
                load_fields()

        # 建立按鈕框架
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="新增欄位", command=add_field).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="刪除欄位", command=delete_field).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="關閉", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

        load_fields()

    def delete_student(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "請選擇要刪除的學員")
            return
        if messagebox.askyesno("確認", f"確定要刪除選取的 {len(selected)} 位學員嗎？"):
            with sqlite3.connect(DB_FILE) as conn:
                c = conn.cursor()
                for sid in selected:
                    c.execute("DELETE FROM students WHERE id=?", (sid,))
                conn.commit()
            self.load_students()

    def import_students(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV檔案", "*.csv")])
        if not file_path:
            return
        try:
            with open(file_path, newline='', encoding="utf-8") as csvfile:
                reader = csv.DictReader(csvfile)
                rows = list(reader)
        except UnicodeDecodeError:
            with open(file_path, newline='', encoding="cp950") as csvfile:
                reader = csv.DictReader(csvfile)
                rows = list(reader)
        
        # 取得所有自定義欄位
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("""
                SELECT id, field_name, field_type 
                FROM custom_fields 
                ORDER BY display_order
            """)
            custom_fields = c.fetchall()
        
        added = 0
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            for row in rows:
                name = row.get("姓名", "").strip()
                dept = row.get("部門", "").strip()
                if not name:
                    continue
                try:
                    h = hash_name(name)
                    # 插入基本資料
                    c.execute("""
                        INSERT INTO students (name, department, hash, gender, phone, dietary) 
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (name, dept, h,
                          row.get("性別", ""),
                          row.get("連絡電話", ""),
                          row.get("餐飲葷素", "")))
                    
                    student_id = c.lastrowid
                    
                    # 插入其他自定義欄位值
                    for field_id, field_name, field_type in custom_fields:
                        if field_id not in [1, 3, 5]:  # 排除已處理的欄位
                            value = row.get(field_name, "").strip()
                            if value:
                                c.execute("""
                                    INSERT INTO student_custom_values (student_id, field_id, field_value)
                                    VALUES (?, ?, ?)
                                """, (student_id, field_id, value))
                    
                    added += 1
                except sqlite3.IntegrityError:
                    continue
            conn.commit()
        messagebox.showinfo("匯入完成", f"成功匯入 {added} 位學員")
        self.load_students()

    def export_students(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV檔案", "*.csv")]
        )
        if not file_path:
            return

        # 取得所有自定義欄位
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("""
                SELECT id, field_name, field_type 
                FROM custom_fields 
                ORDER BY display_order
            """)
            custom_fields = c.fetchall()

            # 取得所有學員資料
            c.execute("""
                SELECT s.id, s.name, s.department, s.gender, s.phone, s.dietary,
                       GROUP_CONCAT(f.field_name || ':' || v.field_value) as custom_values
                FROM students s
                LEFT JOIN student_custom_values v ON v.student_id = s.id
                LEFT JOIN custom_fields f ON f.id = v.field_id
                GROUP BY s.id
                ORDER BY s.name
            """)
            students = c.fetchall()

        # 準備欄位名稱
        fieldnames = ["姓名", "部門", "性別", "連絡電話", "餐飲葷素"]
        for _, field_name, _ in custom_fields:
            if field_name not in ["性別", "連絡電話", "餐飲葷素"]:
                fieldnames.append(field_name)

        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()

                for student in students:
                    sid, name, dept, gender, phone, dietary, custom_values = student
                    row_data = {
                        "姓名": name,
                        "部門": dept,
                        "性別": gender,
                        "連絡電話": phone,
                        "餐飲葷素": dietary
                    }

                    # 處理自定義欄位值
                    if custom_values:
                        for pair in custom_values.split(','):
                            field_name, value = pair.split(':')
                            if field_name not in ["性別", "連絡電話", "餐飲葷素"]:
                                row_data[field_name] = value

                    writer.writerow(row_data)

            messagebox.showinfo("匯出完成", "學員資料已成功匯出")
        except Exception as e:
            messagebox.showerror("錯誤", f"匯出失敗：{str(e)}")

def main():
    root = tk.Tk()
    root.withdraw()
    login_window = None
    app = None

    def show_login():
        nonlocal login_window
        if login_window is not None:
            try:
                login_window.destroy()
            except Exception:
                pass
        login_window = LoginWindow(root, on_login)
        login_window.deiconify()
        login_window.grab_set()
        login_window.focus_force()

    def on_login(success, user_id, is_admin):
        nonlocal app, login_window
        if success:
            if login_window:
                login_window.withdraw()
            root.deiconify()
            # 登入前先清空 root 內容
            for widget in root.winfo_children():
                if not isinstance(widget, LoginWindow):
                    try:
                        widget.destroy()
                    except Exception:
                        pass
            if app:
                app.destroy()
            app = CheckInApp(root)
            app.user_id = user_id
            app.is_admin = is_admin
            if not is_admin:
                app.user_mgmt_btn.grid_remove()
            app.set_logout_callback(logout)
        else:
            root.destroy()

    def logout():
        nonlocal app, login_window
        if app:
            app.destroy()
            app = None
        root.withdraw()
        root.after(100, show_login)

    show_login()
    root.mainloop()

if __name__ == "__main__":
    main()
