import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import sqlite3
import csv
import hashlib
import qrcode
import os
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
        c.execute("""
        CREATE TABLE IF NOT EXISTS attendees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_id INTEGER,
            name TEXT,
            department TEXT,
            hash TEXT,
            FOREIGN KEY (class_id) REFERENCES classes(id)
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER,
            attendee_id INTEGER,
            checkin_time TEXT,
            checkout_time TEXT,
            FOREIGN KEY (session_id) REFERENCES sessions(id),
            FOREIGN KEY (attendee_id) REFERENCES attendees(id)
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS session_attendees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER,
            attendee_id INTEGER,
            FOREIGN KEY (session_id) REFERENCES sessions(id),
            FOREIGN KEY (attendee_id) REFERENCES attendees(id)
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS checkins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER,
            attendee_id INTEGER,
            check_in_time TEXT,
            check_out_time TEXT,
            UNIQUE(session_id, attendee_id)
        )""")
    conn.commit()

class ManageAttendeesDialog(tk.Toplevel):
    def __init__(self, parent, class_id, refresh_callback):
        super().__init__(parent)
        self.class_id = class_id
        self.refresh_callback = refresh_callback
        self.title("管理學員")
        self.geometry("400x400")
        self.resizable(False, False)

        self.tree = ttk.Treeview(self, columns=("name", "dept"), show="headings")
        self.tree.heading("name", text="姓名")
        self.tree.heading("dept", text="部門")
        self.tree.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        self.show_timed_popup = self.show_timed_popup.__get__(self)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=5)

        ttk.Button(btn_frame, text="新增學員", command=self.add_attendee).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="刪除選取", command=self.delete_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="關閉", command=self.destroy).pack(side=tk.RIGHT, padx=5)

        self.load_attendees()

    def load_attendees(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("SELECT id, name, department FROM attendees WHERE class_id=?", (self.class_id,))
            for aid, name, dept in c.fetchall():
                self.tree.insert("", tk.END, iid=aid, values=(name, dept))

    def add_attendee(self):
        def save():
            name = name_var.get().strip()
            dept = dept_var.get().strip()
            if not name:
                messagebox.showwarning("警告", "姓名不能為空")
                return
            with sqlite3.connect(DB_FILE) as conn:
                c = conn.cursor()
                c.execute("SELECT id FROM attendees WHERE class_id=? AND name=?", (self.class_id, name))
                if c.fetchone():
                    messagebox.showwarning("警告", "該學員已存在")
                    return
                h = hash_name(name)
                c.execute("INSERT INTO attendees (class_id, name, department, hash) VALUES (?, ?, ?, ?)",
                          (self.class_id, name, dept, h))
                conn.commit()
            top.destroy()
            self.load_attendees()
            self.refresh_callback()

        top = tk.Toplevel(self)
        top.title("新增學員")
        top.geometry("300x150")
        top.resizable(False, False)

        ttk.Label(top, text="姓名:").pack(pady=5)
        name_var = tk.StringVar()
        ttk.Entry(top, textvariable=name_var).pack(pady=5, fill=tk.X, padx=10)

        ttk.Label(top, text="部門:").pack(pady=5)
        dept_var = tk.StringVar()
        ttk.Entry(top, textvariable=dept_var).pack(pady=5, fill=tk.X, padx=10)

        ttk.Button(top, text="儲存", command=save).pack(pady=10)

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "請先選擇欲刪除的學員")
            return
        confirm = messagebox.askyesno("確認刪除", f"確定刪除選取的 {len(selected)} 位學員嗎？")
        if not confirm:
            return
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            for aid in selected:
                c.execute("DELETE FROM attendees WHERE id=?", (aid,))
            conn.commit()
        self.load_attendees()
        self.refresh_callback()

class LoginWindow(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback = callback
        self.title("登入系統")
        self.geometry("300x200")
        self.resizable(False, False)
        
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
        
        ttk.Label(frame, text="帳號：").pack(pady=5)
        self.username_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.username_var).pack(pady=5, fill=tk.X)
        
        ttk.Label(frame, text="密碼：").pack(pady=5)
        self.password_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.password_var, show="*").pack(pady=5, fill=tk.X)
        
        ttk.Button(frame, text="登入", command=self.login).pack(pady=10)
        
        # 綁定 Enter 鍵
        self.bind("<Return>", lambda e: self.login())
        
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
                self.callback(True, user[0], bool(user[1]))
                self.destroy()
            else:
                messagebox.showerror("錯誤", "帳號或密碼錯誤")

class CheckInApp:
    def __init__(self, root):
        self.root = root
        self.org_info = self.load_org_info()
        self.root.title(f"{self.org_info.get('org_name', '活動(課程)簽到系統')}-活動(課程)簽到系統")
        self.root.geometry("1200x800")

        # 確保資料庫初始化
        init_db()

        self.class_id = None
        self.session_id = None
        self.user_id = None
        self.is_admin = False

        self.setup_ui()
        self.load_classes()
        self.tts_engine = pyttsx3.init()
        self.tts_engine.setProperty("rate", 160)

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')

        top_frame = ttk.Frame(self.root)
        top_frame.pack(pady=10, fill=tk.X)

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
        ttk.Button(top_frame, text="新增學員", command=self.add_attendee).grid(row=0, column=4, padx=5)
        ttk.Button(top_frame, text="管理學員", command=self.open_manage_dialog).grid(row=0, column=5, padx=5)
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
        self.tree = ttk.Treeview(self.root, columns=("姓名", "部門", "簽到時間", "簽退時間"), show="headings")
        self.tree.heading("姓名", text="姓名")
        self.tree.heading("部門", text="部門")
        self.tree.heading("簽到時間", text="簽到時間")
        self.tree.heading("簽退時間", text="簽退時間")
        self.tree.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        # 狀態區：目前時間 + 統計資訊
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)

        self.time_label = ttk.Label(bottom_frame, text="")
        self.time_label.pack(side=tk.LEFT)

        self.stats_label = ttk.Label(bottom_frame, text="", foreground="blue")
        self.stats_label.pack(side=tk.RIGHT)

        self.update_time()
        self.update_stats()

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
                c.execute("SELECT COUNT(*) FROM attendees WHERE class_id=?", (self.class_id,))
                total = c.fetchone()[0]

                c.execute("SELECT COUNT(DISTINCT attendee_id) FROM checkins WHERE session_id=? AND check_in_time IS NOT NULL", (self.session_id,))
                checked_in = c.fetchone()[0]

                unchecked_in = total - checked_in

                c.execute("SELECT COUNT(DISTINCT attendee_id) FROM checkins WHERE session_id=? AND check_out_time IS NOT NULL", (self.session_id,))
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

    def add_attendee(self):
        if not self.class_id:
            messagebox.showwarning("警告", "請先選擇活動(課程)")
            return
        top = tk.Toplevel(self.root)
        top.title("新增學員")
        top.geometry("300x220")
        top.resizable(False, False)

        ttk.Label(top, text="姓名:").pack(pady=5)
        name_var = tk.StringVar()
        ttk.Entry(top, textvariable=name_var).pack(pady=5, fill=tk.X, padx=10)

        ttk.Label(top, text="部門:").pack(pady=5)
        dept_var = tk.StringVar()
        ttk.Entry(top, textvariable=dept_var).pack(pady=5, fill=tk.X, padx=10)

        def save():
            name = name_var.get().strip()
            dept = dept_var.get().strip()
            if not name:
                messagebox.showwarning("警告", "姓名不能為空")
                return
            with sqlite3.connect(DB_FILE) as conn:
                c = conn.cursor()
                c.execute("SELECT id FROM attendees WHERE class_id=? AND name=?", (self.class_id, name))
                if c.fetchone():
                    messagebox.showwarning("警告", "該學員已存在")
                    return
                h = hash_name(name)
                c.execute("INSERT INTO attendees (class_id, name, department, hash) VALUES (?, ?, ?, ?)",
                          (self.class_id, name, dept, h))
                conn.commit()
            top.destroy()
            self.load_attendees()
            self.update_stats()

        ttk.Button(top, text="儲存", command=save).pack(pady=10)

    def load_attendees(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        if not self.class_id or not self.session_id:
            return
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            # 左連接取出所有該活動(課程)該週次學員簽到簽退時間
            c.execute("""
            SELECT a.id, a.name, a.department,
                ci.check_in_time, ci.check_out_time
            FROM attendees a
            LEFT JOIN checkins ci ON ci.attendee_id=a.id AND ci.session_id=?
            WHERE a.class_id=?
            ORDER BY a.name
            """, (self.session_id, self.class_id))
            rows = c.fetchall()
            for aid, name, dept, cin, cout in rows:
                self.tree.insert("", tk.END, iid=aid, values=(
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
                c.execute("SELECT id, name FROM attendees WHERE class_id=?", (self.class_id,))
                attendees = c.fetchall()

                matched_attendee = None
                for aid, name in attendees:
                    hashed = hash_name(name)
                    if code == hashed or code == hashed[:10]:
                        matched_attendee = (aid, name)
                        break

                if not matched_attendee:
                    self.show_timed_popup("查無此學員或備用碼錯誤", popup_type="error", duration=5)
                    return

                aid, name = matched_attendee
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                c.execute("SELECT check_in_time, check_out_time FROM checkins WHERE session_id=? AND attendee_id=?",
                          (self.session_id, aid))
                row = c.fetchone()

                if not row:
                    c.execute("INSERT INTO checkins (session_id, attendee_id, check_in_time) VALUES (?, ?, ?)",
                              (self.session_id, aid, now_str))
                    self.show_timed_popup(f"{name} 簽到成功", popup_type="success", duration=5)
                else:
                    cin, cout = row
                    if cin and not cout:
                        c.execute("UPDATE checkins SET check_out_time=? WHERE session_id=? AND attendee_id=?",
                                  (now_str, self.session_id, aid))
                        self.show_timed_popup(f"{name} 簽退成功", popup_type="success", duration=5)
                    elif cin and cout:
                        self.show_timed_popup(f"{name} 已簽退，無法重複簽到", popup_type="info", duration=5)
                    else:
                        c.execute("UPDATE checkins SET check_in_time=? WHERE session_id=? AND attendee_id=?",
                                  (now_str, self.session_id, aid))
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
            c.execute("SELECT id, name FROM attendees WHERE class_id=?", (self.class_id,))
            attendees = c.fetchall()

            matched_attendee = None
            for aid, name in attendees:
                if hash_name(name) == code:
                    matched_attendee = (aid, name)
                    break

            if not matched_attendee:
                self.show_timed_popup("查無此學員或QR碼錯誤", popup_type="error", duration=5)
                return

            aid, name = matched_attendee
            c.execute("SELECT check_in_time, check_out_time FROM checkins WHERE session_id=? AND attendee_id=?",
                      (self.session_id, aid))
            row = c.fetchone()
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if not row:
                c.execute("INSERT INTO checkins (session_id, attendee_id, check_in_time) VALUES (?, ?, ?)",
                          (self.session_id, aid, now_str))
                self.show_timed_popup(f"{name} 簽到成功", popup_type="success", duration=5)
            else:
                cin, cout = row
                if cin and not cout:
                    c.execute("UPDATE checkins SET check_out_time=? WHERE session_id=? AND attendee_id=?",
                              (now_str, self.session_id, aid))
                    self.show_timed_popup(f"{name} 簽退成功", popup_type="success", duration=5)
                elif cin and cout:
                    self.show_timed_popup(f"{name} 已簽退，無法重複簽到", popup_type="info", duration=5)
                else:
                    c.execute("UPDATE checkins SET check_in_time=? WHERE session_id=? AND attendee_id=?",
                              (now_str, self.session_id, aid))
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
                c.execute("SELECT id FROM attendees WHERE class_id=? AND name=?", (self.class_id, name))
                if c.fetchone():
                    continue
                h = hash_name(name)
                c.execute("INSERT INTO attendees (class_id, name, department, hash) VALUES (?, ?, ?, ?)",
                          (self.class_id, name, dept, h))
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
                FROM attendees a
                LEFT JOIN checkins ci ON ci.attendee_id=a.id AND ci.session_id=?
                WHERE a.class_id=?
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
            c.execute("SELECT name FROM attendees WHERE class_id=?", (self.class_id,))
            names = [row[0] for row in c.fetchall()]
            if not names:
                messagebox.showwarning("警告", "此課堂尚無學員")
                return

            for name in names:
                h = hash_name(name)
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
            for aid in selected:
                c.execute("DELETE FROM attendees WHERE id=?", (aid,))
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

def main():
    root = tk.Tk()
    root.withdraw()  # 隱藏主視窗
    
    def on_login(success, user_id, is_admin):
        if success:
            root.deiconify()  # 顯示主視窗
            app = CheckInApp(root)
            app.user_id = user_id
            app.is_admin = is_admin
            if not is_admin:
                app.user_mgmt_btn.grid_remove()  # 隱藏使用者管理按鈕
            root.mainloop()
        else:
            root.destroy()
    
    login_window = LoginWindow(root, on_login)
    root.mainloop()

if __name__ == "__main__":
    main()
