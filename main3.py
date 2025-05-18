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

class CheckInApp:
    def __init__(self, root):
        self.root = root
        self.root.title("課堂簽到系統")
        self.root.geometry("1000x700")

        init_db()

        self.class_id = None
        self.session_id = None

        self.setup_ui()
        self.load_classes()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')  # 使用較現代化的主題

        top_frame = ttk.Frame(self.root)
        top_frame.pack(pady=10, fill=tk.X)

        ttk.Label(top_frame, text="選擇課堂:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.class_combo = ttk.Combobox(top_frame, state="readonly")
        self.class_combo.grid(row=0, column=1, sticky=tk.W)
        self.class_combo.bind("<<ComboboxSelected>>", lambda e: self.select_class())

        ttk.Button(top_frame, text="新增課堂", command=self.add_class).grid(row=0, column=2, padx=5)
        ttk.Button(top_frame, text="新增週次", command=self.add_session).grid(row=0, column=3, padx=5)
        ttk.Button(top_frame, text="新增學員", command=self.add_attendee).grid(row=0, column=4, padx=5)
        ttk.Button(top_frame, text="管理學員", command=self.open_manage_dialog).grid(row=0, column=5, padx=5)
        ttk.Button(top_frame, text="產生QR Code", command=self.generate_qrcodes).grid(row=0, column=6, padx=5)
        ttk.Button(top_frame, text="刪除選取名單", command=self.delete_selected_attendees).grid(row=0, column=7, padx=5)

        ttk.Label(top_frame, text="選擇週次:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.session_combo = ttk.Combobox(top_frame, state="readonly")
        self.session_combo.grid(row=1, column=1, sticky=tk.W)
        self.session_combo.bind("<<ComboboxSelected>>", lambda e: self.select_session())

        ttk.Button(top_frame, text="匯入名單", command=self.import_attendees).grid(row=1, column=2, padx=5)
        ttk.Button(top_frame, text="匯出記錄", command=self.export_records).grid(row=1, column=3, padx=5)

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

    def update_time(self):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_label.config(text=f"目前時間: {now}")
        self.root.after(1000, self.update_time)

    def update_stats(self):
        if not self.class_id or not self.session_id:
            self.stats_label.config(text="請先選擇課堂及週次")
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
        name = simpledialog.askstring("新增課堂", "輸入課堂名稱")
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
            messagebox.showwarning("警告", "請先選擇課堂")
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
            messagebox.showwarning("警告", "請先選擇課堂")
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
            # 左連接取出所有該課堂該週次學員簽到簽退時間
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

    import platform
    import winsound  # 只適用於 Windows

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

        if platform.system() == "Windows":
            if popup_type == "success":
                winsound.MessageBeep(winsound.MB_OK)
            elif popup_type == "error":
                winsound.MessageBeep(winsound.MB_ICONHAND)
            elif popup_type == "warning":
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
            else:
                winsound.MessageBeep()

        def update_countdown(sec):
            if sec > 0:
                countdown_var.set(f"{sec} 秒後自動關閉")
                self.root.after(1000, lambda: update_countdown(sec - 1))
            else:
                popup.destroy()
                self.scan_entry.focus_set()

        update_countdown(duration)

    def import_attendees(self):
        if not self.class_id:
            messagebox.showwarning("警告", "請先選擇課堂")
            return
        file_path = filedialog.askopenfilename(filetypes=[("CSV檔案", "*.csv")])
        if not file_path:
            return
        with open(file_path, newline='', encoding="utf-8") as csvfile:
            reader = csv.DictReader(csvfile)
            added = 0
            with sqlite3.connect(DB_FILE) as conn:
                c = conn.cursor()
                for row in reader:
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

    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen.canvas import Canvas
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle
    from datetime import datetime

    from reportlab.pdfgen.canvas import Canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import os
    import platform
    from datetime import datetime


    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle
    from datetime import datetime

    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen.canvas import Canvas
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle
    from datetime import datetime

    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog, simpledialog
    import sqlite3
    import csv
    import hashlib
    import qrcode
    import os
    from datetime import datetime
    from PIL import Image, ImageTk
    from tkcalendar import DateEntry
    import platform
    import winsound

    # 新增：PDF 匯出所需
    from reportlab.pdfgen.canvas import Canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    DB_FILE = "checkin.db"
    QR_FOLDER = "qrcodes"
    QR_SEED = "secure_seed_2024"

    if not os.path.exists(QR_FOLDER):
        os.makedirs(QR_FOLDER)

    # ... 其他函數與類別省略（保留原始內容）

    class CheckInApp:
        def __init__(self, root):
            self.root = root
            self.class_id = None
            self.session_id = None
            # 其他初始化...

        # 匯出 PDF 函數（已整合中文與列印日期）
            # 匯出 PDF 函數（已整合中文與列印日期）
    def export_records(self):
        if not self.class_id or not self.session_id:
            messagebox.showwarning("警告", "請先選擇課堂及週次")
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

            # 課堂名稱
            c.execute("SELECT name FROM classes WHERE id=?", (self.class_id,))
            row = c.fetchone()
            class_name = row[0] if row else "(未知課堂)"

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

        # 其餘類別與主程式入口省略...（如 main()）

    # 其餘類別與主程式入口省略...（如 main()）

    def generate_qrcodes(self):
        if not self.class_id:
            messagebox.showwarning("警告", "請先選擇課堂")
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
                qr = qrcode.QRCode(
                    version=1,
                    error_correction=qrcode.constants.ERROR_CORRECT_H,
                    box_size=10,
                    border=4,
                )
                qr.add_data(h)
                qr.make(fit=True)
                img = qr.make_image(fill_color="black", back_color="white")
                img.save(os.path.join(QR_FOLDER, f"{name}.png"))
        messagebox.showinfo("完成", f"QR Code已儲存至資料夾 {QR_FOLDER}")

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
            messagebox.showwarning("警告", "請先選擇課堂")
            return
        ManageAttendeesDialog(self.root, self.class_id, self.update_stats)

def main():
    root = tk.Tk()
    app = CheckInApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
