import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox

def init_db():
    conn = sqlite3.connect("receipts.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS receipts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            receipt_no TEXT,
            customer_name TEXT,
            items TEXT,
            total TEXT,
            upper_total TEXT,
            issuer TEXT,
            date TEXT,
            created_at TEXT
        )
    """)
    conn.commit()
    conn.close()

def save_receipt(data):
    conn = sqlite3.connect("receipts.db")
    cursor = conn.cursor()
    items_str = str(data["items"])
    cursor.execute("""
        INSERT INTO receipts (receipt_no, customer_name, items, total, upper_total, issuer, date, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data["receipt_no"], data["customer_name"], items_str,
        data["total"], data["upper_total"], data["issuer"],
        data["date"], data["created_at"]
    ))
    conn.commit()
    conn.close()

def show_history_window(root):
    win = tk.Toplevel(root)
    win.title("历史收据记录")
    win.geometry("800x500")

    search_var = tk.StringVar()
    ttk.Label(win, text="搜索客户名称：").pack(anchor=tk.W, padx=10, pady=5)
    search_entry = ttk.Entry(win, textvariable=search_var, width=40)
    search_entry.pack(anchor=tk.W, padx=10)

    tree = ttk.Treeview(win, columns=("编号", "客户", "金额", "日期", "开票人"), show="headings")
    for col in tree["columns"]:
        tree.heading(col, text=col)
        tree.column(col, width=150)
    tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def load_data(keyword=""):
        tree.delete(*tree.get_children())
        conn = sqlite3.connect("receipts.db")
        cursor = conn.cursor()
        cursor.execute("""
            SELECT receipt_no, customer_name, total, date, issuer
            FROM receipts
            WHERE customer_name LIKE ?
            ORDER BY created_at DESC
        """, (f"%{keyword}%",))
        for row in cursor.fetchall():
            tree.insert("", tk.END, values=row)
        conn.close()

    def on_search(*args):
        load_data(search_var.get())

    search_var.trace_add("write", on_search)
    load_data()

    def view_selected():
        selected = tree.focus()
        if not selected:
            messagebox.showwarning("未选择", "请先选择一条记录")
            return
        values = tree.item(selected)["values"]
        receipt_no = values[0]
        conn = sqlite3.connect("receipts.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM receipts WHERE receipt_no=?", (receipt_no,))
        record = cursor.fetchone()
        conn.close()
        if not record:
            return

        preview_win = tk.Toplevel(win)
        preview_win.title("收据详情预览")
        preview_win.geometry("800x500")

        ttk.Label(preview_win, text="收款收据", font=("Arial", 16)).pack(pady=5)
        ttk.Label(preview_win, text=f"编号：{record[1]}", foreground="red").pack()
        ttk.Label(preview_win, text=f"客户名称：{record[2]}").pack()
        ttk.Label(preview_win, text=f"开票人：{record[6]}    日期：{record[7]}").pack()

        columns = ("品名及规格", "单位", "数量", "单价", "金额", "备注")
        tree_preview = ttk.Treeview(preview_win, columns=columns, show="headings", height=8)
        for col in columns:
            tree_preview.heading(col, text=col)
            tree_preview.column(col, width=120, anchor="center")
        tree_preview.pack(pady=10)

        try:
            items = eval(record[3])
            for item in items:
                tree_preview.insert("", tk.END, values=item)
        except:
            tree_preview.insert("", tk.END, values=["数据解析失败"] * 6)

        ttk.Label(preview_win, text=f"合计金额：¥{record[4]} （{record[5]}）", font=("Arial", 12)).pack(pady=10)

    ttk.Button(win, text="查看详情", command=view_selected).pack(pady=5)
