import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import sqlite3
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import win32print
import win32api
import os

# 数据库初始化
def init_db():
    conn = sqlite3.connect("receipts.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS receipts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            receipt_no TEXT,
            receiver TEXT,
            payer TEXT,
            amount REAL,
            reason TEXT,
            date TEXT,
            created_at TEXT
        )
    """)
    conn.commit()
    conn.close()

# 保存到数据库
def save_to_db(data):
    conn = sqlite3.connect("receipts.db")
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO receipts (receipt_no, receiver, payer, amount, reason, date, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        data['receipt_no'], data['receiver'], data['payer'],
        data['amount'], data['reason'], data['date'], data['created_at']
    ))
    conn.commit()
    conn.close()

# 导出 PDF
def export_pdf(data):
    filename = f"{data['receipt_no']}.pdf"
    c = canvas.Canvas(filename, pagesize=A4)
    c.setFont("Helvetica", 12)
    c.drawString(100, 800, "收款收据")
    c.drawString(100, 780, f"编号：{data['receipt_no']}")
    c.drawString(100, 760, f"收款人：{data['receiver']}")
    c.drawString(100, 740, f"付款人：{data['payer']}")
    c.drawString(100, 720, f"金额：¥{data['amount']}")
    c.drawString(100, 700, f"事由：{data['reason']}")
    c.drawString(100, 680, f"日期：{data['date']}")
    c.save()
    messagebox.showinfo("导出成功", f"PDF已保存为 {filename}")

# 打印收据
def print_receipt(data):
    filename = f"{data['receipt_no']}.pdf"
    export_pdf(data)  # 先导出 PDF
    win32api.ShellExecute(
        0,
        "print",
        filename,
        None,
        ".",
        0
    )

# 生成收据
def generate_receipt():
    receipt_no = f"RCPT-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    data = {
        'receipt_no': receipt_no,
        'receiver': receiver.get(),
        'payer': payer.get(),
        'amount': amount.get(),
        'reason': reason.get(),
        'date': date.get(),
        'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    # 显示预览
    preview.config(state='normal')
    preview.delete("1.0", tk.END)
    preview.insert(tk.END, f"""
收款收据
编号：{data['receipt_no']}
收款人：{data['receiver']}
付款人：{data['payer']}
金额：¥{data['amount']}
事由：{data['reason']}
日期：{data['date']}
    """)
    preview.config(state='disabled')

    # 保存数据
    save_to_db(data)

    # 存储当前数据用于导出/打印
    global current_receipt
    current_receipt = data

# 主界面
init_db()
root = tk.Tk()
root.title("收款收据生成器")
root.geometry("500x600")

frame = ttk.Frame(root, padding="10")
frame.pack(fill=tk.BOTH, expand=True)

ttk.Label(frame, text="收款人：").grid(row=0, column=0, sticky=tk.W)
receiver = ttk.Entry(frame, width=30)
receiver.grid(row=0, column=1)

ttk.Label(frame, text="付款人：").grid(row=1, column=0, sticky=tk.W)
payer = ttk.Entry(frame, width=30)
payer.grid(row=1, column=1)

ttk.Label(frame, text="金额（¥）：").grid(row=2, column=0, sticky=tk.W)
amount = ttk.Entry(frame, width=30)
amount.grid(row=2, column=1)

ttk.Label(frame, text="日期：").grid(row=3, column=0, sticky=tk.W)
date = ttk.Entry(frame, width=30)
date.insert(0, datetime.now().strftime('%Y-%m-%d'))
date.grid(row=3, column=1)

ttk.Label(frame, text="事由：").grid(row=4, column=0, sticky=tk.W)
reason = ttk.Entry(frame, width=30)
reason.grid(row=4, column=1)

ttk.Button(frame, text="生成收据", command=generate_receipt).grid(row=5, column=0, pady=10)
ttk.Button(frame, text="导出 PDF", command=lambda: export_pdf(current_receipt)).grid(row=5, column=1, pady=10)
ttk.Button(frame, text="打印收据", command=lambda: print_receipt(current_receipt)).grid(row=6, column=0, columnspan=2)

preview = tk.Text(root, height=10, width=60, state='disabled', bg="#f0f0f0", font=("Courier", 10))
preview.pack(padx=10, pady=10)

current_receipt = {}

root.mainloop()
