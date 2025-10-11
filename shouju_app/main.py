import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime
import db
import pdf_export
import printer
from utils import num_to_chinese_upper
import os
os.chdir(os.path.dirname(__file__))


class ReceiptApp:
    def __init__(self, root):
        self.root = root
        self.root.title("收款收据生成器")
        self.root.geometry("900x600")
        self.receipt_no = f"NO {datetime.now().strftime('%Y%m%d%H%M%S')}"
        self.items = []

        db.init_db()
        self.setup_style()
        self.build_ui()
        self.add_row()

    def setup_style(self):
        style = ttk.Style()
        style.configure("Treeview",
                        rowheight=25,
                        borderwidth=1,
                        relief="solid")
        style.map("Treeview", background=[("selected", "#d0e0ff")])
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))

    def build_ui(self):
        ttk.Label(self.root, text="收款收据", font=("Arial", 16)).pack(pady=5)
        ttk.Label(self.root, text=f"编号：{self.receipt_no}", foreground="red").pack()

        frame_top = ttk.Frame(self.root)
        frame_top.pack(pady=5)
        ttk.Label(frame_top, text="客户名称：").grid(row=0, column=0)
        self.customer_name = ttk.Entry(frame_top, width=40)
        self.customer_name.grid(row=0, column=1)

        columns = ("品名及规格", "单位", "数量", "单价", "金额", "备注")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", height=8, style="Treeview")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")
        self.tree.pack(pady=10)
        self.tree.bind("<Button-1>", self.on_click)

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack()
        ttk.Button(btn_frame, text="添加一行", command=self.add_row).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="删除选中行", command=self.delete_row).grid(row=0, column=1, padx=5)

        total_frame = ttk.Frame(self.root)
        total_frame.pack(pady=5)
        ttk.Label(total_frame, text="合计金额：¥").grid(row=0, column=0)
        self.total_amount = tk.StringVar()
        ttk.Label(total_frame, textvariable=self.total_amount).grid(row=0, column=1)
        ttk.Label(total_frame, text="大写金额：").grid(row=0, column=2)
        self.upper_amount = tk.StringVar()
        ttk.Label(total_frame, textvariable=self.upper_amount).grid(row=0, column=3)

        info_frame = ttk.Frame(self.root)
        info_frame.pack(pady=5)
        ttk.Label(info_frame, text="开票人：").grid(row=0, column=0)
        self.issuer = ttk.Entry(info_frame, width=20)
        self.issuer.grid(row=0, column=1)
        ttk.Label(info_frame, text="日期：").grid(row=0, column=2)
        self.date = DateEntry(info_frame, width=12, date_pattern='yyyy-mm-dd')
        self.date.grid(row=0, column=3)

        action_frame = ttk.Frame(self.root)
        action_frame.pack(pady=10)
        ttk.Button(action_frame, text="保存收据", command=self.save_receipt).grid(row=0, column=0, padx=5)
        ttk.Button(action_frame, text="导出 PDF", command=self.export_pdf).grid(row=0, column=1, padx=5)
        ttk.Button(action_frame, text="打印收据", command=self.print_receipt).grid(row=0, column=2, padx=5)
        ttk.Button(action_frame, text="历史记录", command=self.show_history).grid(row=0, column=3, padx=5)

    def add_row(self):
        self.tree.insert("", tk.END, values=("", "", "", "", "", ""))
        self.update_total()

    def delete_row(self):
        selected = self.tree.selection()
        for item in selected:
            self.tree.delete(item)
        self.update_total()

    def update_total(self):
        total = 0
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            try:
                qty = float(values[2])
                price = float(values[3])
                amount = qty * price
                self.tree.set(item, column="金额", value=f"{amount:.2f}")
                total += amount
            except:
                continue
        self.total_amount.set(f"{total:.2f}")
        self.upper_amount.set(num_to_chinese_upper(total))

    def on_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        col_index = int(column.replace("#", "")) - 1
        self.edit_cell(row_id, col_index)

    def edit_cell(self, row_id, col_index):
        x, y, width, height = self.tree.bbox(row_id, f"#{col_index+1}")
        value = self.tree.item(row_id)["values"][col_index]

        entry = tk.Entry(self.root)
        entry.place(x=x + self.tree.winfo_rootx() - self.root.winfo_rootx(),
                    y=y + self.tree.winfo_rooty() - self.root.winfo_rooty(),
                    width=width, height=height)
        entry.insert(0, value)
        entry.focus()

        def save_and_next(event=None):
            new_value = entry.get()
            values = list(self.tree.item(row_id)["values"])
            values[col_index] = new_value
            self.tree.item(row_id, values=values)
            entry.destroy()
            self.update_total()

            next_col = col_index + 1
            if next_col >= len(values):
                next_row = self.tree.next(row_id)
                if next_row:
                    self.edit_cell(next_row, 0)
            else:
                self.edit_cell(row_id, next_col)

        entry.bind("<Return>", save_and_next)
        entry.bind("<Tab>", lambda e: save_and_next())
        entry.bind("<FocusOut>", lambda e: entry.destroy())

    def collect_data(self):
        items = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            items.append(values)
        return {
            "receipt_no": self.receipt_no,
            "customer_name": self.customer_name.get(),
            "items": items,
            "total": self.total_amount.get(),
            "upper_total": self.upper_amount.get(),
            "issuer": self.issuer.get(),
            "date": self.date.get_date().strftime('%Y-%m-%d'),
            "created_at": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

    def save_receipt(self):
        data = self.collect_data()
        db.save_receipt(data)
        messagebox.showinfo("保存成功", "收据已保存到数据库")

    def export_pdf(self):
        data = self.collect_data()
        pdf_export.export(data)
        messagebox.showinfo("导出成功", f"PDF已保存为 {data['receipt_no']}.pdf")

    def print_receipt(self):
        data = self.collect_data()
        printer.print_pdf(data)

    def show_history(self):
        db.show_history_window(self.root)

if __name__ == "__main__":
    root = tk.Tk()
    app = ReceiptApp(root)
    root.mainloop()
