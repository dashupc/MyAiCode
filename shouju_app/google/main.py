import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from datetime import datetime
import os
import sys
import traceback
import json # 用于处理明细项目的JSON存储

# --- ReportLab 相关的库用于 PDF 生成 ---
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- 字体注册 ---
try:
    pdfmetrics.registerFont(TTFont('SimHei', 'simhei.ttf'))
    CHINESE_FONT = 'SimHei'
except Exception:
    CHINESE_FONT = 'Helvetica'
    print("警告：SimHei 字体文件未找到，中文可能无法正常显示。")
# --- 字体注册结束 ---


# --- 辅助函数：金额大写转换 ---
def convert_to_chinese_currency(amount):
    """将数字金额转换为中文大写金额 (简化版)"""
    if amount == 0:
        return "零圆整"
    
    amount_str = f"{amount:.2f}"
    return f"人民币 {amount_str} 元 (大写仅供参考)"


# --- 数据库操作类 ---
class DatabaseManager:
    def __init__(self, db_name="receipts.db"):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self._create_table()

    def _create_table(self):
        """创建收据表，包含所有字段"""
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS receipts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                receipt_date TEXT,
                customer_name TEXT,
                amount REAL,
                description TEXT, -- 存储JSON格式的明细列表
                pdf_path TEXT
            )
        """)
        self.conn.commit()

    def add_receipt(self, date, customer, amount, description_json, pdf_path=""):
        """插入新的收据记录"""
        sql = "INSERT INTO receipts (receipt_date, customer_name, amount, description, pdf_path) VALUES (?, ?, ?, ?, ?)"
        self.cursor.execute(sql, (date, customer, amount, description_json, pdf_path))
        self.conn.commit()
        return self.cursor.lastrowid

    # (其他数据库查询函数保持不变)
    def get_all_receipts(self):
        self.cursor.execute("SELECT id, receipt_date, customer_name, amount, description FROM receipts ORDER BY id DESC")
        return self.cursor.fetchall()
        
    def get_receipt_by_id(self, receipt_id):
        self.cursor.execute("SELECT * FROM receipts WHERE id=?", (receipt_id,))
        cols = [column[0] for column in self.cursor.description]
        row = self.cursor.fetchone()
        return dict(zip(cols, row)) if row else None

    def update_pdf_path(self, receipt_id, pdf_path):
        sql = "UPDATE receipts SET pdf_path = ? WHERE id = ?"
        self.cursor.execute(sql, (pdf_path, receipt_id))
        self.conn.commit()

    def close(self):
        self.conn.close()

# --- 核心 PDF 生成函数 (基于图片模板) ---
def generate_pdf(data):
    """
    根据收据数据生成 PDF 文件，严格按照提供的图片模板布局。
    :param data: 包含收据信息的字典 {id, receipt_date, customer_name, amount, description}
                 注意：description 现在是包含明细项目的 JSON 字符串。
    """
    pdf_dir = "Receipts_PDFs"
    if not os.path.exists(pdf_dir):
        os.makedirs(pdf_dir)
        
    filename = os.path.join(pdf_dir, f"Receipt_{data['id']}.pdf")
    
    try:
        # 解析明细项目
        items = json.loads(data['description'])
        
        c = canvas.Canvas(filename, pagesize=A4)
        width, height = A4
        margin_x = 1.5 * cm
        start_y = height - 2 * cm
        table_start_y = height - 5 * cm
        row_height = 1.2 * cm
        
        # ... (顶部信息和日期绘制代码保持不变) ...
        # --- 顶部信息 ---
        c.setFont(CHINESE_FONT, 20)
        c.drawCentredString(width / 2, start_y, "收 款 收 据")
        c.line(width / 2 - 2 * cm, start_y - 0.05 * cm, width / 2 + 2 * cm, start_y - 0.05 * cm)
        c.line(width / 2 - 2 * cm, start_y - 0.15 * cm, width / 2 + 2 * cm, start_y - 0.15 * cm)

        c.setFont("Helvetica-Bold", 14)
        receipt_no = f"NO {str(data['id']).zfill(7)}"
        c.setFillColor(colors.red)
        c.drawString(width - margin_x - c.stringWidth(receipt_no, "Helvetica-Bold", 14), start_y, receipt_no)
        c.setFillColor(colors.black)
        
        c.setFont(CHINESE_FONT, 12)
        c.drawString(margin_x, start_y - 1.5 * cm, f"客户名称: {data['customer_name']}")
        
        date_parts = data['receipt_date'].split(' ')[0].split('-')
        c.drawString(width - 7 * cm, start_y - 1.5 * cm, f"年   月   日")
        c.setFont("Helvetica", 12) 
        c.drawString(width - 6.5 * cm, start_y - 1.5 * cm, date_parts[0]) 
        c.drawString(width - 4.5 * cm, start_y - 1.5 * cm, date_parts[1]) 
        c.drawString(width - 2.5 * cm, start_y - 1.5 * cm, date_parts[2]) 
        
        # --- 绘制表格 ---
        total_width = width - 2 * margin_x
        # 定义列宽
        col_widths = [
            total_width * 0.35, # 品名及规格 (0)
            total_width * 0.08, # 单位 (1)
            total_width * 0.08, # 数量 (2)
            total_width * 0.08, # 单价 (3)
            total_width * 0.31, # 金额区域 (4)
            total_width * 0.10, # 备注 (5)
        ]
        
        x_positions = [margin_x]
        current_x = margin_x
        for w in col_widths:
            current_x += w
            x_positions.append(current_x)
            
        num_fixed_rows = 4 # 模板固定 4 行
        
        # 1. 表格头部
        c.setLineWidth(1)
        c.setFont(CHINESE_FONT, 10)
        c.rect(x_positions[0], table_start_y, total_width, row_height) # 绘制头部外框

        header_labels = ["品名及规格", "单位", "数量", "单价", "金 额", "备 注"]
        for i in range(len(header_labels)):
            if i < 4 or i == 5:
                c.line(x_positions[i+1], table_start_y, x_positions[i+1], table_start_y + row_height)
            c.drawCentredString(x_positions[i] + col_widths[i] / 2, table_start_y + row_height / 3, header_labels[i])

        # 金额区域内部划分 (同上，简化)
        amount_start_x = x_positions[4]
        amount_width = col_widths[4]
        c.line(amount_start_x, table_start_y + row_height / 2, amount_start_x + amount_width, table_start_y + row_height / 2) 
        for i, label in enumerate(["百|十|万", "千|百|十|元", "角|分"]):
            sub_col_widths = [amount_width * 0.33, amount_width * 0.44, amount_width * 0.23]
            c.setFont(CHINESE_FONT, 8)
            c.drawCentredString(x_positions[4] + sum(sub_col_widths[:i]) + sub_col_widths[i] / 2, table_start_y + row_height * 0.75, label)
            if i < 2:
                c.line(x_positions[4] + sum(sub_col_widths[:i+1]), table_start_y, x_positions[4] + sum(sub_col_widths[:i+1]), table_start_y + row_height) 
        
        # 2. 绘制数据行
        current_y = table_start_y
        c.setFont(CHINESE_FONT, 10)
        
        for i in range(num_fixed_rows):
            current_y -= row_height
            c.rect(x_positions[0], current_y, total_width, row_height) # 绘制行外框

            for j in range(1, len(x_positions) - 1):
                c.line(x_positions[j], current_y, x_positions[j], current_y + row_height)
            
            # 填充数据 (只填充 items 列表中存在的行)
            if i < len(items):
                item = items[i]
                line_total = item['qty'] * item['price']
                
                # 品名及规格
                c.drawString(x_positions[0] + 0.2 * cm, current_y + row_height / 3, item['name'])
                # 单位
                c.drawCentredString(x_positions[1] + col_widths[1] / 2, current_y + row_height / 3, item['unit'])
                # 数量
                c.drawCentredString(x_positions[2] + col_widths[2] / 2, current_y + row_height / 3, str(item['qty']))
                # 单价
                c.drawRightString(x_positions[4] - col_widths[3] * 0.05, current_y + row_height / 3, f"{item['price']:.2f}")
                # 金额
                c.drawRightString(x_positions[5] - 0.2 * cm, current_y + row_height / 3, f"{line_total:.2f}")

        # --- 底部信息 ---
        
        current_y -= 0.5 * cm # 底部从表格下方开始
        c.setLineWidth(1)
        c.line(margin_x, current_y, width - margin_x, current_y) 
        
        # 大写金额
        c.setFont(CHINESE_FONT, 12)
        chinese_amount = convert_to_chinese_currency(data['amount'])
        
        # 绘制大写金额的格子 (简化只画格子)
        chinese_units = ['佰', '拾', '万', '仟', '佰', '拾', '元', '角', '分']
        unit_width = (width - 2 * margin_x) / 11
        start_x = margin_x + 3.5 * cm
        
        # 绘制汉字单位格子
        for i, unit in enumerate(chinese_units):
            x = start_x + i * unit_width
            y = current_y - 2.5 * cm
            c.rect(x, y, unit_width, row_height) 
            c.setFont(CHINESE_FONT, 8)
            c.drawCentredString(x + unit_width / 2, y + row_height + 0.1 * cm, unit)

        c.setFont(CHINESE_FONT, 12)
        c.drawString(margin_x, current_y - 1.5 * cm, "合计金额 (大写)")
        c.drawString(start_x, current_y - 1.5 * cm, chinese_amount)
        
        # 小写金额 ¥
        c.drawString(width - 5 * cm, current_y - 1.5 * cm, "¥")
        c.setFont("Helvetica", 12)
        c.drawRightString(width - 1.5 * cm, current_y - 1.5 * cm, f"{data['amount']:.2f}")

        # 底部签名栏
        c.setFont(CHINESE_FONT, 12)
        bottom_y = current_y - 4 * cm
        c.drawString(margin_x, bottom_y, "单位盖章")
        c.drawString(width / 2 - 1.5 * cm, bottom_y, "收款人")
        c.drawString(width - 4 * cm, bottom_y, "开票人")

        c.save()
        return filename
        
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("PDF 错误", f"生成 PDF 失败: {e}")
        return ""

# --- GUI 应用程序类 ---
class ReceiptApp:
    def __init__(self, master):
        self.master = master
        master.title("收据开具与管理系统 (表格填写版)")
        master.geometry("1000x800")
        
        self.db = DatabaseManager()
        self.item_list = [] # 存储当前收据的所有明细项目
        self.total_amount_var = tk.DoubleVar(value=0.0) # 自动计算的总金额
        
        self.create_widgets()
        self.load_receipts()

    def create_widgets(self):
        # 1. 客户信息和总金额框架
        info_frame = ttk.LabelFrame(self.master, text="客户信息与总计", padding="10")
        info_frame.pack(padx=10, pady=5, fill="x")

        ttk.Label(info_frame, text="客户名称:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.customer_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.customer_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(info_frame, text="总金额 (自动计算):").grid(row=0, column=2, padx=15, pady=5, sticky="w")
        self.total_amount_label = ttk.Label(info_frame, textvariable=self.total_amount_var, font=('Arial', 12, 'bold'), foreground='red')
        self.total_amount_label.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        # 2. 明细输入框架
        item_input_frame = ttk.LabelFrame(self.master, text="收据明细项目填写", padding="10")
        item_input_frame.pack(padx=10, pady=5, fill="x")
        
        item_fields = [
            ("品名及规格", "item_name_var", tk.StringVar),
            ("单位", "item_unit_var", tk.StringVar),
            ("数量", "item_qty_var", tk.DoubleVar),
            ("单价 (¥)", "item_price_var", tk.DoubleVar),
        ]
        
        # 创建明细输入控件
        self.item_vars = {}
        for i, (label_text, var_name, var_type) in enumerate(item_fields):
            ttk.Label(item_input_frame, text=label_text).grid(row=0, column=i * 2, padx=5, pady=5)
            var = var_type()
            self.item_vars[var_name] = var
            ttk.Entry(item_input_frame, textvariable=var, width=15).grid(row=0, column=i * 2 + 1, padx=5, pady=5)

        ttk.Button(item_input_frame, text="添加明细", command=self.add_item).grid(row=0, column=len(item_fields) * 2, padx=10, pady=5)
        ttk.Button(item_input_frame, text="清空明细", command=self.clear_items).grid(row=0, column=len(item_fields) * 2 + 1, padx=5, pady=5)

        # 3. 明细显示表格
        item_display_frame = ttk.Frame(self.master)
        item_display_frame.pack(padx=10, pady=5, fill="x")
        
        self.items_tree = ttk.Treeview(item_display_frame, columns=("Name", "Unit", "Qty", "Price", "Subtotal"), show='headings', height=5)
        self.items_tree.pack(side="left", fill="x", expand=True)

        self.items_tree.heading("Name", text="品名及规格")
        self.items_tree.heading("Unit", text="单位")
        self.items_tree.heading("Qty", text="数量")
        self.items_tree.heading("Price", text="单价 (¥)")
        self.items_tree.heading("Subtotal", text="小计金额 (¥)")
        
        self.items_tree.column("Name", width=250)
        self.items_tree.column("Unit", width=60, anchor='center')
        self.items_tree.column("Qty", width=80, anchor='e')
        self.items_tree.column("Price", width=100, anchor='e')
        self.items_tree.column("Subtotal", width=120, anchor='e')

        # 4. 主操作按钮
        action_frame = ttk.Frame(self.master)
        action_frame.pack(padx=10, pady=10, fill="x")
        ttk.Button(action_frame, text="开具并保存收据 (生成PDF)", command=self.save_receipt).pack(pady=5)

        # 5. 历史记录查询
        query_frame = ttk.LabelFrame(self.master, text="历史收据查询 (双击行进行操作)", padding="10")
        query_frame.pack(padx=10, pady=5, fill="both", expand=True)

        self.receipt_tree = ttk.Treeview(query_frame, columns=("ID", "Date", "Customer", "Amount"), show='headings')
        self.receipt_tree.pack(fill="both", expand=True)

        self.receipt_tree.heading("ID", text="编号")
        self.receipt_tree.heading("Date", text="日期")
        self.receipt_tree.heading("Customer", text="客户名称")
        self.receipt_tree.heading("Amount", text="总金额")

        self.receipt_tree.column("ID", width=60, anchor='center')
        self.receipt_tree.column("Date", width=150, anchor='center')
        self.receipt_tree.column("Amount", width=100, anchor='e')
        self.receipt_tree.column("Customer", width=250)
        
        self.receipt_tree.bind('<Double-1>', self.on_receipt_select)
        
    def calculate_total(self):
        """计算并更新总金额"""
        total = sum(item['qty'] * item['price'] for item in self.item_list)
        self.total_amount_var.set(f"{total:.2f}")

    def add_item(self):
        """添加一项明细到列表"""
        try:
            name = self.item_vars['item_name_var'].get().strip()
            unit = self.item_vars['item_unit_var'].get().strip()
            qty = self.item_vars['item_qty_var'].get()
            price = self.item_vars['item_price_var'].get()
            
            if not name or qty <= 0 or price <= 0:
                messagebox.showerror("输入错误", "品名、数量和单价必须有效且大于零。")
                return

            item = {
                'name': name,
                'unit': unit,
                'qty': qty,
                'price': price
            }
            self.item_list.append(item)
            subtotal = qty * price
            
            # 更新显示表格
            self.items_tree.insert("", tk.END, values=(name, unit, qty, f"{price:.2f}", f"{subtotal:.2f}"))
            
            # 清空输入框
            self.item_vars['item_name_var'].set("")
            self.item_vars['item_unit_var'].set("")
            self.item_vars['item_qty_var'].set(0.0)
            self.item_vars['item_price_var'].set(0.0)

            self.calculate_total()
            
        except tk.TclError:
            messagebox.showerror("输入错误", "数量和单价请输入有效数字。")

    def clear_items(self):
        """清空当前明细列表"""
        if messagebox.askyesno("确认清空", "确定要清除所有已添加的明细项目吗？"):
            self.item_list = []
            for item in self.items_tree.get_children():
                self.items_tree.delete(item)
            self.calculate_total()
            
    def save_receipt(self):
        """处理保存收据的逻辑"""
        customer = self.customer_var.get().strip()
        total_amount = self.total_amount_var.get()
        
        if not customer:
            messagebox.showerror("错误", "客户名称不能为空。")
            return
        
        if not self.item_list:
            messagebox.showerror("错误", "请至少添加一个明细项目。")
            return
            
        receipt_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 将明细列表转换为 JSON 字符串存储
        description_json = json.dumps(self.item_list)
        
        # 1. 存入数据库
        new_id = self.db.add_receipt(receipt_date, customer, float(total_amount), description_json, "")
        
        pdf_data = {
            'id': new_id,
            'receipt_date': receipt_date,
            'customer_name': customer,
            'amount': float(total_amount),
            'description': description_json # 传入JSON字符串
        }
        
        # 2. 调用 PDF 生成函数
        pdf_filename = generate_pdf(pdf_data)
        
        if pdf_filename:
            self.db.update_pdf_path(new_id, pdf_filename)
            messagebox.showinfo("成功", f"收据编号 {new_id} 已成功保存，PDF 已生成。")
        else:
            messagebox.showwarning("部分成功", f"收据编号 {new_id} 已保存，但 PDF 生成失败。")
        
        # 清空当前界面输入
        self.customer_var.set("")
        self.clear_items()
        self.load_receipts()

    def load_receipts(self):
        """从数据库加载并显示收据列表"""
        try:
            for item in self.receipt_tree.get_children():
                self.receipt_tree.delete(item)
                
            for row in self.db.get_all_receipts():
                display_row = list(row)[:4] # 只取前4列 (ID, Date, Customer, Amount)
                display_row[3] = f"¥ {row[3]:.2f}"
                self.receipt_tree.insert("", tk.END, values=display_row)
        except sqlite3.OperationalError:
            messagebox.showerror("数据库错误", "数据库结构不匹配。请删除 'receipts.db' 文件后重试。")
            
    def on_receipt_select(self, event):
        """双击表格中的收据记录时触发的事件：弹出操作菜单"""
        # ... (PDF 查看和打印逻辑保持不变)
        selected_item = self.receipt_tree.focus()
        if not selected_item: return

        values = self.receipt_tree.item(selected_item, 'values')
        receipt_id = values[0]
        
        receipt_data = self.db.get_receipt_by_id(receipt_id)
        if not receipt_data or not receipt_data['pdf_path']:
             messagebox.showwarning("提示", f"收据编号 {receipt_id} 的 PDF 文件路径缺失，请点击保存按钮重新生成。")
             return

        pdf_path = receipt_data['pdf_path']

        menu = tk.Menu(self.master, tearoff=0)
        menu.add_command(label="查看/打开 PDF", command=lambda: self.open_pdf(pdf_path))
        menu.add_command(label="打印收据 (需手动确认)", command=lambda: self.print_pdf(pdf_path))
        
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def open_pdf(self, pdf_path):
        if not os.path.exists(pdf_path):
            messagebox.showerror("错误", "PDF 文件不存在，可能已被删除。")
            return
        if sys.platform == "win32":
            os.startfile(pdf_path)
        elif sys.platform == "darwin":
            os.system(f"open {pdf_path}")
        else:
            os.system(f"xdg-open {pdf_path}")

    def print_pdf(self, pdf_path):
        messagebox.showinfo("打印提示", "程序将打开 PDF 文件，请在 PDF 阅读器中手动点击打印按钮。")
        self.open_pdf(pdf_path)


# --- 主程序入口 ---
if __name__ == "__main__":
    if not os.path.exists("Receipts_PDFs"):
        os.makedirs("Receipts_PDFs")

    root = tk.Tk()
    try:
        app = ReceiptApp(root)
        def on_closing():
            app.db.close()
            root.destroy()
            
        root.protocol("WM_DELETE_WINDOW", on_closing)
        root.mainloop()
    except sqlite3.OperationalError as e:
        messagebox.showerror("致命错误", f"数据库初始化失败：{e}。请确认已删除项目目录下的 receipts.db 文件。")