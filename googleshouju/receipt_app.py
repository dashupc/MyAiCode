import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import datetime
from fpdf2 import FPDF # 实际依赖 fpdf2 库
import os
import sys
import webbrowser 

# --- 0. 资源路径函数 (解决打包后文件找不到的问题) ---
def resource_path(relative_path):
    """获取资源文件的绝对路径，以兼容 PyInstaller 单文件模式"""
    try:
        # PyInstaller 在运行时会将所有 --add-data 的文件解压到 _MEIPASS 目录
        base_path = sys._MEIPASS
    except Exception:
        # 如果不是 PyInstaller 运行，使用当前工作目录
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# --- 1. 数据库操作类 ---
class DatabaseManager:
    """管理 SQLite 数据库连接和操作，包括收据数据和用户设置（收款人信息）"""
    def __init__(self, db_name="receipts.db"):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self._setup_db()

    def _setup_db(self):
        """初始化数据库表结构：Receipts（主表）、ReceiptItems（明细表）、Settings（设置表）"""
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS Receipts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                receipt_no TEXT UNIQUE NOT NULL,
                client_name TEXT NOT NULL,
                issue_date TEXT NOT NULL,
                total_amount_num REAL NOT NULL,
                total_amount_cap TEXT NOT NULL
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS ReceiptItems (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                receipt_id INTEGER,
                item_name TEXT,
                unit TEXT,
                quantity REAL,
                unit_price REAL,
                amount REAL,
                notes TEXT,
                FOREIGN KEY (receipt_id) REFERENCES Receipts(id)
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS Settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        self.conn.commit()

    def generate_receipt_no(self):
        """生成收据编号: 年月日 + 3位流水号"""
        date_str = datetime.datetime.now().strftime("%Y%m%d")
        self.cursor.execute("SELECT COUNT(*) FROM Receipts WHERE receipt_no LIKE ?", (date_str + '%',))
        count = self.cursor.fetchone()[0] + 1
        return f"{date_str}{count:03d}"

    def save_receipt(self, receipt_data, items_data):
        """保存一张完整的收据及其明细"""
        receipt_no = self.generate_receipt_no()
        issue_date = datetime.datetime.now().strftime("%Y-%m-%d")
        
        try:
            self.cursor.execute(
                "INSERT INTO Receipts (receipt_no, client_name, issue_date, total_amount_num, total_amount_cap) VALUES (?, ?, ?, ?, ?)",
                (receipt_no, receipt_data['client_name'], issue_date, receipt_data['total_amount_num'], receipt_data['total_amount_cap'])
            )
            receipt_id = self.cursor.lastrowid
            
            for item in items_data:
                self.cursor.execute(
                    "INSERT INTO ReceiptItems (receipt_id, item_name, unit, quantity, unit_price, amount, notes) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (receipt_id, item['item_name'], item['unit'], item['quantity'], item['unit_price'], item['amount'], item['notes'])
                )
            
            self.conn.commit()
            return receipt_no
        except Exception as e:
            messagebox.showerror("错误", f"保存收据失败: {e}")
            return None

    def save_setting(self, key, value):
        """保存设置（如收款人信息）"""
        self.cursor.execute(
            "INSERT OR REPLACE INTO Settings (key, value) VALUES (?, ?)",
            (key, value)
        )
        self.conn.commit()

    def load_setting(self, key, default=""):
        """加载设置（如收款人信息）"""
        self.cursor.execute("SELECT value FROM Settings WHERE key = ?", (key,))
        result = self.cursor.fetchone()
        return result[0] if result else default

    def fetch_all_receipts(self, query=""):
        """查询所有收据或根据客户名称模糊查询"""
        if query:
            self.cursor.execute(
                "SELECT receipt_no, client_name, issue_date, total_amount_num FROM Receipts WHERE client_name LIKE ? ORDER BY id DESC",
                ('%' + query + '%',)
            )
        else:
            self.cursor.execute("SELECT receipt_no, client_name, issue_date, total_amount_num FROM Receipts ORDER BY id DESC")
        return self.cursor.fetchall()

    def fetch_receipt_details(self, receipt_no):
        """根据收据编号获取详情，并附加最新的收款人设置信息"""
        self.cursor.execute("SELECT id, client_name, issue_date, total_amount_num, total_amount_cap FROM Receipts WHERE receipt_no = ?", (receipt_no,))
        receipt = self.cursor.fetchone()
        
        if not receipt:
            return None, None
            
        receipt_id, client_name, issue_date, total_amount_num, total_amount_cap = receipt
        
        self.cursor.execute(
            "SELECT item_name, unit, quantity, unit_price, amount, notes FROM ReceiptItems WHERE receipt_id = ?",
            (receipt_id,)
        )
        items = self.cursor.fetchall()
        
        # 附加最新的收款人设置信息，用于显示或PDF生成
        payee = self.load_setting("payee", "")
        payee_company = self.load_setting("payee_company", "")
        
        return {
            'receipt_no': receipt_no,
            'client_name': client_name,
            'issue_date': issue_date,
            'total_amount_num': total_amount_num,
            'total_amount_cap': total_amount_cap,
            'payee': payee,
            'payee_company': payee_company
        }, items

    def close(self):
        self.conn.close()

# --- 2. 核心逻辑函数 ---
def convert_to_chinese_caps(num):
    """将数字金额转换为中文大写（简化实现）"""
    num = abs(num) 
    num_str = f"{int(num):,}"
    
    CAPS = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"]
    UNITS = ["", "拾", "佰", "仟"]
    GROUPS = ["", "万", "亿"]
    
    integer_part = str(int(num))
    
    result = ""
    group_index = 0
    
    while integer_part:
        chunk = integer_part[-4:]
        integer_part = integer_part[:-4]
        
        chunk_result = ""
        for i, char in enumerate(reversed(chunk)):
            digit = int(char)
            if digit != 0:
                chunk_result = CAPS[digit] + UNITS[i] + chunk_result
            elif chunk_result and not chunk_result.startswith("零"):
                chunk_result = CAPS[digit] + chunk_result
        
        if chunk_result:
            result = chunk_result.rstrip("零") + GROUPS[group_index] + result
        
        group_index += 1

    result = result.replace("零万", "万").replace("零亿", "亿").strip("零")
    
    if not result:
        return "零元整"
        
    return result + "元整"

# --- 3. PDF 生成类 (包含绘制公章逻辑) ---
class PDFGenerator:
    """使用 fpdf2 生成收据 PDF，支持绘制圆形公章"""
    def __init__(self, receipt_data, items_data):
        self.data = receipt_data
        self.items = items_data
        self.pdf = FPDF('P', 'mm', 'A4')
        
        font_path = resource_path('simhei.ttf') 
        self.pdf_font_name = 'SimHei' 
        
        try:
             self.pdf.add_font(self.pdf_font_name, '', font_path, uni=True) 
             self.pdf.set_font(self.pdf_font_name, '', 12)
        except FileNotFoundError:
             messagebox.showwarning("字体缺失", "未找到 simhei.ttf，PDF将无法正确显示中文，请将字体文件放在程序目录下。")
             self.pdf.set_font('Arial', '', 12) 
             self.pdf_font_name = 'Arial' 

        # ⭐ 公章图片路径和尺寸定义
        self.star_image_path = resource_path('star.png') 
        self.seal_outer_diameter = 40 # 外圈直径 40mm
        self.seal_outer_radius = self.seal_outer_diameter / 2
        self.star_diameter = 10 # 五角星直径 10mm
        
    def generate(self):
        self.pdf.add_page()
        
        # 标题和编号
        self.pdf.set_font(self.pdf_font_name, '', 18)
        self.pdf.cell(0, 10, '收 款 收 据', 0, 1, 'C')
        
        self.pdf.set_font(self.pdf_font_name, '', 12)
        self.pdf.cell(0, 8, f"NO.{self.data['receipt_no']}", 0, 1, 'R')

        # 客户信息
        self.pdf.cell(0, 8, f"客户名称: {self.data['client_name']}", 0, 1, 'L')
        self.pdf.ln(2)

        # 表格头
        col_widths = [40, 20, 20, 20, 30, 60] # mm
        col_names = ['项目名称', '单位', '数量', '单价', '金额', '备注']
        self.pdf.set_fill_color(200, 220, 255)
        self.pdf.set_font(self.pdf_font_name, '', 10)
        
        for i, header in enumerate(col_names):
            self.pdf.cell(col_widths[i], 8, header, 1, 0, 'C', 1)
        self.pdf.ln()

        # 表格数据
        self.pdf.set_font(self.pdf_font_name, '', 10)
        for item in self.items:
            # item 格式: (item_name, unit, quantity, unit_price, amount, notes)
            row = [
                item[0] or '', item[1] or '', 
                f"{item[2]:.2f}", f"{item[3]:.2f}", 
                f"{item[4]:.2f}", item[5] or ''
            ]
            for i, text in enumerate(row):
                align = 'L' if i in [0, 5] else 'R'
                self.pdf.cell(col_widths[i], 8, str(text), 1, 0, align)
            self.pdf.ln()
            
        # 填充空白行
        if len(self.items) < 3:
            for _ in range(3 - len(self.items)):
                for width in col_widths:
                     self.pdf.cell(width, 8, '', 1, 0, 'C')
                self.pdf.ln()

        # 合计行
        total_text = f"合计: 人民币(大写) {self.data['total_amount_cap']}"
        total_num_text = f"合计: {self.data['total_amount_num']:.2f}元"
        
        self.pdf.ln(5)
        self.pdf.set_font(self.pdf_font_name, '', 12)
        
        self.pdf.cell(100, 8, total_text, 0, 0, 'L')
        self.pdf.cell(90, 8, total_num_text, 0, 1, 'R')
        
        self.pdf.ln(5)
        
        # 收款人/收款单位 信息
        payee = self.data.get('payee', '')
        payee_company = self.data.get('payee_company', '')
        issue_date = self.data['issue_date']

        self.pdf.set_font(self.pdf_font_name, '', 10)
        
        # 记录 Y 坐标，用于后续定位公章
        start_y_for_payee_company = self.pdf.get_y()
        
        self.pdf.cell(90, 8, f"收款单位: {payee_company}", 0, 0, 'L')
        self.pdf.cell(90, 8, f"收款人: {payee}", 0, 1, 'R')
        
        # 确定公章的中心位置
        # 公章中心 X 坐标：在页面右侧约 40mm 处（假设右边距为 10mm, 页面宽 210mm）
        seal_center_x = self.pdf.w - self.pdf.r_margin - 40 # 210 - 10 - 40 = 160mm 左右
        # 公章中心 Y 坐标：在收款单位行下方 20mm 处
        seal_center_y = start_y_for_payee_company + 20 
        
        # --- 绘制公章 ---
        self.pdf.save_state()
        self.pdf.set_draw_color(255, 0, 0) # 红色边框
        self.pdf.set_text_color(255, 0, 0) # 红色文字

        # 绘制外圈圆 (直径40mm)
        self.pdf.circle(seal_center_x, seal_center_y, self.seal_outer_radius)

        # 插入五角星图片 (直径10mm)
        if os.path.exists(self.star_image_path):
            try:
                star_x = seal_center_x - self.star_diameter / 2
                star_y = seal_center_y - self.star_diameter / 2
                self.pdf.image(self.star_image_path, x=star_x, y=star_y, 
                               w=self.star_diameter, h=self.star_diameter, type='PNG')
            except Exception as e:
                messagebox.showwarning("五角星图片绘制失败", f"无法在PDF中绘制五角星: {e}")
        
        # 放置“收款单位”文字 (水平放置在上方)
        self.pdf.set_font(self.pdf_font_name, '', 10) 
        text_width_company = self.pdf.get_string_width(payee_company)
        text_x_company = seal_center_x - text_width_company / 2
        # Y 坐标：中心 Y 减去半径，再向下偏移 6mm
        text_y_company = seal_center_y - self.seal_outer_radius + 6 
        
        self.pdf.set_xy(text_x_company, text_y_company)
        self.pdf.cell(text_width_company, 5, payee_company, 0, 0, 'C')

        # 放置日期 (在公章下方)
        self.pdf.set_font(self.pdf_font_name, '', 8) 
        text_width_date = self.pdf.get_string_width(issue_date)
        text_x_date = seal_center_x - text_width_date / 2
        # Y 坐标：中心 Y 加上半径，再向上偏移 8mm
        text_y_date = seal_center_y + self.seal_outer_radius - 8 
        
        self.pdf.set_xy(text_x_date, text_y_date)
        self.pdf.cell(text_width_date, 5, issue_date, 0, 0, 'C')
        
        self.pdf.restore_state()
        # --- 公章绘制结束 ---

        filename = f"Receipt_{self.data['receipt_no']}.pdf"
        self.pdf.output(filename)
        return filename

# --- 4. 主 GUI 应用类 ---
class ReceiptApp:
    def __init__(self, master):
        self.master = master
        
        # 1. 设置窗口图标 (使用 resource_path 确保打包后生效)
        icon_path = resource_path('receipt.ico')
        try:
            self.master.iconbitmap(icon_path)
        except tk.TclError:
            pass 

        master.title("收款收据管理系统")
        
        WINDOW_WIDTH = 850
        WINDOW_HEIGHT = 700
        
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        
        center_x = int((screen_width / 2) - (WINDOW_WIDTH / 2))
        center_y = int((screen_height / 2) - (WINDOW_HEIGHT / 2))
        
        master.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{center_x}+{center_y}")
        
        self.db = DatabaseManager()
        self.entry_editor = None 
        
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(pady=10, padx=10, expand=True, fill="both")

        self.editable_cols = [0, 1, 2, 3, 5] 
        
        # 创建各个标签页
        self.create_receipt_tab()
        self.query_receipt_tab()
        self.create_about_tab() 
        
        self.refresh_query_results()

        self.client_name_entry.focus()
        self._load_default_payee_info() 
        
    def create_about_tab(self):
        """创建关于软件的标签页"""
        about_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(about_frame, text="关于")

        ttk.Label(about_frame, text="电子收据管理系统", font=('Arial', 18, 'bold')).pack(pady=20)
        
        description_text = (
            "本软件旨在提供一个简单、高效的收据开具和管理解决方案。 "
            "主要功能包括快速录入项目明细、自动计算总额、生成标准 PDF 收据文件，"
            "以及方便的历史收据查询和打印功能。"
        )
        ttk.Label(about_frame, text=description_text, wraplength=600, justify='left', font=('Arial', 10)).pack(pady=10, padx=50)

        contact_frame = ttk.LabelFrame(about_frame, text="联系方式", padding="15")
        contact_frame.pack(pady=20, padx=50, fill='x')

        # QQ
        ttk.Label(contact_frame, text="QQ:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        ttk.Label(contact_frame, text="88179096", font=('Arial', 10)).grid(row=0, column=1, sticky='w', padx=10, pady=5)

        # 网址 (带超链接)
        ttk.Label(contact_frame, text="网址:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=5, pady=5)
        
        # 创建超链接标签
        url_label = ttk.Label(contact_frame, text="www.itvip.com.cn", 
                              foreground="blue", cursor="hand2", font=('Arial', 10, 'underline'))
        url_label.grid(row=1, column=1, sticky='w', padx=10, pady=5)
        url_label.bind("<Button-1>", lambda e: self.open_url("http://www.itvip.com.cn"))

    def open_url(self, url):
        """在新浏览器窗口中打开指定的 URL"""
        try:
            webbrowser.open_new(url)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开网址: {e}")

    def create_receipt_tab(self):
        self.receipt_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.receipt_frame, text="开具收据")

        # 客户信息部分
        client_frame = ttk.LabelFrame(self.receipt_frame, text="收据信息", padding="10")
        client_frame.pack(fill="x", pady=5)
        
        ttk.Label(client_frame, text="客户名称:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.client_name_entry = ttk.Entry(client_frame, width=50)
        self.client_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.client_name_entry.bind('<Tab>', self.jump_to_first_item_row)
        
        ttk.Label(client_frame, text="NO.XXXXXX (自动生成)").grid(row=0, column=2, padx=20, pady=5, sticky="e")

        # 项目明细表格
        items_frame = ttk.LabelFrame(self.receipt_frame, text="项目明细 (双击或按Tab键进行输入/跳转)", padding="10")
        items_frame.pack(fill="both", expand=True, pady=10)
        
        columns = ("item_name", "unit", "quantity", "unit_price", "amount", "notes")
        self.items_tree = ttk.Treeview(items_frame, columns=columns, show="headings", height=10)
        
        headings = {"item_name": "项目名称", "unit": "单位", "quantity": "数量", "unit_price": "单价", "amount": "金额", "notes": "备注"}
        widths = [150, 60, 90, 90, 90, 180]
        
        for col, width in zip(columns, widths):
            self.items_tree.heading(col, text=headings[col])
            self.items_tree.column(col, width=width, anchor='center')
        
        self.items_tree.pack(fill="both", expand=True)
        
        self.items_tree.bind('<Button-1>', self.on_tree_click)
        self.items_tree.bind('<Return>', lambda e: self.on_tree_click(e, trigger_key='Return')) 

        # 表格操作按钮
        btn_frame = ttk.Frame(items_frame)
        btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="添加项目行", command=self.add_item_row).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="删除选中项目", command=self.delete_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="计算合计", command=self.calculate_total).pack(side="left", padx=5)
        
        # 合计信息
        total_frame = ttk.Frame(self.receipt_frame)
        total_frame.pack(fill="x", pady=5)
        
        self.total_cap_var = tk.StringVar(value="合计: 人民币(大写) 零元整")
        self.total_num_var = tk.StringVar(value="合计: 0.00元")
        
        ttk.Label(total_frame, textvariable=self.total_cap_var).pack(side="left", padx=10)
        ttk.Label(total_frame, textvariable=self.total_num_var, font=('Arial', 12, 'bold')).pack(side="right", padx=10)
        
        # 收款人/收款单位 输入框
        payee_frame = ttk.Frame(self.receipt_frame)
        payee_frame.pack(fill="x", pady=10)
        
        ttk.Label(payee_frame, text="收款单位:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.payee_company_entry = ttk.Entry(payee_frame, width=30)
        self.payee_company_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        payee_frame.grid_columnconfigure(2, weight=1) 
        
        ttk.Label(payee_frame, text="收款人:").grid(row=0, column=3, padx=5, pady=5, sticky="w")
        self.payee_entry = ttk.Entry(payee_frame, width=20)
        self.payee_entry.grid(row=0, column=4, padx=5, pady=5, sticky="w")
        
        # 主要功能按钮
        main_btn_frame = ttk.Frame(self.receipt_frame)
        main_btn_frame.pack(fill="x", pady=10)
        ttk.Button(main_btn_frame, text="保存收据并生成PDF", command=lambda: self.save_and_generate(True)).pack(side="right", padx=10)
        ttk.Button(main_btn_frame, text="保存收据", command=lambda: self.save_and_generate(False)).pack(side="right", padx=10)
        ttk.Button(main_btn_frame, text="清空", command=self.clear_receipt_form).pack(side="left", padx=10)

    def _load_default_payee_info(self):
        """从数据库加载并设置默认的收款人信息"""
        payee = self.db.load_setting("payee", "")
        payee_company = self.db.load_setting("payee_company", "")
        
        self.payee_entry.delete(0, tk.END)
        self.payee_entry.insert(0, payee)
        
        self.payee_company_entry.delete(0, tk.END)
        self.payee_company_entry.insert(0, payee_company)

    def jump_to_first_item_row(self, event):
        """Tab 键跳转到明细表格"""
        self.master.after(10, self._perform_jump)
        return "break"
        
    def _perform_jump(self):
        children = self.items_tree.get_children()
        if not children:
            self.add_item_row()
        else:
            first_item_id = children[0]
            self.items_tree.focus(first_item_id)
            self.items_tree.selection_set(first_item_id)
            self.on_tree_click(None, item_id=first_item_id, column_id='#1')

    def query_receipt_tab(self):
        self.query_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.query_frame, text="收据查询")

        # 查询条件
        search_frame = ttk.Frame(self.query_frame)
        search_frame.pack(fill="x", pady=5)
        ttk.Label(search_frame, text="客户名称/编号查询:").pack(side="left", padx=5)
        self.search_entry = ttk.Entry(search_frame, width=40)
        self.search_entry.pack(side="left", padx=5)
        ttk.Button(search_frame, text="查询", command=self.refresh_query_results).pack(side="left", padx=10)

        # 查询结果 Treeview
        query_result_frame = ttk.Frame(self.query_frame)
        query_result_frame.pack(fill="both", expand=True, pady=10)
        
        self.query_tree = ttk.Treeview(query_result_frame, columns=("no", "name", "date", "amount"), show="headings", height=15)
        
        # 修正 NameError
        columns = ("no", "name", "date", "amount") 
        
        headings = {"no": "收据编号", "name": "客户名称", "date": "开具日期", "amount": "合计金额"}
        widths = [150, 200, 100, 100]
        
        for col, width in zip(columns, widths):
            self.query_tree.heading(col, text=headings[col])
            self.query_tree.column(col, width=width, anchor='center')
        
        self.query_tree.pack(fill="both", expand=True)
        
        # 绑定双击事件，用于显示详情
        self.query_tree.bind('<Double-1>', self.show_receipt_details)

        # 查询操作按钮
        query_btn_frame = ttk.Frame(self.query_frame)
        query_btn_frame.pack(fill="x", pady=10)
        ttk.Button(query_btn_frame, text="查看详情/生成PDF", command=self.view_details_pdf).pack(side="left", padx=10)
        ttk.Button(query_btn_frame, text="打印收据", command=self.print_receipt).pack(side="left", padx=10)

    def calculate_total(self):
        """计算项目明细的总金额"""
        if self.entry_editor:
             self.entry_editor.destroy()
             self.entry_editor = None

        total = 0.0
        columns = self.items_tree["columns"]
        col_map = {col: i for i, col in enumerate(columns)}
        
        for item_id in self.items_tree.get_children():
            values = list(self.items_tree.item(item_id, 'values'))
            
            try:
                quantity = float(values[col_map['quantity']] or 0) 
                unit_price = float(values[col_map['unit_price']] or 0)
                
                amount = round(quantity * unit_price, 2)
                
                values[col_map['amount']] = f"{amount:.2f}"
                self.items_tree.item(item_id, values=values)
                
                total += amount
            except (ValueError, IndexError):
                pass
        
        self.total_num_var.set(f"合计: {total:.2f}元")
        self.total_cap_var.set(f"合计: 人民币(大写) {convert_to_chinese_caps(total)}")
        return total

    def add_item_row(self):
        """添加一个空的明细行"""
        empty_values = ("", "", "", "", "0.00", "") 
        new_item_id = self.items_tree.insert("", "end", values=empty_values)
        self.items_tree.focus(new_item_id)
        self.items_tree.selection_set(new_item_id)
        self.on_tree_click(None, item_id=new_item_id, column_id='#1') 
        
    def delete_item(self):
        """删除选中的项目明细行"""
        selected_item = self.items_tree.selection()
        if selected_item:
            self.items_tree.delete(selected_item[0])
            self.calculate_total()
        else:
            messagebox.showinfo("提示", "请选择要删除的项目。")

    def on_tree_click(self, event, item_id=None, column_id=None, trigger_key=None):
        """处理 Treeview 的点击事件，实现行内编辑"""
        if self.entry_editor:
             self.entry_editor.destroy()
             self.entry_editor = None

        # 确定被点击的行和列
        if event and trigger_key != 'Return':
             region = self.items_tree.identify('region', event.x, event.y)
             if region != 'cell': return
             
             item = self.items_tree.identify_row(event.y)
             column = self.items_tree.identify_column(event.x)
             
        elif item_id and column_id:
             item = item_id
             column = column_id
        else:
             item = self.items_tree.focus()
             if event:
                 try:
                     column = self.items_tree.identify_column(event.x)
                 except:
                     column = '#1'
             else:
                 column = '#1'

        col_index = int(column.replace('#', '')) - 1
        
        if not item or col_index not in self.editable_cols: 
            return

        x, y, w, h = self.items_tree.bbox(item, column)
        current_value = self.items_tree.set(item, column)
        
        self.entry_editor = ttk.Entry(self.items_tree)
        self.entry_editor.place(x=x, y=y, width=w, height=h)
        self.entry_editor.insert(0, current_value)
        self.entry_editor.focus()
        
        # 绑定 Tab/Enter/失焦事件
        self.entry_editor.bind('<Tab>', 
                               lambda e: self.on_editor_save(item, column, col_index, e, is_tab=True))
        self.entry_editor.bind('<Return>', 
                               lambda e: self.on_editor_save(item, column, col_index, e, is_tab=False))
        self.entry_editor.bind('<FocusOut>', 
                                lambda e: self.on_editor_save(item, column, col_index, e, is_tab=False))

    def on_editor_save(self, item, column, col_index, event, is_tab):
        """保存 Entry 编辑器的值"""
        if not self.entry_editor:
            return

        new_value = self.entry_editor.get()
        
        if col_index in [2, 3]: # 数量(2), 单价(3)
            try:
                float(new_value or 0)
            except ValueError:
                messagebox.showerror("输入错误", "数量和单价必须是有效的数字！")
                self.entry_editor.focus() 
                return

        current_values = list(self.items_tree.item(item, 'values'))
        while len(current_values) <= col_index:
             current_values.append("")
        current_values[col_index] = new_value
        self.items_tree.item(item, values=current_values)

        self.entry_editor.destroy()
        self.entry_editor = None
        self.calculate_total()
        
        if is_tab:
            self.master.after(10, lambda: self.handle_tab_jump(item, col_index))
            return "break" 

    def handle_tab_jump(self, current_item, current_col_index):
        """处理 Tab 键按下后的焦点跳转逻辑"""
        editable_cols = self.editable_cols 
        
        try:
            current_editable_index = editable_cols.index(current_col_index)
            
            if current_editable_index < len(editable_cols) - 1:
                next_col_index = editable_cols[current_editable_index + 1]
                next_column_id = f"#{next_col_index + 1}"
                next_item = current_item
                
            else:
                next_item = self.items_tree.next(current_item)
                
                if not next_item:
                    self.add_item_row()
                    return 
                
                next_col_index = editable_cols[0] 
                next_column_id = f"#{next_col_index + 1}"

            self.items_tree.focus(next_item)
            self.items_tree.selection_set(next_item)
            self.on_tree_click(None, item_id=next_item, column_id=next_column_id)
            
        except ValueError:
            pass
        
    def save_and_generate(self, generate_pdf=False):
        """保存收据到数据库并可选地生成 PDF"""
        self.calculate_total() 
        client_name = self.client_name_entry.get().strip()
        
        if not client_name:
            messagebox.showwarning("警告", "客户名称不能为空！")
            return
            
        items_raw = [self.items_tree.item(i, 'values') for i in self.items_tree.get_children()]
        items_raw = [item for item in items_raw if item and any(item)]
        
        if not items_raw:
            messagebox.showwarning("警告", "项目明细不能为空！")
            return

        total_num = self.calculate_total()
        total_cap = convert_to_chinese_caps(total_num)

        receipt_data = {
            'client_name': client_name,
            'total_amount_num': total_num,
            'total_amount_cap': total_cap
        }
        
        items_data = []
        for item in items_raw:
            try:
                items_data.append({
                    'item_name': item[0] or "",
                    'unit': item[1] or "",
                    'quantity': float(item[2] or 0), 
                    'unit_price': float(item[3] or 0),
                    'amount': float(item[4] or 0),
                    'notes': item[5] or ""
                })
            except (ValueError, IndexError) as e:
                messagebox.showerror("数据错误", f"项目明细中存在无效数据，请检查数量和单价列。错误: {e}")
                return

        receipt_no = self.db.save_receipt(receipt_data, items_data)
        
        if receipt_no:
            # 保存当前的收款人信息作为默认值
            payee_current = self.payee_entry.get().strip()
            payee_company_current = self.payee_company_entry.get().strip()
            self.db.save_setting("payee", payee_current)
            self.db.save_setting("payee_company", payee_company_current)
            
            messagebox.showinfo("成功", f"收据 {receipt_no} 已成功保存到数据库。")
            self.clear_receipt_form()
            self.refresh_query_results()
            
            if generate_pdf:
                # 重新加载数据以确保 PDF 中包含最新的收款人信息
                pdf_data, pdf_items = self.db.fetch_receipt_details(receipt_no)
                
                # 确保 pdf_data 包含当前收款人/单位信息 (如果数据库中未存，则从当前界面获取)
                if 'payee' not in pdf_data:
                     pdf_data['payee'] = payee_current
                     pdf_data['payee_company'] = payee_company_current

                if pdf_data and pdf_items:
                    generator = PDFGenerator(pdf_data, pdf_items)
                    pdf_filename = generator.generate()
                    # 直接打开文件，而不是尝试调用系统打印命令
                    try:
                        if os.name == 'nt': 
                            os.startfile(pdf_filename) 
                        elif os.name == 'posix': 
                            os.system(f"open {pdf_filename}") 
                        messagebox.showinfo("PDF 生成", f"PDF 文件已生成并打开: {pdf_filename}")
                    except Exception as e:
                        messagebox.showwarning("打开失败", f"PDF 文件已生成: {pdf_filename}，但无法自动打开。请手动打开打印。错误: {e}")


    def clear_receipt_form(self):
        """清空开具收据界面 (保留收款人信息)"""
        self.client_name_entry.delete(0, tk.END)
        for item in self.items_tree.get_children():
            self.items_tree.delete(item)
        self.calculate_total() 
        self._load_default_payee_info() # 重新加载默认的收款人信息
        
    def refresh_query_results(self):
        """根据查询条件刷新查询结果列表"""
        query = self.search_entry.get().strip()
        
        for item in self.query_tree.get_children():
            self.query_tree.delete(item)
            
        results = self.db.fetch_all_receipts(query)
        
        for row in results:
            receipt_no, client_name, issue_date, total_amount_num = row
            amount_str = f"{total_amount_num:.2f}"
            self.query_tree.insert("", "end", values=(receipt_no, client_name, issue_date, amount_str))

    def view_details_pdf(self):
        """生成选中收据的PDF文件"""
        selected_item = self.query_tree.focus()
        if not selected_item:
            messagebox.showwarning("警告", "请先在列表中选择一张收据。")
            return
            
        receipt_no = self.query_tree.item(selected_item, 'values')[0]
        receipt_data, items_data = self.db.fetch_receipt_details(receipt_no)
        
        if receipt_data and items_data:
            generator = PDFGenerator(receipt_data, items_data)
            pdf_filename = generator.generate()
            messagebox.showinfo("PDF 生成", f"收据 {receipt_no} 的 PDF 文件已重新生成: {pdf_filename}")
        else:
            messagebox.showerror("错误", f"未找到收据编号 {receipt_no} 的详细信息。")
            
    def print_receipt(self):
        """生成选中收据的PDF并尝试打开文件（从查询页调用）"""
        selected_item = self.query_tree.focus()
        if not selected_item:
            messagebox.showwarning("警告", "请先在列表中选择一张收据进行打印。")
            return
            
        receipt_no = self.query_tree.item(selected_item, 'values')[0]
        # 直接调用内部打印方法
        self._print_receipt_from_details(receipt_no, self.master)
        
    def _print_receipt_from_details(self, receipt_no, parent_window):
        """
        根据指定的收据编号生成PDF并尝试调用系统打开文件，供用户手动打印。
        """
        # 1. 获取数据
        receipt_data, items_data = self.db.fetch_receipt_details(receipt_no)
        if not receipt_data or not items_data:
            messagebox.showerror("错误", f"无法获取收据 {receipt_no} 的详情。", parent=parent_window)
            return
            
        # 2. 生成 PDF
        generator = PDFGenerator(receipt_data, items_data)
        pdf_filename = generator.generate()
        
        # 3. 尝试调用系统打开文件
        try:
            if os.name == 'nt': # Windows 系统
                os.startfile(pdf_filename) 
            elif os.name == 'posix': # Linux/Mac 系统
                # 尝试 Mac/Linux 的打开命令
                os.system(f"open {pdf_filename}") or os.system(f"xdg-open {pdf_filename}")
            else:
                messagebox.showwarning("打印提示", f"已生成 PDF: {pdf_filename}，但无法自动打开文件，请手动打开打印。", parent=parent_window)
                return
            
            messagebox.showinfo("打开成功", f"收据 PDF 文件已生成并打开: {pdf_filename}，请在 PDF 阅读器中手动打印。", parent=parent_window)
            
        except Exception as e:
            messagebox.showerror("文件关联错误", f"已生成 PDF: {pdf_filename}，但系统无法自动打开此文件。请手动找到此文件并用 PDF 阅读器打开打印。\n错误详情: {e}", parent=parent_window)


    def show_receipt_details(self, event):
        """双击查询表格中的收据，弹出详情窗口"""
        selected_item = self.query_tree.focus()
        if not selected_item: return

        receipt_no = self.query_tree.item(selected_item, 'values')[0]
        receipt_data, items_data = self.db.fetch_receipt_details(receipt_no)
        
        if not receipt_data:
            messagebox.showerror("错误", f"未找到收据编号 {receipt_no} 的详细信息。")
            return

        details_window = tk.Toplevel(self.master)
        details_window.title(f"收据详情 - NO.{receipt_no}")
        
        # --- 解决居中问题 ---
        DETAIL_WINDOW_WIDTH = 750 
        DETAIL_WINDOW_HEIGHT = 600
        
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        
        center_x = int((screen_width / 2) - (DETAIL_WINDOW_WIDTH / 2))
        center_y = int((screen_height / 2) - (DETAIL_WINDOW_HEIGHT / 2))
        
        details_window.geometry(f"{DETAIL_WINDOW_WIDTH}x{DETAIL_WINDOW_HEIGHT}+{center_x}+{center_y}")
        
        # --- 解决图标问题 ---
        icon_path = resource_path('receipt.ico')
        try:
            details_window.iconbitmap(icon_path)
        except tk.TclError:
            pass 
        
        details_window.grab_set() 

        main_frame = ttk.Frame(details_window, padding="15")
        main_frame.pack(fill="both", expand=True)

        # 头部信息
        header_frame = ttk.LabelFrame(main_frame, text="收据摘要", padding="10")
        header_frame.pack(fill="x", pady=5)
        
        ttk.Label(header_frame, text="收据编号:", font=('Arial', 10)).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Label(header_frame, text=receipt_data['receipt_no'], font=('Arial', 10, 'bold')).grid(row=0, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(header_frame, text="开具日期:", font=('Arial', 10)).grid(row=0, column=2, padx=20, pady=2, sticky="w")
        ttk.Label(header_frame, text=receipt_data['issue_date'], font=('Arial', 10)).grid(row=0, column=3, padx=5, pady=2, sticky="w")

        ttk.Label(header_frame, text="客户名称:", font=('Arial', 10)).grid(row=1, column=0, padx=5, pady=2, sticky="w")
        ttk.Label(header_frame, text=receipt_data['client_name'], font=('Arial', 10, 'bold')).grid(row=1, column=1, columnspan=3, padx=5, pady=2, sticky="w")

        # 项目明细表格
        items_frame = ttk.LabelFrame(main_frame, text="项目明细", padding="10")
        items_frame.pack(fill="both", expand=True, pady=10)
        
        columns = ("item_name", "unit", "quantity", "unit_price", "amount", "notes")
        detail_tree = ttk.Treeview(items_frame, columns=columns, show="headings", height=10)
        
        headings = {"item_name": "项目名称", "unit": "单位", "quantity": "数量", "unit_price": "单价", "amount": "金额", "notes": "备注"}
        widths = [120, 50, 70, 70, 70, 150]
        
        for col, width in zip(columns, widths):
            detail_tree.heading(col, text=headings[col])
            detail_tree.column(col, width=width, anchor='center')
        
        for item in items_data:
            detail_tree.insert("", "end", values=(
                item[0], item[1], 
                f"{item[2]:.2f}", f"{item[3]:.2f}", 
                f"{item[4]:.2f}", item[5]
            ))
            
        detail_tree.pack(fill="both", expand=True)

        # 底部合计信息
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill="x", pady=5)

        ttk.Label(footer_frame, 
                  text=f"合计 (大写): {receipt_data['total_amount_cap']}",
                  font=('Arial', 11)).pack(side="left", padx=5)
        
        ttk.Label(footer_frame, 
                  text=f"合计 (数字): ¥{receipt_data['total_amount_num']:.2f}", 
                  font=('Arial', 11, 'bold')).pack(side="right", padx=5)

        # 详情窗口中的收款人信息
        payee_details_frame = ttk.Frame(main_frame)
        payee_details_frame.pack(fill="x", pady=5)
        
        ttk.Label(payee_details_frame, text="收款单位:", font=('Arial', 10)).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Label(payee_details_frame, text=receipt_data.get('payee_company', ''), font=('Arial', 10)).grid(row=0, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(payee_details_frame, text="收款人:", font=('Arial', 10)).grid(row=0, column=2, padx=20, pady=2, sticky="w")
        ttk.Label(payee_details_frame, text=receipt_data.get('payee', ''), font=('Arial', 10)).grid(row=0, column=3, padx=5, pady=2, sticky="w")
        payee_details_frame.grid_columnconfigure(4, weight=1)
        
        # 打印按钮
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill="x", pady=10)
        
        # 绑定到内部打印方法，并传入收据编号和当前窗口
        ttk.Button(action_frame, text="打印收据", 
                   command=lambda: self._print_receipt_from_details(receipt_no, details_window)).pack(side="right")
        
        details_window.protocol("WM_DELETE_WINDOW", lambda: details_window.destroy())


# --- 5. 运行程序 ---
if __name__ == "__main__":
    if 'webbrowser' not in sys.modules:
        messagebox.showerror("依赖缺失", "未找到必要的 'webbrowser' 模块。")
        sys.exit(1)
        
    try:
        root = tk.Tk()
        app = ReceiptApp(root)
        root.mainloop()
    finally:
        if 'app' in locals():
            app.db.close()