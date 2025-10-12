import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import webbrowser 

# 从其他文件导入模块 (这些导入必须在 main_app.py 中才能正常工作)
# 确保 database_manager.py, pdf_generator.py, utils.py 存在
from database_manager import DatabaseManager
from pdf_generator import PDFGenerator
from utils import resource_path, convert_to_chinese_caps

# --- 主 GUI 应用类 ---
class ReceiptApp:
    def __init__(self, master):
        self.master = master
        
        # 1. 设置窗口图标 (主窗口)
        icon_path = resource_path('receipt.ico')
        try:
            self.master.iconbitmap(icon_path)
        except tk.TclError:
            pass 

        master.title("电子收据管理系统")
        
        WINDOW_WIDTH = 850
        WINDOW_HEIGHT = 700
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        center_x = int((screen_width / 2) - (WINDOW_WIDTH / 2))
        center_y = int((screen_height / 2) - (WINDOW_HEIGHT / 2))
        master.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{center_x}+{center_y}")
        
        self.db = DatabaseManager()
        self.entry_editor = None 
        # 允许编辑的列索引：项目名称(0), 单位(1), 数量(2), 单价(3), 备注(5)
        self.editable_cols = [0, 1, 2, 3, 5] 
        
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(pady=10, padx=10, expand=True, fill="both")

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

        ttk.Label(contact_frame, text="QQ:", font=('Arial', 10, 'bold')).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Label(contact_frame, text="88179096", font=('Arial', 10)).grid(row=0, column=1, padx=10, pady=5, sticky="w")

        ttk.Label(contact_frame, text="网址:", font=('Arial', 10, 'bold')).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        url_label = ttk.Label(contact_frame, text="www.itvip.com.cn", 
                              foreground="blue", cursor="hand2", font=('Arial', 10, 'underline'))
        url_label.grid(row=1, column=1, padx=10, pady=5, sticky="w")
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
        
        # 6 列：项目名称, 单位, 数量, 单价, 金额, 备注
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
        
        # --- 填票人/收款人/收款单位 全部在一行 ---
        payee_frame = ttk.Frame(self.receipt_frame)
        payee_frame.pack(fill="x", pady=10)
        
        # 1. 填票人 (左侧)
        ttk.Label(payee_frame, text="填票人:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.issuer_entry = ttk.Entry(payee_frame, width=15) # 缩短宽度
        self.issuer_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # 2. 收款人 (中侧)
        ttk.Label(payee_frame, text="收款人:").grid(row=0, column=2, padx=20, pady=5, sticky="w")
        self.payee_entry = ttk.Entry(payee_frame, width=15) # 缩短宽度
        self.payee_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        # 3. 收款单位 (右侧，拉伸占据剩余空间)
        ttk.Label(payee_frame, text="收款单位:").grid(row=0, column=4, padx=20, pady=5, sticky="w")
        self.payee_company_entry = ttk.Entry(payee_frame, width=10) 
        self.payee_company_entry.grid(row=0, column=5, padx=5, pady=5, sticky="ew") # 关键：sticky="ew" 允许拉伸
        
        # 关键：设置第5列 (收款单位输入框所在的列) 权重为1，使其占据剩余空间
        payee_frame.grid_columnconfigure(5, weight=1) 
        
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
        issuer = self.db.load_setting("issuer", "")
        
        self.payee_entry.delete(0, tk.END)
        self.payee_entry.insert(0, payee)
        
        self.payee_company_entry.delete(0, tk.END)
        self.payee_company_entry.insert(0, payee_company)
        
        self.issuer_entry.delete(0, tk.END)
        self.issuer_entry.insert(0, issuer)

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
        
        columns = ("no", "name", "date", "amount") 
        self.query_tree = ttk.Treeview(query_result_frame, columns=columns, show="headings", height=15)
        
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
        ttk.Button(query_btn_frame, text="生成PDF", command=self.view_details_pdf).pack(side="left", padx=10)
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
                quantity_str = values[col_map['quantity']].strip()
                unit_price_str = values[col_map['unit_price']].strip()
                
                if quantity_str:
                    try:
                        quantity = int(quantity_str)
                    except ValueError:
                        quantity = 0 
                else:
                    quantity = 0

                unit_price = float(unit_price_str or 0)
                
                amount = round(quantity * unit_price, 2)
                
                values[col_map['amount']] = f"{amount:.2f}"
                values[col_map['quantity']] = str(quantity) if quantity_str else ""
                
                self.items_tree.item(item_id, values=values)
                
                total += amount
            except (ValueError, IndexError):
                values[col_map['amount']] = "0.00"
                values[col_map['quantity']] = values[col_map['quantity']].strip() or "" 
                values[col_map['unit_price']] = values[col_map['unit_price']].strip() or "" 
                self.items_tree.item(item_id, values=values)
        
        self.total_num_var.set(f"合计: {total:.2f}元")
        self.total_cap_var.set(f"合计: 人民币(大写) {convert_to_chinese_caps(total)}")
        return total

    def add_item_row(self):
        """添加一个空的明细行，数量和单价默认值设置为空"""
        # item_name(0), unit(1), quantity(2), unit_price(3), amount(4), notes(5)
        empty_values = ("", "", "", "", "", "") 
        
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
             # 如果上一个编辑器存在，先清除它
             self.entry_editor.destroy()
             self.entry_editor = None

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

        try:
             col_index = int(column.replace('#', '')) - 1
        except ValueError:
             return

        if not item or col_index not in self.editable_cols: 
            return

        x, y, w, h = self.items_tree.bbox(item, column)
        current_value = self.items_tree.set(item, column)
        
        justify_align = 'right' if col_index in (2, 3) else 'left'
        
        self.entry_editor = ttk.Entry(self.items_tree, justify=justify_align)
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

        new_value = self.entry_editor.get().strip()
        
        # --- 针对数量和单价的输入检查 ---
        if col_index == 2: # 数量 (quantity) 必须是整数
            if new_value:
                 try:
                     int(new_value)
                 except ValueError:
                     messagebox.showerror("输入错误", "数量必须是整数！")
                     self.entry_editor.focus() 
                     return
        
        elif col_index == 3: # 单价 (unit_price) 必须是数字 (可以是小数)
            if new_value:
                 try:
                     float(new_value)
                 except ValueError:
                     messagebox.showerror("输入错误", "单价必须是有效的数字！")
                     self.entry_editor.focus() 
                     return
        # --- 结束输入检查 ---

        current_values = list(self.items_tree.item(item, 'values'))
        while len(current_values) <= col_index:
             current_values.append("")
        current_values[col_index] = new_value
        self.items_tree.item(item, values=current_values)

        self.entry_editor.destroy()
        self.entry_editor = None
        self.calculate_total()
        
        if is_tab:
            # FIX: 使用 lambda 默认参数捕获 col_index 的值，解决 NameError
            self.master.after(10, lambda c_idx=col_index: self.handle_tab_jump(item, c_idx))
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
        total_num = self.calculate_total() 
        client_name = self.client_name_entry.get().strip()
        
        if not client_name:
            messagebox.showwarning("警告", "客户名称不能为空！")
            return
            
        # 2. 准备明细数据
        items_raw = [self.items_tree.item(i, 'values') for i in self.items_tree.get_children()]
        items_raw = [item for item in items_raw if item and any(item)]
        
        if not items_raw:
            messagebox.showwarning("警告", "项目明细不能为空！")
            return

        total_cap = convert_to_chinese_caps(total_num)

        receipt_data = {
            'client_name': client_name,
            'total_amount_num': total_num,
            'total_amount_cap': total_cap
        }
        
        items_data = []
        for item in items_raw:
            try:
                # item 列表/元组有 6 个元素 (item_name, unit, quantity, unit_price, amount, notes)
                
                quantity_value = item[2] or 0
                unit_price_value = item[3] or 0
                
                quantity_num = int(float(quantity_value)) if quantity_value else 0
                unit_price_num = float(unit_price_value) if unit_price_value else 0.0
                amount_num = float(item[4] or 0)
                
                items_data.append([
                    item[0] or "", # 0: item_name
                    item[1] or "", # 1: unit
                    quantity_num,  # 2: quantity (整数)
                    unit_price_num, # 3: unit_price (浮点数)
                    amount_num, # 4: amount (浮点数)
                    item[5] or ""  # 5: notes (备注，确保被获取)
                ])
            except (ValueError, IndexError) as e:
                messagebox.showerror("数据错误", f"项目明细中存在无效数据或非整数数量，请检查。错误: {e}")
                return

        # 3. 数据库保存
        receipt_no = self.db.save_receipt(receipt_data, items_data)
        
        if receipt_no:
            # 4. 保存当前的收款人/填票人信息作为默认值
            payee_current = self.payee_entry.get().strip()
            payee_company_current = self.payee_company_entry.get().strip()
            issuer_current = self.issuer_entry.get().strip() # 获取填票人信息
            
            self.db.save_setting("payee", payee_current)
            self.db.save_setting("payee_company", payee_company_current)
            self.db.save_setting("issuer", issuer_current) # 关键修改：保存填票人信息
            
            messagebox.showinfo("成功", f"收据 {receipt_no} 已成功保存到数据库。")
            self.clear_receipt_form()
            self.refresh_query_results()
            
            if generate_pdf:
                self._generate_and_open_pdf(receipt_no)

    def _generate_and_open_pdf(self, receipt_no, parent_window=None):
        """生成PDF并尝试调用系统打开"""
        
        # 重新加载数据以确保 PDF 中包含最新的收款人信息
        pdf_data, pdf_items = self.db.fetch_receipt_details(receipt_no)
        
        if not pdf_data or not pdf_items:
             messagebox.showerror("错误", f"未找到收据编号 {receipt_no} 的详细信息。")
             return

        generator = PDFGenerator(pdf_data, pdf_items)
        pdf_filename = generator.generate()
        
        try:
            if os.name == 'nt': 
                os.startfile(pdf_filename) 
            elif os.name == 'posix': 
                os.system(f"open {pdf_filename}") or os.system(f"xdg-open {pdf_filename}")
            else:
                 messagebox.showwarning("打开失败", f"PDF 文件已生成: {pdf_filename}，但无法自动打开。", parent=parent_window)
                 return

        except Exception as e:
            messagebox.showwarning("打开失败", f"PDF 文件已生成: {pdf_filename}，但无法自动打开。错误: {e}", parent=parent_window)


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
        """生成选中收据的PDF文件 (从查询页调用)"""
        selected_item = self.query_tree.focus()
        if not selected_item:
            messagebox.showwarning("警告", "请先在列表中选择一张收据。")
            return
            
        receipt_no = self.query_tree.item(selected_item, 'values')[0]
        self._generate_and_open_pdf(receipt_no)
            
    def print_receipt(self):
        """生成选中收据的PDF并尝试打开文件（从查询页调用）"""
        selected_item = self.query_tree.focus()
        if not selected_item:
            messagebox.showwarning("警告", "请先在列表中选择一张收据进行打印。")
            return
            
        receipt_no = self.query_tree.item(selected_item, 'values')[0]
        self._generate_and_open_pdf(receipt_no)
        
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
        
        DETAIL_WINDOW_WIDTH = 750 
        DETAIL_WINDOW_HEIGHT = 600
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        center_x = int((screen_width / 2) - (DETAIL_WINDOW_WIDTH / 2))
        center_y = int((screen_height / 2) - (DETAIL_WINDOW_HEIGHT / 2))
        details_window.geometry(f"{DETAIL_WINDOW_WIDTH}x{DETAIL_WINDOW_HEIGHT}+{center_x}+{center_y}")
        
        # 设置详情窗口图标
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
            # 数量显示为整数
            quantity_display = str(int(item[2])) 
            
            detail_tree.insert("", "end", values=(
                item[0], item[1], 
                quantity_display, f"{item[3]:.2f}", 
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
        
        # 填票人 
        ttk.Label(payee_details_frame, text="填票人:", font=('Arial', 10)).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Label(payee_details_frame, text=receipt_data.get('issuer', ''), font=('Arial', 10)).grid(row=0, column=1, padx=5, pady=2, sticky="w")
        
        # 收款人
        ttk.Label(payee_details_frame, text="收款人:", font=('Arial', 10)).grid(row=0, column=2, padx=20, pady=2, sticky="w")
        ttk.Label(payee_details_frame, text=receipt_data.get('payee', ''), font=('Arial', 10)).grid(row=0, column=3, padx=5, pady=2, sticky="w")
        
        # 收款单位 (右侧拉伸)
        ttk.Label(payee_details_frame, text="收款单位:", font=('Arial', 10)).grid(row=0, column=4, padx=20, pady=2, sticky="w")
        ttk.Label(payee_details_frame, text=receipt_data.get('payee_company', ''), font=('Arial', 10)).grid(row=0, column=5, padx=5, pady=2, sticky="w")
        
        payee_details_frame.grid_columnconfigure(5, weight=1)
        
        # 打印按钮
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill="x", pady=10)
        
        ttk.Button(action_frame, text="打印收据", 
                   command=lambda: self._generate_and_open_pdf(receipt_no, details_window)).pack(side="right")
        
        details_window.protocol("WM_DELETE_WINDOW", lambda: details_window.destroy())


# --- 5. 运行程序 ---
if __name__ == "__main__":
    if 'webbrowser' not in sys.modules and sys.platform != 'win32':
        pass
        
    try:
        root = tk.Tk()
        app = ReceiptApp(root)
        root.mainloop()
    except Exception as e:
         # 如果程序启动失败，弹出错误提示
         messagebox.showerror("程序启动错误", f"程序启动失败: {e}")
         sys.exit(1)
    finally:
        # 确保数据库连接关闭
        if 'app' in locals():
            try:
                app.db.close()
            except Exception:
                pass