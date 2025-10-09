from fpdf import FPDF
import os
import tkinter as tk
from tkinter import messagebox
from utils import resource_path # 相对导入 utils 模块

class PDFGenerator:
    """使用 fpdf2 生成收据 PDF，不包含公章绘制"""
    def __init__(self, receipt_data, items_data):
        self.data = receipt_data
        self.items = items_data
        self.pdf = FPDF('P', 'mm', 'A4')
        
        font_path = resource_path('simhei.ttf') 
        self.pdf_font_name = 'SimHei' 
        
        try:
             # 加载字体
             self.pdf.add_font(self.pdf_font_name, '', font_path) 
             self.pdf.set_font(self.pdf_font_name, '', 12)
        except FileNotFoundError:
             # 字体缺失时弹窗提示，并回退
             messagebox.showwarning("字体缺失", "未找到 simhei.ttf，PDF将无法正确显示中文，请将字体文件放在程序目录下。")
             self.pdf.set_font('Arial', '', 12) 
             self.pdf_font_name = 'Arial' 
        
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
            # item 是一个元组/列表，结构为 (item_name, unit, quantity(int), unit_price(float), amount(float), notes)
            
            # ！！！ 数量显示为整数 ！！！
            quantity_display = str(int(item[2])) 

            row = [
                item[0] or '', item[1] or '', 
                quantity_display, f"{item[3]:.2f}", 
                f"{item[4]:.2f}", item[5] or ''
            ]
            
            # ！！！ 统一设置对齐为居中 (C) ！！！
            for i, text in enumerate(row):
                self.pdf.cell(col_widths[i], 8, str(text), 1, 0, 'C') 
            self.pdf.ln()
            
        # 填充空白行（保持表格高度一致）
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
        
        # 合计信息占据两列
        self.pdf.cell(100, 8, total_text, 0, 0, 'L')
        self.pdf.cell(90, 8, total_num_text, 0, 1, 'R')
        
        self.pdf.ln(10) # 底部留出更多空间

        # 收款人/收款单位 信息
        payee = self.data.get('payee', '')
        payee_company = self.data.get('payee_company', '')
        issue_date = self.data['issue_date']

        self.pdf.set_font(self.pdf_font_name, '', 10)
        
        # ！！！ 调整顺序和对齐：收款人(左对齐), 收款单位(右对齐) ！！！
        self.pdf.cell(90, 8, f"收款人: {payee}", 0, 0, 'L')
        self.pdf.cell(90, 8, f"收款单位: {payee_company}", 0, 1, 'R')
        
        # 日期信息
        self.pdf.ln(5)
        self.pdf.cell(0, 8, f"开具日期: {issue_date}", 0, 1, 'R')


        filename = f"Receipt_{self.data['receipt_no']}.pdf"
        self.pdf.output(filename)
        return filename