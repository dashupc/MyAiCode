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
        
        # --- 核心修改: 宽度 170mm，高度 0 (自动适应) ---
        # 使用字典定义自定义格式，兼容性更佳
        custom_format = {
            'w': 170, # 宽度 170mm
            'h': 0    # 高度 0 表示随内容自动适应 (fpdf2 特性)
        }
        self.pdf = FPDF('P', 'mm', custom_format) 
        
        # 定义内容区域宽度 (宽度 170 - 左右边距 20*2 = 130mm)
        self.content_width = 130 

        font_path = resource_path('simhei.ttf') 
        self.pdf_font_name = 'SimHei' 
        
        try:
             # 加载字体
             self.pdf.add_font(self.pdf_font_name, '', font_path) 
             self.pdf.set_font(self.pdf_font_name, '', 12)
        except FileNotFoundError:
             messagebox.showwarning("字体缺失", "未找到 simhei.ttf，PDF将无法正确显示中文，请将字体文件放在程序目录下。")
             self.pdf.set_font('Arial', '', 12) 
             self.pdf_font_name = 'Arial' 
        
    def generate(self):
        self.pdf.add_page()
        
        # 设置页边距 (20mm)
        self.pdf.set_margins(left=20, top=20, right=20) 
        
        # 标题和编号
        self.pdf.set_font(self.pdf_font_name, '', 18)
        self.pdf.cell(0, 10, "收  据", 0, 1, 'C') # 居中标题
        
        self.pdf.set_font(self.pdf_font_name, '', 11)
        self.pdf.cell(0, 7, f"NO: {self.data['receipt_no']}", 0, 1, 'R') # 编号靠右
        
        self.pdf.ln(5)

        # 客户信息
        client_name = self.data.get('client_name', ' ')
        issue_date = self.data.get('issue_date', ' ')
        
        self.pdf.set_font(self.pdf_font_name, '', 12)
        # 宽度分配：左边 80，右边占据剩余 (50)
        self.pdf.cell(80, 8, f"付款单位/个人: {client_name}", 0, 0, 'L')
        self.pdf.cell(0, 8, f"日期: {issue_date}", 0, 1, 'R')
        
        self.pdf.ln(5)

        # 表格配置：col_widths总和必须等于 self.content_width (130)
        col_widths = [8, 40, 15, 15, 25, 27] 
        
        # 表头
        self.pdf.set_font(self.pdf_font_name, '', 10)
        self.pdf.set_fill_color(220, 220, 220) # 浅灰色背景
        headers = ["序号", "项目名称", "单位", "数量", "单价(元)", "金额(元)"]
        for i, header in enumerate(headers):
            self.pdf.cell(col_widths[i], 8, header, 1, 0, 'C', 1)
        self.pdf.ln()
        
        # 表格内容
        items_formatted = [
            (i + 1, item[0], item[1], f"{item[2]:.2f}", f"{item[3]:.2f}", f"{item[4]:.2f}")
            for i, item in enumerate(self.items)
        ]

        # 表格随内容自适应高度
        self.pdf.set_font(self.pdf_font_name, '', 10)
        for row in items_formatted:
            for i, text in enumerate(row):
                self.pdf.cell(col_widths[i], 8, str(text), 1, 0, 'C') 
            self.pdf.ln()

        # 合计行
        total_text = f"合计: 人民币(大写) {self.data['total_amount_cap']}"
        total_num_text = f"合计: {self.data['total_amount_num']:.2f}元"
        
        self.pdf.ln(5)
        self.pdf.set_font(self.pdf_font_name, '', 12)
        
        # 合计信息占据两列 (宽度 80 和 剩余)
        self.pdf.cell(80, 8, total_text, 0, 0, 'L')
        self.pdf.cell(0, 8, total_num_text, 0, 1, 'R')
        
        self.pdf.ln(10)

        # 底部信息框架
        payee = self.data.get('payee', ' ')
        payee_company = self.data.get('payee_company', ' ')
        issuer = self.data.get('issuer', ' ')

        # 打印 填票人 和 收款人
        self.pdf.set_font(self.pdf_font_name, '', 11)
        
        # 填票人 (宽 50) 和 收款人 (剩余宽度)
        self.pdf.cell(50, 7, f"填票人: {issuer}", 0, 0, 'L')
        self.pdf.cell(0, 7, f"收款人: {payee}", 0, 1, 'L') 

        # 收款单位 (独立一行)
        self.pdf.cell(0, 7, f"收款单位: {payee_company}", 0, 1, 'L')
        
        
    def output(self, filename="receipt.pdf"):
        """输出 PDF 文件"""
        self.pdf.output(filename, 'F')
        return os.path.abspath(filename)