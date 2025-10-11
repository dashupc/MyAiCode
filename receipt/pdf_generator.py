from fpdf import FPDF # 兼容老版 fpdf
import os
import tkinter as tk
from tkinter import messagebox
from utils import resource_path 

class PDFGenerator:
    """使用 fpdf (兼容老版本) 生成收据 PDF，使用标准 A4 尺寸，并只显示有内容的行"""
    def __init__(self, receipt_data, items_data):
        self.data = receipt_data
        
        # --- 兼容老版本 fpdf：使用标准 'A4' 字符串 ---
        self.pdf = FPDF('P', 'mm', 'A4') 
        
        # 使用 A4 尺寸：内容宽度为 210mm - 左右边距 20mm*2 = 170mm
        self.margin = 20
        self.content_width = 210 - (self.margin * 2) # 170mm

        font_path = resource_path('simhei.ttf') 
        self.pdf_font_name = 'SimHei' 
        
        try:
             # 加载字体 (fpdf 同样支持 .add_font)
             self.pdf.add_font(self.pdf_font_name, '', font_path) 
             self.pdf.set_font(self.pdf_font_name, '', 12)
        except FileNotFoundError:
             messagebox.showwarning("字体缺失", "未找到 simhei.ttf，PDF将无法正确显示中文，请将字体文件放在程序目录下。")
             self.pdf.set_font('Arial', '', 12) 
             self.pdf_font_name = 'Arial' 
        
        # 核心过滤逻辑：移除项目名称为空的行
        self.items = [
            item for item in items_data 
            if item and item[0] and item[0].strip()
        ]
        
    def _draw_content(self):
        """内部方法：负责将所有内容绘制到 PDF 页面上"""
        
        # 设置页边距 (20mm)
        self.pdf.set_margins(left=self.margin, top=self.margin, right=self.margin) 
        self.pdf.add_page() 
        
        # --- 头部信息 ---
        # 标题和编号
        self.pdf.set_font(self.pdf_font_name, '', 18)
        self.pdf.cell(0, 10, "收  据", 0, 1, 'C') 
        
        self.pdf.set_font(self.pdf_font_name, '', 11)
        self.pdf.cell(0, 7, f"NO: {self.data['receipt_no']}", 0, 1, 'R') 
        
        self.pdf.ln(5)

        # 客户信息
        client_name = self.data.get('client_name', ' ')
        issue_date = self.data.get('issue_date', ' ')
        
        self.pdf.set_font(self.pdf_font_name, '', 12)
        self.pdf.cell(80, 8, f"付款单位/个人: {client_name}", 0, 0, 'L')
        self.pdf.cell(0, 8, f"日期: {issue_date}", 0, 1, 'R')
        
        self.pdf.ln(5)

        # --- 项目明细表格 ---
        # 7 列宽度调整，总和仍为 170mm： [序号, 项目名称, 单位, 数量, 单价(元), 金额(元), 备注]
        col_widths = [10, 40, 15, 20, 30, 30, 25] # 总和: 170
        
        # 表头
        self.pdf.set_font(self.pdf_font_name, '', 10)
        self.pdf.set_fill_color(220, 220, 220) 
        # --- 关键修改：添加 '备注' 列 ---
        headers = ["序号", "项目名称", "单位", "数量", "单价(元)", "金额(元)", "备注"]
        for i, header in enumerate(headers):
            self.pdf.cell(col_widths[i], 8, header, 1, 0, 'C', 1)
        self.pdf.ln()
        
        # 格式化数据 (item[2] 数量已保证是整数)
        items_formatted = [
            # 7个元素: 序号, item[0], item[1], item[2](数量), item[3](单价), item[4](金额), item[5](备注)
            (i + 1, item[0], item[1], str(int(item[2])), f"{item[3]:.2f}", f"{item[4]:.2f}", item[5]) 
            for i, item in enumerate(self.items)
        ]

        # 表格内容主体
        self.pdf.set_font(self.pdf_font_name, '', 10)
        for row in items_formatted:
            for i, text in enumerate(row):
                # 表格内容居中显示 (备注列可以考虑左对齐，但此处保持居中)
                self.pdf.cell(col_widths[i], 8, str(text), 1, 0, 'C') 
            self.pdf.ln()

        # --- 底部合计信息和收款人 ---
        total_text = f"合计: 人民币(大写) {self.data['total_amount_cap']}"
        total_num_text = f"合计: {self.data['total_amount_num']:.2f}元"
        
        self.pdf.ln(5)
        self.pdf.set_font(self.pdf_font_name, '', 12)
        
        # 合计信息占据两列
        self.pdf.cell(80, 8, total_text, 0, 0, 'L')
        self.pdf.cell(0, 8, total_num_text, 0, 1, 'R')
        
        self.pdf.ln(10)

        # 底部信息：填票人、收款人、收款单位
        payee = self.data.get('payee', ' ')
        payee_company = self.data.get('payee_company', ' ')
        # --- 确保获取了填票人数据 ---
        issuer = self.data.get('issuer', ' ')

        # 底部布局：
        self.pdf.set_font(self.pdf_font_name, '', 11)
        
        # 1. 填票人 (左对齐，宽 50) 
        self.pdf.cell(50, 7, f"填票人: {issuer}", 0, 0, 'L')
        
        # 2. 收款人 (左对齐，宽 50)
        self.pdf.cell(50, 7, f"收款人: {payee}", 0, 0, 'L') 
        
        # 3. 收款单位 (右对齐，占据剩余宽度)
        self.pdf.cell(0, 7, f"收款单位: {payee_company}", 0, 1, 'R') 

    def generate(self):
        """生成 PDF 文件并返回文件名"""
        
        self._draw_content() 

        receipt_no = self.data.get('receipt_no', 'TEMP') 
        filename = f"Receipt_{receipt_no}.pdf"
        
        self.pdf.output(filename, 'F') 
        
        return filename