from fpdf import FPDF
import os
import sys
from PIL import Image
from utils import resource_path

# FPDF 的核心是使用 mm 作为单位
PAGE_WIDTH_MM = 210
PAGE_HEIGHT_MM = 150
MARGIN_MM = 20 # 左右边距 20mm

class PDFGenerator:
    """使用 fpdf2 生成收据 PDF，包含印章绘制和粗体字体注册"""
    def __init__(self, receipt_data, items_data):
        self.receipt_data = receipt_data
        self.items_data = items_data
        
        self.pdf = FPDF('P', 'mm', 'A4')
        self.pdf.set_auto_page_break(auto=True, margin=15)

        # 内容区域宽度 (210 - 20*2 = 170mm)
        self.content_width = PAGE_WIDTH_MM - 2 * MARGIN_MM 

        font_path = resource_path('simhei.ttf') 
        self.pdf_font_name = 'SimHei' 
        
        try:
             # 注册普通字体
             self.pdf.add_font(self.pdf_font_name, '', font_path) 
             # 注册粗体字体
             self.pdf.add_font(self.pdf_font_name, 'B', font_path) 
        except FileNotFoundError:
             self.pdf_font_name = 'Arial' 
        
        self.pdf.set_font(self.pdf_font_name, '', 12)

    def generate(self):
        """生成 PDF 文件，并返回文件名"""
        
        self.pdf.add_page()
        self.pdf.set_margin(MARGIN_MM)
        
        receipt_no = self.receipt_data['receipt_no']
        client_name = self.receipt_data['client_name']
        
        safe_client_name = "".join(c for c in client_name if c.isalnum() or c in (' ', '_')).rstrip()
        filename = f"收据_{receipt_no}_{safe_client_name}.pdf"

        # --- 1. 标题 ---
        self.pdf.set_font(self.pdf_font_name, 'B', 24) 
        self.pdf.cell(self.content_width, 15, '收 款 收 据', 0, 1, 'C') 
        self.pdf.ln(5)

        # --- 2. 收据编号和日期 ---
        self.pdf.set_font(self.pdf_font_name, '', 10)
        self.pdf.cell(self.content_width / 2, 5, f"开具日期：{self.receipt_data['issue_date']}", 0, 0, 'L')
        self.pdf.cell(self.content_width / 2, 5, f"NO. {receipt_no}", 0, 1, 'R')
        self.pdf.ln(5)

        # --- 3. 客户名称 ---
        self.pdf.set_font(self.pdf_font_name, '', 12)
        self.pdf.cell(30, 8, '客户名称:', 0, 0, 'L')
        self.pdf.set_font(self.pdf_font_name, 'B', 12) 
        self.pdf.cell(self.content_width - 10, 8, self.receipt_data['client_name'], 0, 1, 'L')
        self.pdf.ln(2)

        # --- 4. 表格头部 【添加序号，调整列宽和对齐】---
        # 新列宽: 序号(10) + 项目名称(50) + 单位(15) + 数量(20) + 单价(25) + 金额(25) + 备注(25) = 170mm
        col_widths = [10, 50, 15, 20, 25, 25, 25] 
        col_names = ["序号", "项目名称", "单位", "数量", "单价(元)", "金额(元)", "备注"]
        row_height = 8

        self.pdf.set_fill_color(240, 240, 240)
        self.pdf.set_draw_color(0, 0, 0)
        self.pdf.set_line_width(0.3)
        self.pdf.set_font(self.pdf_font_name, 'B', 10) 

        # 绘制表头
        for i, name in enumerate(col_names):
            # 序号(0), 单位(2), 数量(3), 单价(4), 金额(5) 居中 ('C')
            align = 'C' if i in [0, 2, 3, 4, 5, 6] else 'L'
            self.pdf.cell(col_widths[i], row_height, name, 1, 0, align, True)
        self.pdf.ln()

        # --- 5. 表格内容 ---
        self.pdf.set_font(self.pdf_font_name, '', 10) 
        
        items_formatted = []
        # 在数据中添加序号
        for i, item in enumerate(self.items_data):
            quantity_display = str(int(item[2])) 
            items_formatted.append([
                str(i + 1),               # 序号
                item[0], item[1], quantity_display, 
                f"{item[3]:.2f}", f"{item[4]:.2f}", item[5]
            ])

        for row in items_formatted:
            # 0: 序号 - 居中 ('C')
            self.pdf.cell(col_widths[0], row_height, row[0], 1, 0, 'C') 
            # 1: 项目名称 - 左对齐 ('L')
            self.pdf.cell(col_widths[1], row_height, row[1], 1, 0, 'L') 
            # 2: 单位 - 居中 ('C')
            self.pdf.cell(col_widths[2], row_height, row[2], 1, 0, 'C') 
            # 3: 数量 - 居中 ('C') 【对齐修改】
            self.pdf.cell(col_widths[3], row_height, row[3], 1, 0, 'C') 
            # 4: 单价 - 居中 ('C') 【对齐修改】
            self.pdf.cell(col_widths[4], row_height, row[4], 1, 0, 'C') 
            # 5: 金额 - 居中 ('C') 【对齐修改】
            self.pdf.cell(col_widths[5], row_height, row[5], 1, 0, 'C') 
            # 6: 备注 - 左对齐 ('C')
            self.pdf.cell(col_widths[6], row_height, row[6], 1, 1, 'L') 
        
        # --- 6. 合计行 【更新列宽计算】---
        self.pdf.set_font(self.pdf_font_name, 'B', 10) 
        total_text = f"合计 (大写)：{self.receipt_data['total_amount_cap']}"
        total_num_text = f"合计 (小写)：¥{self.receipt_data['total_amount_num']:.2f}"
        
        # 序号(10)+项目名称(50)+单位(15)+数量(20)+单价(25) = 120mm
        cap_width = col_widths[0] + col_widths[1] + col_widths[2] + col_widths[3] + col_widths[4] 
        # 金额(25) + 备注(25) = 50mm
        num_width = col_widths[5] + col_widths[6]
        
        self.pdf.set_fill_color(220, 220, 220)
        
        self.pdf.cell(cap_width, row_height, total_text, 1, 0, 'L', True)
        self.pdf.cell(num_width, row_height, total_num_text, 1, 1, 'R', True)
        
        self.pdf.ln(10)

        # --- 7. 底部信息：填票人 | 收款人 | 收款单位(盖章) ---
        self.pdf.set_font(self.pdf_font_name, '', 12)
        
        payee_company = self.receipt_data.get('payee_company', 'N/A')
        issuer = self.receipt_data.get('issuer', 'N/A')
        payee = self.receipt_data.get('payee', 'N/A')
        
        # 定义三列宽度 (填票人减少 20mm, 后两列左移)
        col_1_w = 30  # 填票人 
        col_2_w = 50  # 收款人 
        col_3_w = self.content_width - col_1_w - col_2_w # 收款单位 
        
        # 1. 填票人：
        self.pdf.cell(col_1_w, 5, f"填票人：{issuer}", 0, 0, 'L')
        # 2. 收款人：
        self.pdf.cell(col_2_w, 5, f"收款人：{payee}", 0, 0, 'L')
        # 3. 收款单位(盖章)：
        self.pdf.cell(col_3_w, 5, f"收款单位(盖章)：{payee_company}", 0, 1, 'L')
        
        # --- 8. 盖章图绘制 (调整尺寸和位置) ---
        
        stamp_size_mm = 40 # 40mm
        
        # 计算收款单位文字的起始 X 坐标
        start_x_payee_company = MARGIN_MM + col_1_w + col_2_w 
        
        # 印章 X 坐标：起始 X + “收款单位(盖章)”标签的宽度 (约 35mm) - 半个印章宽度 (20mm) + 右移 20mm
        stamp_x_mm = start_x_payee_company + 35 - stamp_size_mm / 2 + 20 
        
        # Y 坐标：当前 Y 坐标 (文字基线) 向上抬高 10mm
        stamp_y_mm = self.pdf.get_y() - 10 - stamp_size_mm / 2 

        stamp_path = resource_path('yinzhang.PNG')
        
        if os.path.exists(stamp_path):
            try:
                self.pdf.image(stamp_path, x=stamp_x_mm, y=stamp_y_mm, w=stamp_size_mm, h=stamp_size_mm)
            except Exception as e:
                print(f"警告：无法绘制印章图片。错误: {e}")

        # --- 保存 PDF ---
        self.pdf.output(filename)
        return filename