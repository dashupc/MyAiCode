from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
import os

# 使用系统字体路径（宋体）
font_path = "C:/Windows/Fonts/simsun.ttf"

# 检查字体文件是否存在
if not os.path.exists(font_path):
    raise FileNotFoundError(f"字体文件未找到：{font_path}")

# 注册宋体字体
pdfmetrics.registerFont(TTFont("SimSun", font_path))

def export(data):
    filename = f"{data['receipt_no']}.pdf"
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    # 标题
    c.setFont("SimSun", 16)
    c.drawCentredString(width / 2, height - 50, "收款收据")

    # 编号
    c.setFont("SimSun", 10)
    c.setFillColor(colors.red)
    c.drawRightString(width - 50, height - 70, f"编号：{data['receipt_no']}")
    c.setFillColor(colors.black)

    # 客户名称、开票人、日期
    c.setFont("SimSun", 11)
    c.drawString(50, height - 90, f"客户名称：{data['customer_name']}")
    c.drawString(50, height - 110, f"开票人：{data['issuer']}")
    c.drawString(250, height - 110, f"日期：{data['date']}")

    # 表格数据
    table_data = [["品名及规格", "单位", "数量", "单价", "金额", "备注"]]
    for item in data["items"]:
        table_data.append(item)

    # 表格样式
    table = Table(table_data, colWidths=[90, 50, 50, 60, 60, 100])
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, -1), "SimSun"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey)
    ]))

    # 表格位置
    table.wrapOn(c, width, height)
    table.drawOn(c, 50, height - 300)

    # 合计金额
    c.setFont("SimSun", 11)
    c.drawString(50, height - 320, f"合计金额：¥{data['total']}")
    c.drawString(250, height - 320, f"大写金额：{data['upper_total']}")

    # 盖章区域
    c.setFont("SimSun", 10)
    c.drawString(50, height - 350, "单位盖章：__________________")

    c.save()
