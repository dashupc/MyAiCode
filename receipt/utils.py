import os
import sys

# --- 资源路径函数 (解决打包后文件找不到的问题) ---
def resource_path(relative_path):
    """获取资源文件的绝对路径，以兼容 PyInstaller 单文件模式"""
    try:
        # PyInstaller 在运行时会将所有 --add-data 的文件解压到 _MEIPASS 目录
        base_path = sys._MEIPASS
    except Exception:
        # 如果不是 PyInstaller 运行，使用当前工作目录
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# --- 金额大写转换函数 ---
def convert_to_chinese_caps(num):
    """将数字金额转换为中文大写（简化实现）"""
    num = abs(num) 
    # 只处理整数部分，因为收据金额通常为整数元加整角分，这里简化为元整
    
    CAPS = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"]
    UNITS = ["", "拾", "佰", "仟"]
    GROUPS = ["", "万", "亿"]
    
    # 将金额四舍五入到两位小数，并分离整数部分
    num_str = f"{num:.2f}"
    integer_part = num_str.split('.')[0]
    
    result = ""
    group_index = 0
    
    while integer_part:
        chunk = integer_part[-4:]
        integer_part = integer_part[:-4]
        
        chunk_result = ""
        # 遍历千、百、十、个
        for i, char in enumerate(reversed(chunk)):
            digit = int(char)
            if digit != 0:
                chunk_result = CAPS[digit] + UNITS[i] + chunk_result
            elif chunk_result and not chunk_result.startswith("零"):
                chunk_result = CAPS[digit] + chunk_result
        
        if chunk_result:
            result = chunk_result.rstrip("零") + GROUPS[group_index] + result
        
        group_index += 1

    # 优化处理连续零和单位
    result = result.replace("零万", "万").replace("零亿", "亿").strip("零")
    
    if not result:
        return "零元整"
        
    return result + "元整"