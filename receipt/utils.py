import os
import sys

def resource_path(relative_path):
    """获取脚本或打包后的程序目录下的资源文件路径"""
    try:
        # PyInstaller 临时目录
        base_path = sys._MEIPASS
    except Exception:
        # 脚本运行目录
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def convert_to_chinese_caps(number):
    """将数字金额转换为人民币大写"""
    num_str = "{:.2f}".format(float(number))
    
    # 转换逻辑（简化版，您可能需要更完整的实现）：
    chinese_nums = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']
    units = ['', '拾', '佰', '仟']
    big_units = ['', '万', '亿']
    
    integer_part, decimal_part = num_str.split('.')
    
    def _convert_part(part):
        output = ''
        i = 0
        for num in part[::-1]:
            n = int(num)
            if n != 0:
                output = chinese_nums[n] + units[i % 4] + output
            elif i % 4 == 0 and output:
                 output = big_units[i // 4] + output
            elif output and output[0] != '零':
                 output = '零' + output
                 
            i += 1
        
        # 清理多余的“零”
        while '零零' in output:
            output = output.replace('零零', '零')
        if output.endswith('零'):
            output = output[:-1]

        # 补全大单位
        for bu in big_units[1:]:
             if bu in output and not output.endswith(bu):
                 output = output.replace(bu, bu + '零')

        return output.replace('零万', '万').replace('零亿', '亿').replace('亿万', '亿')
        
    integer_converted = _convert_part(integer_part)
    
    # 处理整数部分
    if integer_converted:
        result = integer_converted + '元'
    else:
        result = '零元'
        
    # 处理小数部分
    jiao = int(decimal_part[0])
    fen = int(decimal_part[1])
    
    if jiao == 0 and fen == 0:
        result += '整'
    elif jiao != 0 and fen == 0:
        result += chinese_nums[jiao] + '角整'
    elif jiao == 0 and fen != 0:
        result += '零' + chinese_nums[fen] + '分'
    else:
        result += chinese_nums[jiao] + '角' + chinese_nums[fen] + '分'
        
    return result