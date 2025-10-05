import random
import string

def generate_registration_key():
    """生成有效的注册码"""
    # 生成前三部分随机字符（大写字母和数字）
    part1 = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
    part2 = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
    part3 = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
    
    # 计算校验和（与验证函数保持一致）
    checksum = sum(ord(c) for c in part1 + part2 + part3) % 10000
    part4 = f"{checksum:04d}"  # 确保是4位数字，不足补零
    
    # 组合成注册码格式
    return f"{part1}-{part2}-{part3}-{part4}"

if __name__ == "__main__":
    # 生成10个示例注册码
    print("生成的注册码：")
    for _ in range(10):
        print(generate_registration_key())
