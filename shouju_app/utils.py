def num_to_chinese_upper(num):
    units = ["元", "拾", "佰", "仟", "万", "拾", "佰", "仟", "亿"]
    nums = "零壹贰叁肆伍陆柒捌玖"
    result = ""
    num_str = str(int(float(num)))
    for i, digit in enumerate(reversed(num_str)):
        result = nums[int(digit)] + units[i] + result
    return result + "整"
