import re

def contains_chinese(string):
    # 正则表达式匹配中文字符
    pattern = re.compile(r'[\u4e00-\u9fff]')
    return bool(pattern.search(string))