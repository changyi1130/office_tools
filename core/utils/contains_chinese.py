import re


def contains_chinese(string):
    """检查字符串中是否包含中文"""

    pattern = re.compile(r'[\u4e00-\u9fff]')
    return bool(pattern.search(string))
