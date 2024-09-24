# 打开 Word 运行指定宏

import win32com.client

def open_word(path):
    # 打开 Word 并运行宏

    # 创建 Word 对象
    word_app = win32com.client.Dispatch('Word.Application')

    # 打开 Word 文档
    try:
        return word_app.Documents.Open(path)
    except Exception as e:
        print(f"打开 Word 文档时出错：{e}")
        return None