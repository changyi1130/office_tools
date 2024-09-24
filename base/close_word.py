# 关闭 Word 文档

import win32com.client

def close_word(doc, type: int):
    # 创建 Word 对象
    word_app = win32com.client.Dispatch('Word.Application')
    
    # 关闭 Word 文档
    try:
        doc.Close(type)
    except Exception as e:
        print(f"关闭 Word 文档进程出错: {e}")