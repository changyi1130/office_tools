from tkinter import filedialog
import os

# 文件筛选器
doc_file_filter = [('Word文档', '*.doc*'), ('所有文件', '*')]

def open_single_file(window_title, file_filter=doc_file_filter):
    """
    获取文件地址（单个文件）
    
    :param window_title: 窗口标题
    :param file_filter: 过滤器 (default: *.doc*)
    """
    
    original_file_path = filedialog.askopenfilename(title=window_title,
                                                    filetypes=file_filter)

    if original_file_path == '':
        print("用户取消选择")
        return None
    else:
        # 规范化路径格式
        decoded_path = os.path.normpath(original_file_path)

        return decoded_path