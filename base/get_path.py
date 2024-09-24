from tkinter import filedialog
import os

def get_path(window_title, file_filter):
    """
    获取文件路径，打开单个文件
    
    :param window_title: 窗口标题
    :param file_filter: 文件过滤器
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