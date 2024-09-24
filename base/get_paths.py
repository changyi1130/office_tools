from tkinter import filedialog
import os

def get_paths(window_title, file_filter):
    """
    获取文件路径，可打开多个文件
    
    :param window_title: 窗口标题
    :param file_filter: 文件过滤器
    """
    original_file_paths = filedialog.askopenfilenames(title=window_title,
                                                      filetypes=file_filter)

    if original_file_paths == '':
        print("用户取消选择")
        return None
    else:
        # 规范化路径格式
        decoded_paths = []
        for original_file_path in original_file_paths:
            decoded_paths.append(os.path.normpath(original_file_path))
        
        return decoded_paths