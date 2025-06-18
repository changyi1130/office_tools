import os
import tkinter as tk
from tkinter import filedialog


def open_file_dialog(
        window_title: str,
        file_filter=None,
        multi_select: bool = False
) -> list | str | None:
    """
    通用文件选择对话框

    :param window_title: 窗口标题
    :param file_filter: 文件类型过滤器（默认筛选Word文档）
    :param multi_select: 是否允许多选（默认单选）
    :return: 单选返回字符串路径，多选返回路径列表，取消返回None
    """
    if file_filter is None:
        file_filter = [('Word文档', '*.doc*'), ('所有文件', '*')]

    root = tk.Tk()
    root.withdraw()

    # 根据选择模式调用不同方法
    dialog_method = filedialog.askopenfilenames if multi_select else filedialog.askopenfilename
    selected_paths = dialog_method(title=window_title, filetypes=file_filter)

    root.destroy()  # 销毁根窗口

    # 统一处理取消操作
    if not selected_paths:
        print("用户取消选择")
        return None

    # 规范路径格式
    normalize = lambda p: os.path.normpath(p)

    # 返回类型处理
    if multi_select:
        return [normalize(path) for path in selected_paths]
    else:
        return normalize(selected_paths)
