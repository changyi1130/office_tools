import os
import tkinter as tk
from tkinter import filedialog
from contextlib import contextmanager
from typing import Optional, Union, List

# 全局Tk实例
_root_instance: Optional[tk.Tk] = None


@contextmanager
def _get_tk_root():
    """获取或创建Tk根窗口的上下文管理器"""
    global _root_instance
    try:
        if _root_instance is None:
            _root_instance = tk.Tk()
            _root_instance.withdraw()  # 隐藏主窗口
        yield _root_instance
    finally:
        # 不销毁实例，保留供后续使用
        pass


def _cleanup_tk_root():
    """清理Tk根窗口（程序退出时调用）"""
    global _root_instance
    if _root_instance is not None:
        try:
            _root_instance.destroy()
        except Exception:
            pass
        finally:
            _root_instance = None


def open_file_dialog(
        window_title: str,
        file_filter=None,
        multi_select: bool = False
) -> Union[List[str], str, None]:
    """
    通用文件选择对话框

    :param window_title: 窗口标题
    :param file_filter: 文件类型过滤器（默认筛选Word文档）
    :param multi_select: 是否允许多选（默认单选）
    :return: 单选返回字符串路径，多选返回路径列表，取消返回None
    """
    if file_filter is None:
        file_filter = [('Word文档', '*.doc*'), ('所有文件', '*')]

    with _get_tk_root():
        # 根据选择模式调用不同方法
        dialog_method = filedialog.askopenfilenames if multi_select else filedialog.askopenfilename
        selected_paths = dialog_method(title=window_title, filetypes=file_filter)

        # 统一处理取消操作
        if not selected_paths:
            return None

        # 规范路径格式
        normalize = lambda p: os.path.normpath(p)

        # 返回类型处理
        if multi_select:
            return [normalize(path) for path in selected_paths]
        else:
            return normalize(selected_paths)
