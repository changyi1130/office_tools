"""
运行中文预编宏
"""
from pathlib import Path
from typing import List, Callable

from natsort import os_sorted

from core.tasks.rename_files import is_file_hidden, select_directory
from core.utils.open_file_dialog import open_file_dialog


def hidden_untranslated_content(update_info: Callable[[str], None]):
    """"""

    # 选择目录
    selected_dir = select_directory()
    if not selected_dir:
        update_info("操作已取消：未选择目录")
        return

    update_info(f"正在处理文件，请稍候...")

    # 对文件执行宏


def get_files(update_info: Callable[[str], None]) -> List[str]:
    """获取要处理文件的路径"""

    files = open_file_dialog(
        window_title="选择文件",
        file_filter=[
            ("Word文档", "*.doc*"),
            ("所有文件", "*.*")],
        multi_select=True)
