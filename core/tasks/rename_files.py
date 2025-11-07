"""
批量文件重命名功能
包含：
- 添加前缀编号
- 删除前缀编号
"""

import ctypes
import os
import re
from pathlib import Path
from tkinter import filedialog
from typing import List, Callable, Optional

from natsort import os_sorted


def is_file_hidden(file_path: str) -> bool:
    """检查文件是否为隐藏文件"""
    try:
        attrs = ctypes.windll.kernel32.GetFileAttributesW(file_path)
        return attrs != -1 and attrs & 2  # FILE_ATTRIBUTE_HIDDEN = 0x00000002
    except Exception as e:
        print(f"检查是否为隐藏文件错误：{e}")
        return False


def get_files(path: Path) -> List[str]:
    """获取目录中排序后的文件和子目录"""
    try:
        """获取目录下所有文件（包括子文件）"""
        # 符号链接，避免循环
        if path.is_symlink():
            return []

        items = list(path.iterdir())
        print(items)
        all_files = []

        # 分离文件和目录
        dirs = []
        files = []
        for item in items:
            if item.is_dir():
                dirs.append(item)
            elif item.is_file() and not is_file_hidden(str(item)):
                files.append(item)

        # 按 Windows 习惯排序
        files = os_sorted(files)
        dirs = os_sorted(dirs)

        # 添加当前目录的文件
        all_files.extend(files)

        # 递归处理子目录
        for d in dirs:
            all_files.extend(get_files(d))

        return all_files

    except PermissionError as e:
        print(f"目录无权限：{e}")
        return []
    except Exception as e:
        print(f"获取目录内容错误：{e}")
        return []


def add_numbered_prefix(file_paths: List[str]) -> List[str]:
    """为文件列表添加编号前缀"""
    processed_files = []
    for idx, file_path in enumerate(file_paths, start=1):
        dir_path = os.path.dirname(file_path)
        file_name = os.path.basename(file_path)
        new_name = f"{idx:03}--{file_name}"
        processed_files.append(os.path.join(dir_path, new_name))
    return processed_files


def remove_numbered_prefix(file_paths: List[str]) -> List[str]:
    """移除文件列表中的编号前缀"""
    processed_files = []
    prefix_pattern = re.compile(r'^\d{3}--')

    for file_path in file_paths:
        dir_path = os.path.dirname(file_path)
        file_name = os.path.basename(file_path)

        if prefix_pattern.match(file_name):
            new_name = file_name[5:]
            processed_files.append(os.path.join(dir_path, new_name))
        else:
            processed_files.append(file_path)  # 保留未编号的文件

    return processed_files


def rename_files(original_paths: List[str], new_paths: List[str]) -> bool:
    """批量重命名文件"""
    # 原始文件名列表数量与新文件名列表数量是否一致
    if len(original_paths) != len(new_paths):
        return False

    success = True
    for orig, new in zip(original_paths, new_paths):
        try:
            if orig != new:  # 避免不必要的重命名
                os.rename(orig, new)
        except OSError as e:
            print(f"重命名失败：{orig} -> {new}，错误：{e}")
            success = False

    return success


def select_directory() -> Optional[Path]:
    """选择要处理的目录"""

    dir_path = filedialog.askdirectory(
        title="选择要处理的目录",
        mustexist=True
    )

    return Path(dir_path) if dir_path else None


def process_selected_directory(
        operation_name: str,
        file_processor: Callable[[List[str]], List[str]],
        success_msg: str,
        update_info: Callable[[str], None]
):
    """处理用户选择的目录"""
    # 获取用户选择
    selected_dir = select_directory()

    if not selected_dir:
        update_info("未选择目录")
        return

    update_info(f"请稍后...")

    # 获取目录中所有文件
    files = get_files(selected_dir)

    if not files:
        update_info(f"目录中没有可处理的文件：{selected_dir}")
        return

    # 处理文件
    processed_files = file_processor(files)
    success = rename_files(files, processed_files)

    if success:
        update_info(f"{success_msg} ({len(files)}个文件)")
    else:
        update_info(f"{operation_name}部分失败，请检查文件")


# 公共接口函数
def batch_add_prefix_numbers(update_info: Callable[[str], None]) -> None:
    """添加编号"""
    process_selected_directory(
        operation_name="添加编号",
        file_processor=add_numbered_prefix,
        success_msg="已为目录下所有文件添加编号",
        update_info=update_info)


def batch_remove_prefix_numbers(update_info: Callable[[str], None]) -> None:
    """删除编号"""
    process_selected_directory(
        operation_name="删除编号",
        file_processor=remove_numbered_prefix,
        success_msg="已删除目录下所有文件的编号",
        update_info=update_info
    )
