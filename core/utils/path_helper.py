import os
import sys
from pathlib import Path

def get_resource_path(relative_path: str) -> Path:
    """
    获取资源文件的绝对路径。兼容开发环境和PyInstaller打包后的单文件exe环境。

    :param relative_path: 资源文件相对于项目根目录的路径，
                          例如 "core/vba_libs/word_tools.dotm"
    :return: Path对象，指向资源的真实绝对路径。
    :raises FileNotFoundError: 如果资源文件不存在。
    """
    try:
        # PyInstaller创建单文件exe后，会生成一个临时文件夹，路径存储在 _MEIPASS 中
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        # 未打包时（开发环境），使用当前文件的父级目录向上回溯到项目根目录
        # 假设此文件在 `项目根目录/core/utils/path_helper.py`
        base_path = Path(__file__).parent.parent.parent  # 回溯到项目根目录

    # 拼接并返回资源的绝对路径
    full_path = (base_path / relative_path).resolve()

    # 检查文件是否存在，提供友好错误信息
    if not full_path.exists():
        raise FileNotFoundError(
            f"资源文件未找到：{full_path}\n"
            f"请检查相对路径是否正确：{relative_path}\n"
            f"当前基路径：{base_path}"
        )
    return full_path