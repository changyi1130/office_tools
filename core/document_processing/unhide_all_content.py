"""取消 Word 中所有隐藏"""
from pathlib import Path
from typing import Callable

import win32com.client as win32

from core.utils.WordAppManager import WordAppManager
from core.utils.exceptions import DocumentProcessingError
from core.utils.open_file_dialog import open_file_dialog
from core.utils.extract_path_components import extract_path_components

def unhide_all_content(document):
    """文档内容取消隐藏"""
    document.Content.Font.Hidden = False

def process_word_file(document, original_path: Path) -> str:
    """
    处理、保存文件

    :param document: Word 文档对象
    :param original_path: 原始文件路径
    :return: 处理后的文件路径
    """

    # 调用取消隐藏
    unhide_all_content(document)

    # 创建输出目录路径
    output_dir = original_path.parent / "取消隐藏"
    output_dir.mkdir(parents=True, exist_ok=True)

    # 生成输文件名
    new_filename = f"{original_path.stem}-取消隐藏{original_path.suffix}"
    output_path = output_dir / new_filename

    # 保存处理后的文档
    document.SaveAs2(FileName=str(output_path))
    return str(output_path)

def execute_unhide_workflow(update_info: Callable[[str], None]):
    """
    主处理函数

    :param update_info: 状态更新回调函数
    """
    try:
        # 选择Word文档
        update_info("请选择 Word 文档...")
        file_paths = open_file_dialog(
            "选择 Word 文档",
            file_filter=[("Word 文档", "*.doc*")],
            multi_select=True
        )

        if not file_paths:
            if update_info:
                update_info("已取消选择文件")
            return

        # 转换路径对象
        file_paths = [Path(f) for f in file_paths]

        # 提示信息
        if update_info:
            update_info(f"开始处理 {len(file_paths)} 个文件...")

        with WordAppManager() as word_app:
            # 打开文档
            for file_path in file_paths:
                document = word_app.Documents.Open(str(file_path))

                try:
                    # 处理
                    process_word_file(document, original_path=file_path)

                    # 保存
                    document.Close()
                    update_info(extract_path_components(str(file_path), 'full_name'))

                except Exception as e:
                    if document:
                        document.Close(SaveChanges=False)
                    update_info(f"处理失败 {file_path.name}: {str(e)}")

                # finally:
                #     # 确保文档关闭
                #     document.Close(SaveChanges=False)

        # 提示信息
        if update_info:
            update_info(f"所有文件已取消隐藏。")

    except DocumentProcessingError as e:
        update_info(f"处理失败: {str(e)}")
    except Exception as e:
        update_info(f"发生意外错误: {str(e)}")