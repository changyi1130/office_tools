
"""
隐藏Word文档中不包含中文的段落
"""
from pathlib import Path
from typing import Callable

import win32com.client as win32

from core.services.logger_service import setup_logger
from core.utils.WordAppManager import WordAppManager
from core.utils.exceptions import DocumentProcessingError
from core.utils.open_file_dialog import open_file_dialog
from core.utils.run_vba_macro import execute_vba_on_document
from core.utils.extract_path_components import extract_path_components

# 设置日志记录器
logger = setup_logger('hide_non_chinese_paragraphs', 'hide_non_chinese_paragraphs.log')

# VBA宏名称
MACRO_NAME = "hideNonChineseParagraphs.hideNonChineseParagraphs"


def select_word_files() -> list[Path]:
    """
    选择要处理的Word文档

    :return: 选择的文件路径列表
    """
    logger.info("开始选择Word文档")

    file_paths = open_file_dialog(
        "选择Word文档",
        file_filter=[("Word 文档", "*.doc*")],
        multi_select=True
    )

    if not file_paths:
        logger.info("用户取消了文件选择")
        return []

    # 转换为Path对象
    files = [Path(f) for f in file_paths]
    logger.info(f"已选择 {len(files)} 个文件")

    return files


def open_document(word_app, file_path: Path):
    """
    打开Word文档

    :param word_app: Word应用程序对象
    :param file_path: 文档路径
    :return: 文档对象
    """
    logger.info(f"正在打开文档: {file_path.name}")
    document = word_app.Documents.Open(str(file_path))
    logger.debug(f"文档已打开: {file_path.name}")
    return document


def execute_hide_macro(document):
    """
    执行隐藏非中文段落的VBA宏

    :param document: Word文档对象
    """
    logger.info("开始执行隐藏非中文段落宏")

    try:
        # 使用工具函数执行宏
        execute_vba_on_document(document, MACRO_NAME)
        logger.info("VBA宏执行成功")
    except Exception as e:
        logger.error(f"执行VBA宏时出错: {str(e)}", exc_info=True)
        raise DocumentProcessingError(f"执行VBA宏失败: {str(e)}")


def save_document(document, original_path: Path) -> str:
    """
    保存处理后的文档

    :param document: Word文档对象
    :param original_path: 原始文件路径
    :return: 保存后的文件路径
    """
    logger.info(f"开始保存文档: {original_path.name}")

    # 创建输出目录
    output_dir = original_path.parent / "中文预编"
    output_dir.mkdir(parents=True, exist_ok=True)
    logger.debug(f"创建输出目录: {output_dir}")

    # 生成新文件名
    new_filename = f"{original_path.stem}-中文预编{original_path.suffix}"
    output_path = output_dir / new_filename

    # 保存文档
    document.SaveAs2(FileName=str(output_path))
    logger.info(f"文档已保存到: {output_path}")

    return str(output_path)


def close_document(document, save_changes: bool = False):
    """
    关闭Word文档

    :param document: Word文档对象
    :param save_changes: 是否保存更改
    """
    logger.debug("关闭文档")
    document.Close(SaveChanges=save_changes)


def process_single_file(word_app, file_path: Path, progress_callback: Callable = None) -> str:
    """
    处理单个Word文档

    :param word_app: Word应用程序对象
    :param file_path: 文档路径
    :param progress_callback: 进度回调函数
    :return: 处理后的文件路径
    """
    document = None
    try:
        # 打开文档
        document = open_document(word_app, file_path)

        # 执行隐藏宏
        execute_hide_macro(document)

        # 保存文档
        output_path = save_document(document, file_path)

        return output_path

    except Exception as e:
        error_msg = f"处理文件失败 {file_path.name}: {str(e)}"
        logger.error(error_msg, exc_info=True)
        raise DocumentProcessingError(error_msg)
    finally:
        # 关闭文档
        if document:
            close_document(document, save_changes=False)


def execute_hide_non_chinese_workflow(progress_callback: Callable = None):
    """
    执行隐藏非中文段落的主工作流

    :param progress_callback: 进度更新回调函数 (current, total, message)
    """
    logger.info("开始执行隐藏非中文段落工作流")

    try:
        # 选择文件
        if progress_callback:
            progress_callback(message="请选择Word文档...")

        file_paths = select_word_files()

        if not file_paths:
            if progress_callback:
                progress_callback(message="已取消选择文件")
            return

        total = len(file_paths)
        logger.info(f"开始处理 {total} 个文件")

        if progress_callback:
            progress_callback(message=f"开始处理 {total} 个文件...")

        # 处理每个文件
        with WordAppManager() as word_app:
            for idx, file_path in enumerate(file_paths):
                try:
                    logger.info(f"正在处理文件 ({idx+1}/{total}): {file_path.name}")

                    # 处理文件
                    output_path = process_single_file(word_app, file_path, progress_callback)

                    # 更新进度
                    if progress_callback:
                        progress_callback(
                            idx + 1, 
                            total, 
                            extract_path_components(str(file_path), 'full_name')
                        )

                    logger.info(f"文件处理完成: {file_path.name}")

                except Exception as e:
                    error_msg = f"处理失败 {file_path.name}: {str(e)}"
                    logger.error(error_msg, exc_info=True)

                    if progress_callback:
                        progress_callback(idx + 1, total, error_msg)
                    else:
                        print(error_msg)

        # 完成提示
        logger.info("所有文件处理完成")
        if progress_callback:
            progress_callback(total, total, "所有文件处理完成，等待进程结束...")

    except DocumentProcessingError as e:
        logger.error(f"文档处理错误: {str(e)}", exc_info=True)
        if progress_callback:
            progress_callback(message=f"处理失败：{str(e)}")
    except Exception as e:
        logger.error(f"未知错误: {str(e)}", exc_info=True)
        if progress_callback:
            progress_callback(message=f"处理失败：{str(e)}")
