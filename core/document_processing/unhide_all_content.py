"""取消 Word 中所有隐藏"""
from pathlib import Path
from typing import Callable

import win32com.client as win32

from core.services.logger_service import setup_logger
from core.utils.WordAppManager import WordAppManager
from core.utils.exceptions import DocumentProcessingError
from core.utils.open_file_dialog import open_file_dialog
from core.utils.extract_path_components import extract_path_components

# 设置日志记录器
logger = setup_logger('unhide_all_content', 'unhide_all_content.log')

def unhide_all_content(document):
    """文档内容取消隐藏"""
    logger.info("开始取消文档隐藏内容")
    document.Content.Font.Hidden = False
    logger.info("已成功取消文档隐藏内容")

def process_word_file(document, original_path: Path) -> str:
    """
    处理、保存文件

    :param document: Word 文档对象
    :param original_path: 原始文件路径
    :return: 处理后的文件路径
    """
    logger.info(f"开始处理文件: {original_path}")

    # 调用取消隐藏
    unhide_all_content(document)

    # 创建输出目录路径
    output_dir = original_path.parent / "取消隐藏"
    output_dir.mkdir(parents=True, exist_ok=True)
    logger.debug(f"创建输出目录: {output_dir}")

    # 生成输文件名
    new_filename = f"{original_path.stem}-取消隐藏{original_path.suffix}"
    output_path = output_dir / new_filename

    # 保存处理后的文档
    document.SaveAs2(FileName=str(output_path))
    logger.info(f"文件已保存到: {output_path}")
    return str(output_path)

def execute_unhide_workflow(progress_callback: Callable[..., None] = None):
    """
    主处理函数

    :param progress_callback: 进度更新回调函数 (current, total, message)
    """
    logger.info("开始执行取消隐藏工作流")
    try:
        # 选择Word文档
        if progress_callback:
            progress_callback(message="请选择 Word 文档...")

        file_paths = open_file_dialog(
            "选择 Word 文档",
            file_filter=[("Word 文档", "*.doc*")],
            multi_select=True
        )

        if not file_paths:
            logger.info("用户取消了文件选择")
            if progress_callback:
                progress_callback(message="已取消选择文件")
            return

        # 转换路径对象
        file_paths = [Path(f) for f in file_paths]
        total = len(file_paths)
        logger.info(f"已选择 {total} 个文件进行处理")

        # 提示信息
        if progress_callback:
            progress_callback(message=f"开始处理 {total} 个文件...")

        with WordAppManager() as word_app:
            # 打开文档
            for idx, file_path in enumerate(file_paths):
                logger.info(f"正在处理文件 ({idx+1}/{total}): {file_path.name}")
                document = word_app.Documents.Open(str(file_path))

                try:
                    # 处理
                    process_word_file(document, original_path=file_path)

                    # 保存
                    document.Close()

                    # 更新进度
                    if progress_callback:
                        progress_callback(idx + 1, total, extract_path_components(str(file_path), 'full_name'))

                except Exception as e:
                    error_msg = f"处理失败 {file_path.name}: {str(e)}"
                    logger.error(error_msg, exc_info=True)
                    if document:
                        document.Close(SaveChanges=False)

                    if progress_callback:
                        progress_callback(idx + 1, total, error_msg)
                    else:
                        print(error_msg)

        # 提示信息
        logger.info("所有文件已取消隐藏")
        if progress_callback:
            progress_callback(total, total, "所有文件已取消隐藏。")

    except DocumentProcessingError as e:
        logger.error(f"文档处理错误: {str(e)}", exc_info=True)
        if progress_callback:
            progress_callback(message=f"处理失败：{str(e)}")
    except Exception as e:
        logger.error(f"未知错误: {str(e)}", exc_info=True)
        if progress_callback:
            progress_callback(message=f"处理失败：{str(e)}")
