"""Word 文档比较功能"""

from pathlib import Path
from time import sleep
from typing import Callable

from core.services.logger_service import setup_logger
from core.utils.WordAppManager import WordAppManager
from core.utils.exceptions import DocumentComparisonError  # 自定义异常
from core.utils.open_file_dialog import open_file_dialog  # 假设已优化文件对话框

# 设置日志记录器
logger = setup_logger('compare_word_documents', 'compare_word_documents.log')

def compare_word_documents(original_path: Path, modified_path: Path) -> Path:
    """
    比较两个 Word 文档并生成比较结果

    :param original_path: 原始版本路径
    :param modified_path: 更新版本路径
    :return: 生成的比较文档路径
    :raises DocumentComparisonError: 文档比较失败时抛出
    """
    # 验证文件存在
    if not original_path.exists():
        raise FileNotFoundError(f"原始版本不存在: {original_path}")
    if not modified_path.exists():
        raise FileNotFoundError(f"更新版本不存在: {modified_path}")

    # 生成比较文档路径
    output_path = modified_path.with_stem(f"{modified_path.stem}——比较文档")

    try:
        with WordAppManager() as word_app:
            # 打开文档
            original_doc = word_app.Documents.Open(str(original_path))
            sleep(2)
            modified_doc = word_app.Documents.Open(str(modified_path))
            sleep(2)

            # 比较文档
            comparison_doc = word_app.CompareDocuments(original_doc, modified_doc)
            sleep(2)

            # 保存比较结果
            comparison_doc.SaveAs2(FileName=str(output_path))
            sleep(2)

            # 关闭文档（不保存原始修改）
            original_doc.Close(SaveChanges=False)
            modified_doc.Close(SaveChanges=False)
            comparison_doc.Close(SaveChanges=True)

            return output_path

    except Exception as e:
        # 封装具体异常信息
        raise DocumentComparisonError(f"文档比较失败: {str(e)}") from e


def compare_documents_with_ui(progress_callback: Callable[..., None] = None):
    """文档比较流程"""
    try:
        # 选择原始版本
        original_path = open_file_dialog("选择原始版本", file_filter=[("Word文档", "*.doc*")])
        if not original_path:
            progress_callback(message="已取消选择原始版本")
            logger.info("未选择原始版本文件")
            return
        else:
            progress_callback(message=f"原始版本: {original_path}")
            logger.info(f"原始版本：{original_path}")

        # 选择更新版本
        modified_path = open_file_dialog("选择更新版本", file_filter=[("Word文档", "*.doc*")])
        if not modified_path:
            progress_callback(message="已取消选择更新版本")
            logger.info("未选择更新版本文件")
            return
        else:
            progress_callback(message=f"更新版本: {modified_path}")
            logger.info(f"更新版本：{original_path}")

        # 验证不同文件
        if Path(original_path) == Path(modified_path):
            progress_callback(message="错误: 一个文件不能同时为更新版本和原始版本！")
            logger.error("原始与新版为同一文件")
            return

        # 执行比较
        progress_callback(message="正在比较文档...")
        logger.info("开始比较文档")
        result_path = compare_word_documents(Path(original_path), Path(modified_path))

        # 显示结果
        progress_callback(1, 1, message=f"比较完成: 结果已保存至\n{result_path}")
        logger.info(f"比较完成：{result_path}")

    except DocumentComparisonError as e:
        progress_callback(message=f"比较失败: {str(e)}")
        logger.error(f"比较失败: {str(e)}")
    except Exception as e:
        progress_callback(message=f"发生意外错误: {str(e)}")
        logger.error(f"发生意外错误: {str(e)}")
