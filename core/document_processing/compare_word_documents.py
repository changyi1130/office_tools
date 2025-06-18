"""Word 文档比较功能"""

from pathlib import Path
from typing import Callable

from core.utils.WordAppManager import WordAppManager
from core.utils.exceptions import DocumentComparisonError  # 自定义异常
from core.utils.open_file_dialog import open_file_dialog  # 假设已优化文件对话框


def compare_word_documents(original_path: Path, modified_path: Path) -> Path:
    """
    比较两个Word文档并生成比较结果

    :param original_path: 原始文档路径
    :param modified_path: 修改后文档路径
    :return: 生成的比较文档路径
    :raises DocumentComparisonError: 文档比较失败时抛出
    """
    # 验证文件存在
    if not original_path.exists():
        raise FileNotFoundError(f"原始文档不存在: {original_path}")
    if not modified_path.exists():
        raise FileNotFoundError(f"修改后文档不存在: {modified_path}")

    # 生成比较文档路径
    output_path = modified_path.with_stem(f"{modified_path.stem}——比较文档")

    try:
        with WordAppManager() as word_app:
            # 打开文档
            original_doc = word_app.Documents.Open(str(original_path))
            modified_doc = word_app.Documents.Open(str(modified_path))

            # 比较文档
            comparison_doc = word_app.CompareDocuments(original_doc, modified_doc)

            # 保存比较结果
            comparison_doc.SaveAs2(FileName=str(output_path))

            # 关闭文档（不保存原始修改）
            original_doc.Close(SaveChanges=False)
            modified_doc.Close(SaveChanges=False)
            comparison_doc.Close(SaveChanges=True)

            return output_path

    except Exception as e:
        # 封装具体异常信息
        raise DocumentComparisonError(f"文档比较失败: {str(e)}") from e


def compare_documents_with_ui(update_info: Callable[[str], None]):
    """带用户界面的文档比较流程"""
    try:
        # 选择原始文档
        update_info("请选择原始文档...")
        original_path = open_file_dialog("选择原始文档", file_filter=[("Word文档", "*.doc*")])
        if not original_path:
            update_info("已取消选择原始文档")
            return

        # 选择修改后文档
        update_info("请选择修改后文档...")
        modified_path = open_file_dialog("选择修改后文档", file_filter=[("Word文档", "*.doc*")])
        if not modified_path:
            update_info("已取消选择修改后文档")
            return

        # 验证不同文件
        if Path(original_path) == Path(modified_path):
            update_info("错误: 不能选择同一个文件作为原始和修改文档")
            return

        # 执行比较
        update_info("正在比较文档...")
        result_path = compare_word_documents(Path(original_path), Path(modified_path))

        # 显示结果
        update_info(f"比较完成: 结果已保存至\n{result_path}")

    except DocumentComparisonError as e:
        update_info(f"比较失败: {str(e)}")
    except Exception as e:
        update_info(f"发生意外错误: {str(e)}")
