"""Word文档修订内容高亮工具"""

import win32com.client as win32
from pathlib import Path
from typing import Callable
from core.utils.WordAppManager import WordAppManager
from core.utils.open_file_dialog import open_file_dialog
from core.utils.exceptions import DocumentProcessingError

# Word修订类型常量（更易读的命名）
REVISION_FORMATTING = 10  # 格式修订
REVISION_CONFLICT = 13  # 冲突修订
REVISION_INSERT = 1  # 插入内容
REVISION_DELETE = 2  # 删除内容


def highlight_revisions(document):
    """
    高亮显示文档中的非格式修订内容

    :param document: Word文档对象
    """
    # 检查文档是否有修订
    if document.Revisions.Count == 0:
        raise DocumentProcessingError("文档中没有修订内容")

    # 遍历并高亮所有非格式修订
    for revision in document.Revisions:
        # 跳过格式修订和冲突修订
        if revision.Type in (REVISION_FORMATTING, REVISION_CONFLICT):
            continue

        # 高亮显示修订内容
        revision.Range.HighlightColorIndex = win32.constants.wdYellow


def process_document_revisions(document, original_path: Path) -> Path:
    """
    处理文档修订并保存

    :param document: Word 文档对象
    :param original_path: 原始文件路径
    :return: 处理后的文件路径
    """
    # 高亮修订内容
    highlight_revisions(document)

    # 生成输出文件名
    output_path = original_path.with_stem(f"{original_path.stem}-高亮修订")

    # 保存处理后的文档
    document.SaveAs2(FileName=str(output_path))
    return output_path


def highlight_document_revisions(update_info: Callable[[str], None]):
    """
    主处理函数：高亮显示Word文档中的修订内容

    :param update_info: 状态更新回调函数
    """
    try:
        # 选择Word文档
        update_info("请选择 Word 文档...")
        file_path = open_file_dialog(
            "选择 Word 文档",
            file_filter=[("Word 文档", "*.doc*")]
        )

        if not file_path:
            update_info("已取消选择文档")
            return

        file_path = Path(file_path)
        update_info(f"处理中: {file_path.name}")

        with WordAppManager() as word_app:
            # 打开文档
            document = word_app.Documents.Open(str(file_path))

            try:
                # 处理并保存文档
                output_path = process_document_revisions(document, file_path)
                update_info(f"处理完成: 结果已保存至\n{output_path}")

            finally:
                # 确保文档关闭
                document.Close(SaveChanges=False)

    except DocumentProcessingError as e:
        update_info(f"处理失败: {str(e)}")
    except Exception as e:
        update_info(f"发生意外错误: {str(e)}")
        # 记录完整错误日志
        # logger.exception("高亮修订处理失败")
