"""Word文档格式转换工具"""

from pathlib import Path
from typing import Callable

from core.utils.WordAppManager import WordAppManager
from core.utils.exceptions import DocumentConversionError
from core.utils.open_file_dialog import open_file_dialog

# 转换类型枚举
CONVERSION_TYPES = {
    "doc_to_docx": {
        "filter": [("Word 97-2003 文档", "*.doc")],
        "function": "convert_to_docx",
        "success_msg": "已转存为高版本 DOCX"
    },
    "docx_to_doc": {
        "filter": [("Word 文档", "*.docx")],
        "function": "convert_to_doc",
        "success_msg": "已转存为低版本 DOC"
    },
    "to_pdf": {
        "filter": [("Word 文档", "*.doc*")],
        "function": "convert_to_pdf",
        "success_msg": "已转存为 PDF"
    }
}


def convert_document(
        conversion_type: str,
        update_info: Callable[[str], None]
) -> None:
    """
    通用文档转换函数

    :param conversion_type: 转换类型 (doc_to_docx, docx_to_doc, to_pdf)
    :param update_info: 状态更新回调函数
    """
    # 验证转换类型
    if conversion_type not in CONVERSION_TYPES:
        raise DocumentConversionError(f"无效的转换类型: {conversion_type}")

    config = CONVERSION_TYPES[conversion_type]

    # 选择文件
    file_paths = open_file_dialog(
        f"选择要转换的文档",
        file_filter=config["filter"],
        multi_select=True
    )

    if not file_paths:
        update_info("已取消转换")
        return

    total_files = len(file_paths)
    update_info(f"开始转换 {total_files} 个文件...")

    success_count = 0
    with WordAppManager() as word_app:
        for i, file_path in enumerate(file_paths, 1):
            file_path = Path(file_path)
            try:
                # 打开文档
                doc = word_app.Documents.Open(str(file_path))

                # 执行转换
                if config["function"] == "convert_to_docx":
                    output_path = convert_to_docx(file_path, doc)
                elif config["function"] == "convert_to_doc":
                    output_path = convert_to_doc(file_path, doc)
                else:  # convert_to_pdf
                    output_path = convert_to_pdf(file_path, doc)

                # 更新状态
                success_count += 1
                update_info(f"处理中：{i} / {total_files}\n{output_path.name}")

            except Exception as e:
                error_msg = f"转换失败: {file_path.name} - {str(e)}"
                update_info(f"({i} / {total_files}) {error_msg}")
                # 记录详细错误日志
                # logger.error(f"{file_path}: {traceback.format_exc()}")

            finally:
                # 确保文档关闭
                if 'doc' in locals():
                    doc.Close(SaveChanges=False)

    # 最终状态报告
    result_msg = f"转换完成: {success_count} / {total_files} 成功"
    if success_count < total_files:
        result_msg += f"，{total_files - success_count} 个文件失败"

    update_info(result_msg)


def convert_to_docx(original_path: Path, doc) -> Path:
    """将文档转换为 DOCX 格式"""
    output_path = original_path.with_suffix(".docx")
    doc.SaveAs2(FileName=str(output_path), FileFormat=16)  # wdFormatDocumentDefault
    return output_path


def convert_to_doc(original_path: Path, doc) -> Path:
    """将文档转换为 DOC 格式"""
    output_path = original_path.with_suffix(".doc")
    doc.SaveAs2(FileName=str(output_path), FileFormat=0)  # wdFormatDocument
    return output_path


def convert_to_pdf(original_path: Path, doc) -> Path:
    """将文档转换为 PDF 格式"""
    # 隐藏修订标记
    doc.ShowRevisions = False

    output_path = original_path.with_suffix(".pdf")
    doc.SaveAs2(FileName=str(output_path), FileFormat=17)  # wdFormatPDF
    return output_path
