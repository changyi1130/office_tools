"""PPT转PDF工具"""

from pathlib import Path
from typing import Callable

from core.services.logger_service import setup_logger
from core.utils.open_file_dialog import open_file_dialog
from core.utils.exceptions import DocumentConversionError

logger = setup_logger('convert_ppt_to_pdf', 'convert_ppt_to_pdf.log')


def convert_ppt_to_pdf(
        progress_callback: Callable[..., None] = None
) -> None:
    """
    将 PowerPoint 演示文稿转换为 PDF 文件

    使用 PowerPoint COM 自动化打开 PPT/PPTX 文件，另存为 PDF。
    如需将 PDF 进一步转换为 Word，请使用 Adobe Acrobat 手动操作。

    :param progress_callback: 进度更新回调函数
    """
    # 选择文件
    file_paths = open_file_dialog(
        "选择要转换的 PPT 文件",
        file_filter=[
            ("PowerPoint 演示文稿", "*.pptx"),
            ("PowerPoint 97-2003 演示文稿", "*.ppt")
        ],
        multi_select=True
    )

    if not file_paths:
        if progress_callback:
            progress_callback(message="未选择文件")
        logger.info("未选择文件")
        return

    total_files = len(file_paths)
    if progress_callback:
        progress_callback(message=f"开始转换 {total_files} 个 PPT 文件...")
    logger.info(f"开始转换 {total_files} 个 PPT 文件...")

    from win32com.client import Dispatch
    success_count = 0
    ppt = Dispatch("PowerPoint.Application")
    try:
        for i, file_path in enumerate(file_paths, 1):
            file_path = Path(file_path)
            try:
                if progress_callback:
                    progress_callback(i, total_files, f"正在处理：{file_path.name}")

                pres = ppt.Presentations.Open(str(file_path), WithWindow=False)
                try:
                    pdf_path = file_path.with_suffix(".pdf")
                    pres.SaveAs(str(pdf_path), 32)  # ppSaveAsPDF
                finally:
                    pres.Close()

                success_count += 1
                if progress_callback:
                    progress_callback(i, total_files, f"已完成：{file_path.name} → {pdf_path.name}")
                logger.info(f"已转换：{file_path.name} → {pdf_path.name}")

            except Exception as e:
                error_msg = f"转换失败：{file_path.name} - {str(e)}"
                if progress_callback:
                    progress_callback(i, total_files, error_msg)
                logger.error(f"转换失败：{file_path}: {str(e)}")
    finally:
        ppt.Quit()

    # 最终状态报告
    result_msg = f"转换完成：{success_count} / {total_files} 成功"
    if success_count < total_files:
        result_msg += f"，{total_files - success_count} 个文件失败"

    if progress_callback:
        progress_callback(total_files, total_files, result_msg)

    # 提示 PDF→Word 需要手动操作
    if success_count > 0:
        manual_hint = (
            "提示：如需继续将 PDF 转换为 Word，\n"
            "请用 Adobe Acrobat 手动操作：\n"
            "  1. 用 Acrobat 打开生成的 PDF\n"
            "  2. 文件 → 导出到 → Microsoft Word\n"
            "  3. 另存为 .docx 文件"
        )
        if progress_callback:
            progress_callback(message=manual_hint)
        logger.info("提示用户 PDF→Word 需手动操作")

    logger.info(f"转换完成：{success_count}/{total_files} 成功")
