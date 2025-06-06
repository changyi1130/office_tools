"""高效批量文件页数统计工具"""

import pymupdf
from pathlib import Path
from typing import Dict, List, Optional, Callable
import win32com.client as win32
from core.utils.write_text_to_file import write_text_to_file
from core.utils.exceptions import DocumentProcessingError
from core.utils.open_file_dialog import open_file_dialog


class FileType:
    """文件类型分类器"""

    @staticmethod
    def get_type(file_path: Path) -> str:
        ext = file_path.suffix.lower()
        if ext == '.pdf':
            return 'pdf'
        elif ext in ('.doc', '.docx'):
            return 'word'
        elif ext in ('.xls', '.xlsx'):
            return 'excel'
        elif ext in ('.ppt', '.pptx'):
            return 'ppt'
        else:
            return 'unsupported'


class OfficeAppManager:
    """Office 应用程序管理器"""

    def __init__(self):
        self.word_app = None
        self.ppt_app = None

    def get_word_app(self) -> win32.CDispatch:
        """获取或创建 Word 应用实例"""
        if not self.word_app:
            self.word_app = win32.gencache.EnsureDispatch("Word.Application")
            self.word_app.Visible = False
        return self.word_app

    def get_ppt_app(self) -> win32.CDispatch:
        """获取或创建 PowerPoint 应用实例"""
        if not self.ppt_app:
            self.ppt_app = win32.gencache.EnsureDispatch("PowerPoint.Application")
        return self.ppt_app

    def close_all(self):
        """关闭所有 Office 应用程序并释放资源"""
        if self.word_app:
            try:
                # 关闭所有打开的 Word 文档
                while self.word_app.Documents.Count > 0:
                    doc = self.word_app.Documents(1)
                    doc.Close(SaveChanges=False)

                # 退出 Word 应用
                self.word_app.Quit()
            except Exception as e:
                print(f"关闭 Word 失败: {str(e)}")
            finally:
                self.word_app = None

        if self.ppt_app:
            try:
                # 关闭所有打开的 PPT 文档
                while self.ppt_app.Presentations.Count > 0:
                    pres = self.ppt_app.Presentations(1)
                    pres.Close()

                # 退出 PowerPoint 应用
                self.ppt_app.Quit()
            except Exception as e:
                print(f"关闭 PPT 失败: {str(e)}")
            finally:
                self.ppt_app = None


class PageCountResult:
    """页数统计结果封装类"""

    def __init__(self,
                 file_path: Path,
                 page_count: Optional[int] = None,
                 error: Optional[str] = None):
        self.file_path = file_path
        self.page_count = page_count
        self.error = error

    def __str__(self) -> str:
        if self.error:
            return f"{self.file_path.name}\t错误: {self.error}"
        return f"{self.file_path.name}\t{self.page_count}"


def count_pdf_pages(file_path: Path) -> PageCountResult:
    """统计PDF文件页数"""
    try:
        with pymupdf.open(file_path) as doc:
            return PageCountResult(file_path, len(doc))
    except Exception as e:
        return PageCountResult(file_path, error=f"PDF处理错误: {str(e)}")


def count_word_pages(word_app: win32.CDispatch, file_path: Path) -> PageCountResult:
    """使用已打开的Word应用统计页数"""
    try:
        doc = word_app.Documents.Open(str(file_path))
        page_count = doc.ComputeStatistics(Statistic=2)  # wdStatisticPages
        doc.Close(SaveChanges=False)
        return PageCountResult(file_path, page_count)
    except Exception as e:
        return PageCountResult(file_path, error=f"Word处理错误: {str(e)}")


def count_ppt_pages(ppt_app: win32.CDispatch, file_path: Path, include_hidden: bool = False) -> PageCountResult:
    """使用已打开的PPT应用统计页数（默认不包含隐藏页面）"""
    try:
        pres = ppt_app.Presentations.Open(str(file_path), WithWindow=False)

        # 统计可见页数
        visible_slides = 0
        for slide in pres.Slides:
            if include_hidden or not slide.SlideShowTransition.Hidden:
                visible_slides += 1

        pres.Close()
        return PageCountResult(file_path, visible_slides)
    except Exception as e:
        return PageCountResult(file_path, error=f"PPT处理错误: {str(e)}")


def count_excel_pages(file_path: Path) -> PageCountResult:
    """Excel文件不统计页数，返回特殊标记"""
    return PageCountResult(file_path, page_count=None, error="Excel文件不统计页数")


def batch_count_file_pages(
        file_paths: List[Path],
        update_info: Optional[Callable[[str], None]] = None
) -> List[PageCountResult]:
    """高效批量统计文件页数"""
    results = []
    app_manager = OfficeAppManager()
    total_files = len(file_paths)

    try:
        # 按文件类型分组
        file_groups = {}
        for file_path in file_paths:
            file_type = FileType.get_type(file_path)
            if file_type not in file_groups:
                file_groups[file_type] = []
            file_groups[file_type].append(file_path)

        # 处理Word文件
        if 'word' in file_groups:
            word_app = app_manager.get_word_app()
            for i, file_path in enumerate(file_groups['word'], 1):
                if update_info:
                    update_info(f"处理Word文件 ({i}/{len(file_groups['word'])}): {file_path.name}")
                results.append(count_word_pages(word_app, file_path))

        # 处理PPT文件
        if 'ppt' in file_groups:
            ppt_app = app_manager.get_ppt_app()
            for i, file_path in enumerate(file_groups['ppt'], 1):
                if update_info:
                    update_info(f"处理PPT文件 ({i}/{len(file_groups['ppt'])}): {file_path.name}")
                results.append(count_ppt_pages(ppt_app, file_path))

        # 处理PDF文件
        if 'pdf' in file_groups:
            for i, file_path in enumerate(file_groups['pdf'], 1):
                if update_info:
                    update_info(f"处理PDF文件 ({i}/{len(file_groups['pdf'])}): {file_path.name}")
                results.append(count_pdf_pages(file_path))

        # 处理Excel文件
        if 'excel' in file_groups:
            for i, file_path in enumerate(file_groups['excel'], 1):
                if update_info:
                    update_info(f"跳过Excel文件 ({i}/{len(file_groups['excel'])}): {file_path.name}")
                results.append(count_excel_pages(file_path))

        # 处理不支持的文件
        if 'unsupported' in file_groups:
            for i, file_path in enumerate(file_groups['unsupported'], 1):
                if update_info:
                    update_info(f"不支持的文件类型 ({i}/{len(file_groups['unsupported'])}): {file_path.name}")
                results.append(PageCountResult(file_path, error="不支持的文件类型"))

    except Exception as e:
        # 整体错误处理
        error_result = PageCountResult(Path(""), error=f"批量处理失败: {str(e)}")
        results = [error_result] * total_files

    finally:
        # 确保释放所有资源
        app_manager.close_all()

    return results


def generate_page_count_report(
        results: List[PageCountResult],
        output_dir: Path
) -> Path:
    """生成页数统计报告"""
    # 准备报告内容
    report_lines = ["文件\t页数\t状态"]
    for result in results:
        report_lines.append(str(result))

    # 生成报告文件路径
    report_path = output_dir / "文件页数统计报告.txt"

    # 写入文件
    write_text("\n".join(report_lines), report_path)
    return report_path


def process_page_count_collection(
        update_info: Optional[Callable[[str], None]] = None
) -> None:
    """带UI的页数统计流程"""
    try:
        # 选择文件
        if update_info:
            update_info("请选择要统计的文件...")

        file_paths = open_multiple_files(
            "选择文件",
            file_filter=[
                ("PDF文件", "*.pdf"),
                ("Word文档", "*.doc*"),
                ("PPT演示文稿", "*.ppt*"),
                ("Excel文件", "*.xls*"),
                ("所有文件", "*.*")
            ]
        )

        if not file_paths:
            if update_info:
                update_status("已取消选择文件")
            return

        # 转换路径对象
        file_paths = [Path(f) for f in file_paths]
        output_dir = file_paths[0].parent

        # 批量统计
        if update_info:
            update_info(f"开始统计 {len(file_paths)} 个文件的页数...")

        results = batch_count_file_pages(file_paths, update_info)

        # 生成报告
        report_path = generate_page_count_report(results, output_dir)

        if update_info:
            update_info(f"统计完成！报告已保存至:\n{report_path}")

    except Exception as e:
        if update_info:
            update_info(f"统计失败: {str(e)}")
