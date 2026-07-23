"""Word 文档信息统计工具"""

from __future__ import annotations

from enum import IntEnum
from pathlib import Path
from typing import Callable, TYPE_CHECKING

from natsort import os_sorted

if TYPE_CHECKING:
    import win32com.client as win32

from core.utils.CountResult import CountResult
from core.utils.WordAppManager import WordAppManager
from core.utils.exceptions import DocumentProcessingError
from core.utils.open_file_dialog import open_file_dialog
from core.utils.write_report_to_excel import write_report_to_excel
from core.utils.format_excel import format_excel_file


# 定义统计类型枚举（更清晰）
class WordStatisticType(IntEnum):
    WORDS = 0  # 字数
    LINES = 1  # 行数
    PAGES = 2  # 页数
    CHARACTERS_NO_SPACES = 3  # 字符数(不计空格)
    PARAGRAPHS = 4  # 段落数
    CHARACTERS_WITH_SPACES = 5  # 字符数(计空格)
    FAR_EAST_CHARACTERS = 6  # 中文字符和朝鲜语单词


# 统计类型描述映射
STATISTIC_DESCRIPTIONS = {
    WordStatisticType.WORDS: "字数",
    WordStatisticType.LINES: "行数",
    WordStatisticType.PAGES: "页数",
    WordStatisticType.CHARACTERS_NO_SPACES: "字符数(不计空格)",
    WordStatisticType.PARAGRAPHS: "段落数",
    WordStatisticType.CHARACTERS_WITH_SPACES: "字符数(计空格)",
    WordStatisticType.FAR_EAST_CHARACTERS: "中文字符和朝鲜语单词"
}

# 页眉类型
class HeaderFooterType(IntEnum):
    wdHeaderFooterPrimary = 1     # 主页眉/页脚（奇数页）
    wdHeaderFooterFirstPage = 2   # 首页页眉/页脚
    wdHeaderFooterEvenPages = 3   # 偶数页页眉/页脚


def _count_shapes(shapes, statistic_type: WordStatisticType) -> int:
    """
    递归统计形状集合中所有文字的统计量
    对应 VBA: CountShapesText（只处理 TextFrame，递归处理组合形状 GroupItems）
    注意：HeaderFooter.Shapes 返回的是全局集合，此函数应只调用一次。
    """
    total = 0
    for shape in shapes:
        try:
            if shape.TextFrame.HasText:
                total += shape.TextFrame.TextRange.ComputeStatistics(statistic_type)
        except:
            pass

        # 递归处理组合形状（msoGroup = 6）
        try:
            if shape.Type == 6:
                total += _count_shapes(shape.GroupItems, statistic_type)
        except:
            pass
    return total


def get_document_statistics(
        document: win32.CDispatch,
        statistic_type: WordStatisticType,
        include_notes: bool = True,
        include_header_footer: bool = False
) -> int:
    """
    获取 Word 文档的统计信息

    :param document: Word 文档对象
    :param statistic_type: 统计类型
    :param include_notes: 是否包含脚注和尾注
    :param include_header_footer: 是否统计页眉页脚（自动跳过链接到前一节的）
    :return: 统计结果
    :raises DocumentProcessingError: 统计失败时抛出
    """
    try:
        import win32com.client as win32

        # 确保显示最终状态（不显示修订标记）
        document.ShowRevisions = False

        # 1. 基础统计（正文 + 脚注/尾注）
        total = document.ComputeStatistics(
            Statistic=statistic_type,
            IncludeFootnotesAndEndnotes=include_notes
        )

        # 2. 页眉页脚统计（仅在开启且统计类型不是"页数"时执行）
        if include_header_footer and statistic_type != WordStatisticType.PAGES:
            # 2a. 遍历所有节，累加各未链接页眉/页脚的主体文本
            for sec in document.Sections:
                for hf in sec.Headers:
                    if hf.Exists and not hf.LinkToPrevious:
                        total += hf.Range.ComputeStatistics(statistic_type)
                for hf in sec.Footers:
                    if hf.Exists and not hf.LinkToPrevious:
                        total += hf.Range.ComputeStatistics(statistic_type)

            # 2b. 统计形状（只调用一次——HeaderFooter.Shapes 返回全局集合）
            try:
                first_header = document.Sections(1).Headers(1)
                if first_header.Exists:
                    total += _count_shapes(first_header.Shapes, statistic_type)
                else:
                    total += _count_shapes(document.Sections(1).Footers(1).Shapes, statistic_type)
            except:
                pass

        return total

    except Exception as e:
        raise DocumentProcessingError(f"统计信息失败: {str(e)}") from e


def process_word_statistics(
        progress_callback: Callable[..., None],
        statistic_type: WordStatisticType = WordStatisticType.PAGES,
        include_notes: bool = True,
        include_header_footer: bool = False
) -> None:
    """
    主处理函数：收集 Word 文档统计信息

    :param statistic_type: 统计类型（默认页数）
    :param include_notes: 是否包含脚注和尾注（默认 True）
    :param include_header_footer: 是否统计页眉页脚（默认 False）
    :param progress_callback: 状态更新回调
    """
    try:
        # 选择Word文档
        progress_callback(message="请选择要统计的 Word 文档...")
        file_paths = open_file_dialog(
            "选择 Word 文档",
            file_filter=[("Word 文档", "*.doc*")],
            multi_select=True
        )

        if not file_paths:
            progress_callback(message="已取消选择文档")
            return

        # 准备处理
        total_files = len(file_paths)
        results = []
        output_dir = Path(file_paths[0]).parent
        stat_desc = STATISTIC_DESCRIPTIONS[statistic_type]

        progress_callback(message=f"开始统计 {total_files} 个文档的{stat_desc}...")

        # 处理每个文档
        with WordAppManager() as word_app:
            for i, file_path in enumerate(file_paths, 1):
                file_path = Path(file_path)
                progress_callback(message=f"处理中：{i} / {total_files}")

                doc = None
                try:
                    # 打开文档
                    doc = word_app.Documents.Open(str(file_path))

                    # 获取统计信息
                    count = get_document_statistics(
                        doc,
                        statistic_type,
                        include_notes,
                        include_header_footer
                    )
                    results.append(CountResult(file_path=file_path, page_count=count))

                except DocumentProcessingError as e:
                    # 记录错误但继续处理其他文件
                    results.append(CountResult(file_path=file_path, error=f"错误: {str(e)}"))
                finally:
                    # 确保文档关闭
                    if doc is not None:
                        doc.Close(SaveChanges=False)

        # 处理结果
        result_data = [result.to_row() for result in results]
        result_data = os_sorted(result_data, key=lambda r: r[0])

        # 生成报告
        report_name = f"000--文档统计-{stat_desc}.xlsx"
        report_path = output_dir / report_name
        column_headers = ['文件名称', stat_desc]
        write_report_to_excel(report_data=result_data,
                              column_headers=column_headers,
                              output_path=report_path)

        # 美化 Excel 格式（字体、对齐、冻结窗格、列宽等）
        format_excel_file(str(report_path))

        progress_callback(message=f"统计完成！报告已保存至:\n{report_path}")

    except Exception as e:
        progress_callback(message=f"统计失败: {str(e)}")
