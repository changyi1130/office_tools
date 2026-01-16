"""Word 文档信息统计工具"""

from enum import IntEnum
from pathlib import Path
from typing import Callable

import win32com.client as win32
from natsort import os_sorted

from core.utils.CountResult import CountResult
from core.utils.WordAppManager import WordAppManager
from core.utils.exceptions import DocumentProcessingError
from core.utils.open_file_dialog import open_file_dialog
from core.utils.write_report_to_excel import write_report_to_excel


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


def get_document_statistics(
        document: win32.CDispatch,
        statistic_type: WordStatisticType,
        include_notes: bool = True
) -> int:
    """
    获取 Word 文档的统计信息

    :param document: Word 文档对象
    :param statistic_type: 统计类型
    :param include_notes: 是否包含页眉、页脚和尾注
    :return: 统计结果
    :raises DocumentProcessingError: 统计失败时抛出
    """
    try:
        # 确保显示最终状态（不显示修订标记）
        document.ShowRevisions = False

        # 获取统计信息
        return document.ComputeStatistics(
            Statistic=statistic_type,
            IncludeFootnotesAndEndnotes=include_notes
        )
    except Exception as e:
        raise DocumentProcessingError(f"统计信息失败: {str(e)}") from e


def process_word_statistics(
        statistic_type: WordStatisticType = WordStatisticType.PAGES,
        include_notes: bool = True,
        update_info: Callable[[str], None] = None
) -> None:
    """
    主处理函数：收集 Word 文档统计信息

    :param statistic_type: 统计类型（默认页数）
    :param include_notes: 是否包含页眉页脚（默认 True）
    :param update_info: 状态更新回调
    """
    try:
        # 选择Word文档
        update_info("请选择要统计的 Word 文档...")
        file_paths = open_file_dialog(
            "选择 Word 文档",
            file_filter=[("Word 文档", "*.doc*")],
            multi_select=True
        )

        if not file_paths:
            update_info("已取消选择文档")
            return

        # 准备处理
        total_files = len(file_paths)
        results = []
        output_dir = Path(file_paths[0]).parent
        stat_desc = STATISTIC_DESCRIPTIONS[statistic_type]

        update_info(f"开始统计 {total_files} 个文档的{stat_desc}...")

        # 处理每个文档
        with WordAppManager() as word_app:
            for i, file_path in enumerate(file_paths, 1):
                file_path = Path(file_path)
                update_info(f"处理中：{i} / {total_files}")

                try:
                    # 打开文档
                    doc = word_app.Documents.Open(str(file_path))

                    # 获取统计信息
                    count = get_document_statistics(doc, statistic_type, include_notes)
                    results.append(CountResult(file_path=file_path, page_count=count))
                    # results.append(f"{file_path.name}\t{count}")

                except DocumentProcessingError as e:
                    # 记录错误但继续处理其他文件
                    results.append(CountResult(file_path=file_path, error=f"错误: {str(e)}"))
                    # results.append(f"{file_path.name}\t错误: {str(e)}")
                finally:
                    # 确保文档关闭
                    if 'doc' in locals():
                        doc.Close(SaveChanges=False)

        # 处理结果
        result_data = [result.to_row() for result in results]
        result_data = os_sorted(result_data, key=lambda r: r[0])

        # 生成报告
        report_name = f"000--文档统计-{stat_desc}.xlsx"
        report_path = output_dir / report_name
        column_headers = ['文件名称', stat_desc]
        # write_text_to_file(contents=results, target_path=report_path)
        write_report_to_excel(report_data=result_data,
                              column_headers=column_headers,
                              output_path=report_path)

        update_info(f"统计完成！报告已保存至:\n{report_path}")

    except Exception as e:
        update_info(f"统计失败: {str(e)}")
