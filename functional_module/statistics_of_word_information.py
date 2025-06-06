"""
提取 Word 文件信息

名称                               值  描述
wdStatisticWords                   0   字数
wdStatisticLines                   1   行
wdStatisticPages                   2   页数
wdStatisticCharacters              3   字符数(不计空格)
wdStatisticParagraphs              4   段落数
wdStatisticCharactersWithSpaces    5   字符数(计空格)
wdStatisticFarEastCharacters       6   中文字符和朝鲜语单词
"""

from core.utils.WordAppManager import WordAppManager
from file_processing.open_multiple_files import open_multiple_files
from core.utils.extract_path_components import extract_file_name
from core.utils.write_text_to_file import write_text

# 不知道有什么用
# global_progress = None

def statistics_of_word_information(doc, statistic_type, include_head_and_foot):
    """
    提取 Word 文档的统计信息。
    
    :param doc: Word 对象
    :param statistic_type: 统计类型
    :param include_head_and_foot: 是否包含页眉和页脚
    :return: 统计结果
    """

    # 最终状态
    doc.ShowRevisions = False

    try:
        count = doc.ComputeStatistics(Statistic=statistic_type, 
            IncludeFootnotesAndEndnotes=include_head_and_foot)
        return count
    except Exception as e:
        print(f"statistics_of_word_information: {e}")
        return '!!!' + e
    
def process_word_statistics(statistic_type, include_head_and_foot, update_info):
    """选择文档并统计 Word 信息"""

    file_paths = open_multiple_files('选择文档')

    # 检查是否选择了文件
    if file_paths is None:
        update_info("已取消")
        return None
    
    # 更新提示信息
    total_files = len(file_paths)
    current_file = 0
    update_info(f"进度：{current_file} / {total_files}，请稍后……")
    
    directory = extract_file_name(file_path=file_paths[0], type='directory')
    results = []

    with WordAppManager() as word_app:
        for file_path in file_paths:
            try:
                doc = word_app.Documents.Open(file_path)

                count = statistics_of_word_information(doc, statistic_type, include_head_and_foot)

                filename = extract_file_name(file_path=file_path, type='full_name')
                results.append(f"{filename}\t{count}")

            except Exception as e:
                print(f"process_word_statistics: {e}")

            finally:
                if 'doc' in locals():
                    doc.Close(SaveChanges=False)  # 关闭文档，不保存更改
        
            # 更新状态信息
            current_file += 1
            update_info(f"进度：{current_file} / {total_files}，请稍后……")
    
    write_text(results, directory, '000_count_word_pages.txt')
    update_info(f"已检查 {total_files} 个文件，报告保存在文件目录下")

    print("统计 Word 信息完成")