# 提取 Word 文件信息

import win32com.client
import tkinter.messagebox

from base.get_paths import get_paths
from base.get_file_info import get_file_info
from base.write_to_txt import write_to_txt

global_progress = None

def compute_word_statistics(doc, statistic_type, include_head_and_foot):
    """
    提取 Word 文档的统计信息。
    
    :param doc: Word 文档对象
    :param statistic_type: 统计类型
    :param include_head_and_foot: 是否包含页眉和页脚
    :return: 统计信息的数量
    """
    # 设置查看模式为最终状态，不显示修订标记
    doc.ShowRevisions = False   # 隐藏修订

    try:
        count = doc.ComputeStatistics(Statistic=statistic_type, 
            IncludeFootnotesAndEndnotes=include_head_and_foot)
        return count
    except Exception as e:
        print(f"提取统计信息时发生错误: {e}")
        return None
    
# 名称                               值  描述
# wdStatisticWords                   0   字数
# wdStatisticLines                   1   行
# wdStatisticPages                   2   页数
# wdStatisticCharacters              3   字符数(不计空格)
# wdStatisticParagraphs              4   段落数
# wdStatisticCharactersWithSpaces    5   字符数(计空格)
# wdStatisticFarEastCharacters       6   中文字符和朝鲜语单词

def process_word_statistics(statistic_type, include_head_and_foot, callback):
    """调用 compute_word_statistics 函数"""
    file_filter = [('Word 文档', '*.doc*')]
    file_paths = get_paths('打开文件', file_filter=file_filter)

    # 检查是否选择了文件
    if file_paths is None:
        callable("已取消")
        return None
    
    word_app = win32com.client.DispatchEx('Word.Application')
    work_directory = get_file_info(file_path=file_paths[0], type='directory')
    print(work_directory)

    results = []

    # 窗口进度标签信息
    total_files = len(file_paths)
    current_file = 0
    callback(f"进度：{current_file} / {total_files}")

    for file_path in file_paths:
        try:
            doc = word_app.Documents.Open(file_path)
            
            count = compute_word_statistics(doc, statistic_type, include_head_and_foot)
            print(f"Page: {count}")

            filename = get_file_info(file_path=file_path, type='all_name')
            print(filename)

            results.append(f"{filename}\t{count}")
        except Exception as e:
            print(f"调用 compute_word_statistics 时出错: {e}")
        
        finally:
            if 'doc' in locals():
                doc.Close(SaveChanges=False)  # 关闭文档，不保存更改
    
        current_file += 1
        callback(f"进度：{current_file} / {total_files}")
    
    word_app.Quit()

    write_to_txt(results, work_directory, '000_count_word_pages.txt')
    # tkinter.messagebox.showinfo('提示', '统计完成！')