import win32com.client as win32
import logging
import os

from core.utils.WordAppManager import WordAppManager
from file_processing.open_single_file import open_single_file
from core.utils.contains_chinese import contains_chinese
from core.utils.extract_path_components import extract_file_name

# 配置日志
# 获取用户文档目录
user_documents = os.path.expanduser('~\\Documents')
log_directory = os.path.join(user_documents, 'highlight_logs')

# 确保日志目录存在
os.makedirs(log_directory, exist_ok=True)

# 配置日志
logging.basicConfig(filename=os.path.join(log_directory, 'highlight_terms.log'), level=logging.INFO, format='%(asctime)s - %(message)s')

def read_terms_from_txt(file_path):
    """从 txt 文件读取术语"""
    with open(file_path, 'r', encoding='utf-8') as file:
        terms = [line.strip().lower() for line in file if line.strip()]
    return terms

def highlight_terms_in_word(doc, terms, update_info):
    """在 Word 文档中查找并高亮术语"""
    highlighted_count = 0

    for paragraph in doc.Paragraphs:
        range = paragraph.Range
        print(range.Text)

        if contains_chinese(range.Text):
            for term in terms:
                term_lower = term.lower()
                start_index = range.Text.lower().find(term_lower)
                print(range.Text.lower())
                print(term_lower)

                while start_index != -1: # 找到术语，并进行高亮
                    end_index = start_index + len(term_lower)

                    # 创建要高亮的文本范围
                    highlight_range = range.Duplicate # 复制段落的范围
                    highlight_range.Start += start_index # 移动开始位置
                    highlight_range.End = highlight_range.Start + len(term_lower) # 设置结束位置

                    # 高亮显示术语
                    highlight_range.HighlightColorIndex = win32.constants.wdYellow # 使用黄色高亮
                    # highlight_range.HighlightColorIndex = 7 # 使用黄色高亮

                    # 更新计数和日志
                    highlighted_count += 1
                    logging.info(f'Highlighted term "{term}" in paragraph: {paragraph.Range.Text.strip()}')

                    # 继续查找同一段落中的下一个实例
                    start_index = range.Text.lower().find(term_lower, end_index)

        # 更新提示信息
        update_info(f"正在标记术语 ({highlighted_count})，请稍后……")

    return highlighted_count

def process_highlight_term(update_info):
    """选择文件并执行 highlight_terms_in_word"""

    # 选择术语文件
    file_filter = [('文本文档', '*.txt')]
    txt_file_path = open_single_file('请选择术语文件', file_filter)

    # 检查是否选择了文件
    if txt_file_path is None:
        update_info("未选择文件")
        return None

    # 选择 Word 文档
    file_filter = [('Word 文档', '*.docx')]
    file_path = open_single_file('请选择 Word 文档', file_filter)

    # 检查是否选择了文件
    if file_path is None:
        update_info("未选择文件")
        return None

    with WordAppManager() as word_app:
        doc = word_app.Documents.Open(file_path)

        terms = read_terms_from_txt(txt_file_path)
        count = highlight_terms_in_word(doc, terms, update_info)

        # 另存处理完的文件
        save_as_name = extract_file_name(file_path, 'except') + \
                    '-高亮术语' + \
                    extract_file_name(file_path, 'ext')
        doc.SaveAs2(FileName=save_as_name)

        # 关闭文档，不保存更改
        doc.Close(SaveChanges=False)

        # 返回高亮术语的总数
        update_info(f"已完成，共标记 {count} 个术语。")
    
    print("高亮术语完成")