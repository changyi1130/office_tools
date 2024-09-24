import win32com.client as win32
import logging
import os

from base.get_path import get_path
from base.get_file_info import get_file_info

# 配置日志
# 获取用户文档目录
user_documents = os.path.expanduser('~\\Documents')
log_directory = os.path.join(user_documents, 'highlight_logs')

# 确保日志目录存在
os.makedirs(log_directory, exist_ok=True)

# 配置日志
logging.basicConfig(filename=os.path.join(log_directory, 'highlight_terms.log'), level=logging.INFO, format='%(asctime)s - %(message)s')

def read_terms_from_txt(file_path):
    """从 TXT 文件读取术语"""
    with open(file_path, 'r', encoding='utf-8') as file:
        terms = [line.strip().lower() for line in file if line.strip()]
    return terms

def highlight_terms_in_word(doc, terms, callback):
    """在 Word 文档中查找并高亮术语"""

    highlighted_count = 0

    for paragraph in doc.Paragraphs:
        range_obj = paragraph.Range
        for term in terms:
            term_lower = term.lower()
            start_index = range_obj.Text.lower().find(term_lower)

            while start_index != -1:  # 找到术语，并进行高亮
                end_index = start_index + len(term_lower)

                # 创建要高亮的文本范围
                highlight_range = range_obj.Duplicate  # 复制段落的范围
                highlight_range.Start += start_index  # 移动开始位置
                highlight_range.End = highlight_range.Start + len(term_lower)  # 设置结束位置

                # 高亮显示术语
                highlight_range.HighlightColorIndex = win32.constants.wdYellow  # 使用黄色高亮

                # 更新计数和日志
                highlighted_count += 1
                logging.info(f'Highlighted term "{term}" in paragraph: {paragraph.Range.Text.strip()}')

                # 继续查找同一段落中的下一个实例
                start_index = range_obj.Text.lower().find(term_lower, end_index)

        # 回调函数
        callback(f"已标记 {highlighted_count} 个术语，请不要关闭窗口")

    return highlighted_count

def run_highlight_term(callback):
    # 选择术语文件
    file_filter = [('文本文档', '*.txt')]
    txt_file_path = get_path('请选择术语', file_filter)

    # 选择 Word 文档
    file_filter = [('Word 文档', '*.docx')]
    file_path = get_path('请选择 Word 文档', file_filter)

    # 启动 Word 应用程序
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False  # 设置为 True 以查看 Word 操作过程

    # 打开文档
    doc = word.Documents.Open(file_path)

    terms = read_terms_from_txt(txt_file_path)
    count = highlight_terms_in_word(doc, terms, callback)

    # 保存位置
    directory = get_file_info(file_path, 'directory')
    save_path = directory + '\\highlighted_terms.docx'

    # 保存修改后的 Word 文档
    doc.SaveAs(save_path)
    doc.Close(False)  # 关闭文档，不保存更改

    word.Quit()  # 退出 Word 应用程序

    # 返回高亮术语的总数
    callback(f"已完成，共标记 {count} 个术语。")

    print(f'Total highlighted terms: {count}')