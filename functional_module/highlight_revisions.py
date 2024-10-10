"""高亮内容修订"""

import win32com.client as win32

from document_processing.WordAppManager import WordAppManager
from file_processing.open_single_file import open_single_file
from other_functions.extract_file_name import extract_file_name

def highlight_revisions(doc):
    """高亮修订内容"""

    for revision in doc.Revisions:
        print(revision.Type)
        if revision.Type not in (13, 10): # 根据打印显示 13 和 10 应该是格式修订
            revision.Range.HighlightColorIndex = win32.constants.wdYellow

def process_highlight_revisions(update_info):
    """处理文件"""
    file_path = open_single_file('请选择 Word 文档')

    # 检查是否选择了文件
    if file_path is None:
        update_info("未选择文件")
        return None

    with WordAppManager() as word_app:
        # 打开文档
        doc = word_app.Documents.Open(file_path)

        # doc.TrackRevisions = False

        update_info('正在处理文档，请稍后…………')

        # 调用处理函数
        highlight_revisions(doc)
        
        # 另存处理完的文件
        save_name = extract_file_name(file_path, 'except') + \
                    '-高亮修订内容' + \
                    extract_file_name(file_path, 'ext')
        doc.SaveAs2(FileName=save_name)

        # 关闭文档，不保存更改
        doc.Close(SaveChanges=False)

        update_info('已将文档中所有修订内容标记黄色高亮。')

    print("高亮修订内容完成")