import win32com.client as win32

from base.get_path import get_path
from base.get_file_info import get_file_info

def highlight_revisions(doc):
    """高亮修订内容"""
    doc.TrackRevisions = False

    for revision in doc.Revisions:
        revision.Range.HighlightColorIndex = win32.constants.wdYellow

def run_highlight_revisions(callback):
    """处理文件"""
    # 选择 Word 文档
    file_filter = [('Word 文档', '*.docx')]
    file_path = get_path('请选择 Word 文档', file_filter)

    # 启动 Word 应用程序
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False  # 设置为 True 以查看 Word 操作过程

    # 打开文档
    doc = word.Documents.Open(file_path)

    # 调用处理函数
    highlight_revisions(doc)
    
    callback('正在处理文档，请不要关闭窗口……')

    # 保存位置
    directory = get_file_info(file_path, 'directory')
    filename = get_file_info(file_path, 'name')
    ext = get_file_info(file_path, 'ext')
    save_path = directory + '\\' + filename + '-高亮修订内容' + ext
    print(save_path)

    # 保存修改后的 Word 文档
    doc.SaveAs(save_path)
    doc.Close(False)  # 关闭文档，不保存更改

    word.Quit()  # 退出 Word 应用程序

    callback('已将文档中所有修订内容标记黄色高亮。')