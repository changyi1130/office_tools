"""高版本 Word 转低版本，低版本 Word 转高版本"""

import win32com.client
import tkinter.messagebox

from base.get_paths import get_paths
from base.get_file_info import get_file_info

def convert_doc_to_docx(filename, doc):
    """另存为高版本"""
    save_name = get_file_info(filename, 'except') + '.docx'

    doc.SaveAs2(FileName=save_name, FileFormat=16) # 16: wdFormatDocumentDefault

def convert_docx_to_doc(filename, doc):
    """另存为低版本"""
    save_name = get_file_info(filename, 'except') + '.doc'

    doc.SaveAs2(FileName=save_name, FileFormat=0) # 0: wdFormatDocument

def convert_doc_to_pdf(filename, doc):
    """另存为 PDF"""

    # 设置查看模式为最终状态，不显示修订标记
    doc.ShowRevisions = False   # 隐藏修订

    save_name = get_file_info(filename, 'except') + '.pdf'

    doc.SaveAs2(FileName=save_name, FileFormat=17) # 17: wdFormatPDF

def convert_doc(callback):
    """调用 convert_doc_to_docx，存为高版本"""
    file_filter = [('Word 文档', '*.doc')]
    file_paths = get_paths('打开文件', file_filter=file_filter)

    # 检查是否选择了文件
    if file_paths is None:
        callback("已取消")
        return None
    
    word_app = win32com.client.DispatchEx('Word.Application')

    # 状态栏信息
    total_files = len(file_paths)
    current_file = 0
    callback(f"进度：{current_file} / {total_files}")

    for file_path in file_paths:
        try:
            doc = word_app.Documents.Open(file_path)
            
            convert_doc_to_docx(file_path, doc)

        except Exception as e:
            print(f"转存高版本时出错: {e}")
        
        finally:
            if 'doc' in locals():
                doc.Close(SaveChanges=False) # 关闭文档，不保存更改
    
        # 更新进度条
        current_file += 1
        callback(f"进度：{current_file} / {total_files}")
    
    word_app.Quit()

    callback("所有文件已转存为高版本")

def convert_docx(callback):
    """调用 convert_docx_to_doc，存为低版本"""
    file_filter = [('Word 文档', '*.docx')]
    file_paths = get_paths('打开文件', file_filter=file_filter)

    # 检查是否选择了文件
    if file_paths is None:
        callback("已取消")
        return None
    
    word_app = win32com.client.DispatchEx('Word.Application')

    # 状态栏信息
    total_files = len(file_paths)
    current_file = 0
    callback(f"进度：{current_file} / {total_files}")

    for file_path in file_paths:
        try:
            doc = word_app.Documents.Open(file_path)
            
            convert_docx_to_doc(file_path, doc)

        except Exception as e:
            print(f"转存低版本时出错: {e}")
        
        finally:
            if 'doc' in locals():
                doc.Close(SaveChanges=False) # 关闭文档，不保存更改
    
        # 更新进度条
        current_file += 1
        callback(f"进度：{current_file} / {total_files}")
    
    word_app.Quit()

    callback("所有文件已转存为低版本")

def convert_to_pdf(callback):
    """调用 convert_doc_to_pdf，存为 PDF"""
    file_filter = [('Word 文档', '*.doc*')]
    file_paths = get_paths('打开文件', file_filter=file_filter)

    # 检查是否选择了文件
    if file_paths is None:
        callback("已取消")
        return None
    
    word_app = win32com.client.DispatchEx('Word.Application')

    # 状态栏信息
    total_files = len(file_paths)
    current_file = 0
    callback(f"进度：{current_file} / {total_files}")

    for file_path in file_paths:
        try:
            doc = word_app.Documents.Open(file_path)
            
            convert_doc_to_pdf(file_path, doc)

        except Exception as e:
            print(f"转存 PDF 时出错: {e}")
        
        finally:
            if 'doc' in locals():
                doc.Close(SaveChanges=False) # 关闭文档，不保存更改
    
        # 更新进度条
        current_file += 1
        callback(f"进度：{current_file} / {total_files}")
    
    word_app.Quit()

    callback("所有文件已转存为 PDF")