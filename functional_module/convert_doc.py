"""高版本 Word 转低版本，低版本 Word 转高版本"""

from core.utils.WordAppManager import WordAppManager
from file_processing.open_multiple_files import open_multiple_files
from core.utils.extract_path_components import extract_file_name

def convert_doc_to_docx(filename, doc):
    """另存为高版本"""
    save_name = extract_file_name(filename, 'except') + '.docx'

    doc.SaveAs2(FileName=save_name, FileFormat=16) # 16: wdFormatDocumentDefault

def convert_docx_to_doc(filename, doc):
    """另存为低版本"""
    save_name = extract_file_name(filename, 'except') + '.doc'

    doc.SaveAs2(FileName=save_name, FileFormat=0) # 0: wdFormatDocument

def convert_doc_to_pdf(filename, doc):
    """另存为 PDF"""

    # 设置查看模式为最终状态，不显示修订标记
    doc.ShowRevisions = False   # 隐藏修订

    save_name = extract_file_name(filename, 'except') + '.pdf'

    doc.SaveAs2(FileName=save_name, FileFormat=17) # 17: wdFormatPDF

def convert_doc(update_info):
    """调用 convert_doc_to_docx，存为高版本"""
    file_filter = [('Word 文档', '*.doc')]
    file_paths = open_multiple_files('打开文件', file_filter=file_filter)

    # 检查是否选择了文件
    if file_paths is None:
        update_info("已取消")
        return None

    # 更新状态信息
    total_files = len(file_paths)
    current_file = 0
    update_info(f"进度：{current_file} / {total_files}，请稍后……")
    
    with WordAppManager() as word_app:
        for file_path in file_paths:
            try:
                doc = word_app.Documents.Open(file_path)
                
                convert_doc_to_docx(file_path, doc)

            except Exception as e:
                update_info(f"转存高版本时出错: {e}")
            
            finally:
                if 'doc' in locals():
                    doc.Close(SaveChanges=False) # 关闭文档，不保存更改
        
            # 更新状态信息
            current_file += 1
            update_info(f"进度：{current_file} / {total_files}，请稍后……")
    
    update_info("所有文件已转存为高版本")
    print("转存完成")

def convert_docx(update_info):
    """调用 convert_docx_to_doc，存为低版本"""
    file_filter = [('Word 文档', '*.docx')]
    file_paths = open_multiple_files('打开文件', file_filter=file_filter)

    # 检查是否选择了文件
    if file_paths is None:
        update_info("已取消")
        return None
    
    # 更新状态信息
    total_files = len(file_paths)
    current_file = 0
    update_info(f"进度：{current_file} / {total_files}，请稍后……")

    with WordAppManager() as word_app:
        for file_path in file_paths:
            try:
                doc = word_app.Documents.Open(file_path)
                
                convert_docx_to_doc(file_path, doc)

            except Exception as e:
                update_info(f"转存低版本时出错: {e}")
            
            finally:
                if 'doc' in locals():
                    doc.Close(SaveChanges=False) # 关闭文档，不保存更改
        
            # 更新状态信息
            current_file += 1
            update_info(f"进度：{current_file} / {total_files}，请稍后……")
    
    update_info("所有文件已转存为低版本")
    print("转存完成")

def convert_to_pdf(update_info):
    """调用 convert_doc_to_pdf，存为 PDF"""
    file_filter = [('Word 文档', '*.doc*')]
    file_paths = open_multiple_files('打开文件', file_filter=file_filter)

    # 检查是否选择了文件
    if file_paths is None:
        update_info("已取消")
        return None
    
    # 更新状态信息
    total_files = len(file_paths)
    current_file = 0
    update_info(f"进度：{current_file} / {total_files}，请稍后……")

    with WordAppManager() as word_app:
        for file_path in file_paths:
            try:
                doc = word_app.Documents.Open(file_path)
                
                convert_doc_to_pdf(file_path, doc)

            except Exception as e:
                update_info(f"转存 PDF 时出错: {e}")
            
            finally:
                if 'doc' in locals():
                    doc.Close(SaveChanges=False) # 关闭文档，不保存更改
        
            # 更新状态信息
            current_file += 1
            update_info(f"进度：{current_file} / {total_files}，请稍后……")
    
    update_info("所有文件已转存为 PDF")
    print("转存完成")