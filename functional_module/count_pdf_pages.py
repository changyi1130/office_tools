"""提取 PDF 页数"""

import fitz # PyMuPDF
import tkinter.messagebox

from file_processing.open_multiple_files import open_multiple_files
from other_functions.extract_file_name import extract_file_name
from other_functions.write_text import write_text

def count_pdf_pages(file_path):
    """提取 PDF 页数"""
    try:
        with fitz.open(file_path) as doc:
            num_pages = len(doc)
    except Exception as e:
        return str(e)
    
    return num_pages

def process_pdf_pages(update_info):
    """已不再单独使用此功能"""
    # 文件筛选器
    file_filter = [('PDF 文件', '*.pdf'), ('所有文件', '*')]
    file_paths = open_multiple_files('打开文件', file_filter=file_filter)

    # 检查是否选择了文件
    if file_paths is None:
        update_info("已取消")
        return None

    results = []

    # 检查是否选择了文件
    if file_paths:
        # 提取工作目录
        work_directory = extract_file_name(file_path=file_paths[0], type='directory')
        print(work_directory)

        # 更新提示信息
        total_files = len(file_paths)
        current_file = 0
        update_info(f"进度：{current_file} / {total_files}，请稍后……")

        for file_path in file_paths:
            num_pages = count_pdf_pages(file_path=file_path)
            filename = extract_file_name(file_path=file_path, type='full_name')
            results.append(f"{filename}\t{num_pages}")

            current_file += 1
            update_info(f"进度：{current_file} / {total_files}，请稍后……")

        # 写入 txt 文件
        write_text(texts=results, directory=work_directory, filename='000_count_PDF_pages.txt')