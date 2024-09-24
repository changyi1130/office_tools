# 提取 PDF 页数

import fitz # PyMuPDF
import tkinter.messagebox

from base.get_paths import get_paths
from base.get_file_info import get_file_info
from base.write_to_txt import write_to_txt

def count_pdf_pages(file_path):
    """提取 PDF 页数"""
    try:
        with fitz.open(file_path) as doc:
            num_pages = len(doc)
    except Exception as e:
        return str(e)
    
    return num_pages

def process_pdf_pages(callback):
    # 文件筛选器
    file_filter = [('PDF 文件', '*.pdf'), ('所有文件', '*')]
    file_paths = get_paths('打开文件', file_filter=file_filter)

    # 检查是否选择了文件
    if file_paths is None:
        callable("已取消")
        return None

    results = []

    # 检查是否选择了文件
    if file_paths:
        # 提取工作目录
        work_directory = get_file_info(file_path=file_paths[0], type='directory')
        print(work_directory)

        # 窗口进度标签信息
        total_files = len(file_paths)
        current_file = 0
        callback(f"进度：{current_file} / {total_files}")

        for file_path in file_paths:
            num_pages = count_pdf_pages(file_path=file_path)
            filename = get_file_info(file_path=file_path, type='all_name')
            results.append(f"{filename}\t{num_pages}")

            current_file += 1
            callback(f"进度：{current_file} / {total_files}")

        # 写入 txt 文件
        write_to_txt(texts=results,
                     directory=work_directory,
                     filename='000_count_PDF_pages.txt')

        # tkinter.messagebox.showinfo('提示', '统计完成！')