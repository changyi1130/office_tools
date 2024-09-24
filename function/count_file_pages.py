# 提取 Word、PDF 页数
import win32com.client
import tkinter.messagebox

from base.get_paths import get_paths
from base.get_file_info import get_file_info
from base.write_to_txt import write_to_txt

from function.compute_pdf_pages import count_pdf_pages
from function.compute_word_statistics import compute_word_statistics

def count_file_pages(callback):
    file_filter = [('支持文件', '*.pdf;*.doc*'), ('所有文件', '*')]
    file_paths = get_paths('打开文件', file_filter=file_filter)

    # 检查是否选择了文件
    if file_paths is None:
        callable("未 选择文件")
        return None

    # 如有 Word 文件，则启动 Word 应用程序
    if any(f.lower().endswith('.doc') or f.lower().endswith('.docx') for f in file_paths):
        # 尝试启动 Word 应用程序
        try:
            word_app = win32com.client.DispatchEx('Word.Application')
            print("Word 应用程序已启动")
            # 使 Word 应用程序不可见
            # word_app.Visible = False
        except Exception as e:
            tkinter.messagebox.showerror('错误', f'无法启动 Word 应用程序: {e}')
            exit()
    
    results = []

    if file_paths:
        # 提取工作目录
        work_directory = get_file_info(file_path=file_paths[0], type='directory')
        print(work_directory)

        # 窗口进度标签信息
        total_files = len(file_paths)
        current_file = 0
        callback(f"进度：{current_file} / {total_files}")
        
        for file_path in file_paths:
            if file_path.lower().endswith('.pdf'):
                num_pages = count_pdf_pages(file_path=file_path)
                filename = get_file_info(file_path=file_path, type='all_name')
                results.append(f"{filename}\t{num_pages}")
            elif file_path.lower().endswith('.doc') or file_path.lower().endswith('.docx'):
                try:
                    doc = word_app.Documents.Open(file_path)
                    count = compute_word_statistics(doc, 2, False)
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
        
        # 尝试关闭 Word 应用程序
        try:
            word_app.Quit()
        except Exception as e:
            print(f"未打开 Word: {e}")

        # 写入 txt 文件
        results.sort()
        write_to_txt(texts=results,
                     directory=work_directory,
                     filename='000_count_files_pages.txt')

        # tkinter.messagebox.showinfo('提示', '统计完成！')