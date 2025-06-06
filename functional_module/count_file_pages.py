# 提取 Word、PDF 页数
import win32com.client
import tkinter.messagebox

from file_processing.open_multiple_files import open_multiple_files
from core.utils.extract_path_components import extract_file_name
from core.utils.write_text_to_file import write_text

from functional_module.count_pdf_pages import count_pdf_pages
from functional_module.statistics_of_word_information import statistics_of_word_information

def count_file_pages(update_info):
    file_filter = [('支持文件', '*.pdf;*.doc*'), ('所有文件', '*')]
    file_paths = open_multiple_files('打开文件', file_filter=file_filter)

    # 检查是否选择了文件
    if file_paths is None:
        update_info("未选择文件")
        return None

    # 如有 Word 文件，则启动 Word 应用程序
    if any(f.lower().endswith('.doc') or f.lower().endswith('.docx') for f in file_paths):
        # 尝试启动 Word 应用程序
        try:
            word_app = win32com.client.gencache.EnsureDispatch('Word.Application')
            print("Word 应用程序已启动")
            
        except Exception as e:
            tkinter.messagebox.showerror('错误', f"无法启动 Word 应用程序: {e}")
            return
    
    results = []

    # 提取工作目录
    work_directory = extract_file_name(file_path=file_paths[0], type='directory')
    print(work_directory)

    # 更新提示信息
    total_files = len(file_paths)
    current_file = 0
    update_info(f"进度：{current_file} / {total_files}，请稍后……")
        
    for file_path in file_paths:
        if file_path.lower().endswith('.pdf'):
            num_pages = count_pdf_pages(file_path=file_path)
            filename = extract_file_name(file_path=file_path, type='full_name')
            results.append(f"{filename}\t{num_pages}")
        elif file_path.lower().endswith('.doc') or file_path.lower().endswith('.docx'):
            try:
                doc = word_app.Documents.Open(file_path)
                count = statistics_of_word_information(doc, 2, False)
                filename = extract_file_name(file_path=file_path, type='full_name')
                result = f"{filename}\t{count}"
                print(result)
                results.append(result)

            except Exception as e:
                update_info(f"调用 statistics_of_word_information 时出错: {e}")
            
            finally:
                if 'doc' in locals():
                    doc.Close(SaveChanges=False)

        # 更新提示信息
        current_file += 1
        update_info(f"进度：{current_file} / {total_files}，请稍后……")
        
    # 尝试关闭 Word 应用程序
    try:
        word_app.Quit()

    except Exception as e:
        print(f"关闭失败: {e}")

    # 写入 txt 文件
    results.sort()
    write_text(texts=results, directory=work_directory, filename='000_count_file_pages.txt')
    
    update_info(f"已统计 {total_files} 个文件，报告保存在文件目录下")

    print("统计页数完成")