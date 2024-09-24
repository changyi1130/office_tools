# 比较两个 Word 文档

from base.get_path import get_path
import win32com.client
import tkinter.messagebox

def compare_document(doc1_path, doc2_path):
    # 创建 Word 对象
    word_app = win32com.client.Dispatch('Word.Application')

    # 打开 Word 文档
    doc1 = word_app.Documents.Open(doc1_path)
    doc2 = word_app.Documents.Open(doc2_path)

    # 设定对比文件保存路径
    full_name = doc1.FullName
    point_position = full_name.rfind('.')
    compare_file_name = full_name[:point_position] \
                      + '——比较文件' \
                      + full_name[point_position:]

    # 比较两个文档
    compare_file = word_app.CompareDocuments(doc1, doc2)

    # 关闭 Word
    # 0:  wdDoNotSaveChanges
    # -1: wdSaveChanges
    doc1.Close(0)
    doc2.Close(0)

    # 保存比较文件
    compare_file.SaveAs2(compare_file_name)
    compare_file.Close(0)

    # 显示文档
    # word_app.Visible = True

def run_compare(callback):
    # 文件筛选器
    file_filter = [('Word文档', '*.doc*'), ('所有文件', '*')]

    # 打开原文件
    first_file_path = get_path('原文件', file_filter)

    # 检查是否选择了文件
    if first_file_path:
        # 打开修改后文件
        second_file_path = get_path('修改后文件', file_filter)

        # 检查是否选择了文件
        if second_file_path:
            if first_file_path == second_file_path:
                callback("选择的原文件与修改后文件相同")
                # tkinter.messagebox.showinfo(
                #     '提示', '选择的原文件与修改后文件相同')
            else:
                compare_document(first_file_path, second_file_path)
                callback("比较文件已保存在文件目录下")
                # tkinter.messagebox.showinfo(
                #     '提示', '比较文件已保存')