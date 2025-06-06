from time import sleep
from core.utils.open_file_dialog import open_file_dialog
import pymupdf


def run_task(callback):
    print("开始测试……")
    # for i in range(10):
    #     callback(f"进行中：{i} / 10", (i / 10))
    #     print(f"i = {i}")
    #     sleep(1)

    filename = open_file_dialog(window_title="测试打开文件", multi_select=True)

    print(filename)

    callback("Done", (10 / 10))

    pymupdf_test()


def pymupdf_test():
    """测试 PyMuPDF"""
    print(pymupdf.__doc__)

    file_filter = [('PDF 文档', '*.pdf')]

    path = open_file_dialog(window_title="测试 PyMuPDF", file_filter=file_filter)
    doc = pymupdf.open(path)

    print(doc.page_count)
