import pymupdf

from core.utils.open_file_dialog import open_file_dialog


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


def select_files_and_directories() -> Tuple[List[str], List[str]]:
    """选择文件和目录"""
    root = Tk()
    root.withdraw()  # 隐藏主窗口

    # 选择文件
    file_paths = filedialog.askopenfilenames(
        title="选择要处理的文件",
        filetypes=[("所有文件", "*.*")]
    )
    file_paths = list(file_paths)

    # 选择目录
    dir_path = filedialog.askdirectory(title="选择要处理的目录")
    dir_paths = [dir_path] if dir_path else []

    return file_paths, dir_paths
