"""比较 Word"""

from document_processing.WordAppManager import WordAppManager
from file_processing.open_single_file import open_single_file
from other_functions.extract_file_name import extract_file_name

def compare_documents(doc1_path, doc2_path):
    """比较 Word"""

    # 设定对比文件保存地址
    compare_file_name = extract_file_name(doc2_path, 'except') + \
                        '——比较文件' + \
                        extract_file_name(doc2_path, 'ext')
    
    with WordAppManager() as word_app:
        # 打开 Word 文档
        doc1 = word_app.Documents.Open(doc1_path)
        doc2 = word_app.Documents.Open(doc2_path)

        # 比较两个文档
        compare_file = word_app.CompareDocuments(doc1, doc2)

        # 关闭 Word
        doc1.Close(SaveChanges=False)
        doc2.Close(SaveChanges=False)

        # 保存比较文件
        compare_file.SaveAs2(FileName=compare_file_name)
        compare_file.Close(SaveChanges=True)

def select_and_compare_doc(update_info):
    # 打开原文件
    first_file_path = open_single_file('原文件')

    # 检查是否选择了文件
    if first_file_path:
        # 打开修改后文件
        second_file_path = open_single_file('修改后文件')

        # 检查是否选择了文件
        if second_file_path:
            if first_file_path == second_file_path:
                update_info("选择的原文件与修改后文件相同")
            else:
                update_info("正在比较，请稍后……")
                compare_documents(first_file_path, second_file_path)
                update_info("比较文件已保存在文件目录下")

    print("比较完成")