"""文档处理"""
# Word 文档比较功能
from core.document_processing.compare_word_documents import compare_documents_with_ui
# 高亮修订内容
from core.document_processing.highlight_revisions import highlight_document_revisions
# 取消 Word 中的隐藏
from core.document_processing.unhide_all_content import execute_unhide_workflow

"""信息统计"""
# 统计文件页数
from core.document_processing.count_file_pages import process_page_count_collection
# 统计 Word 文档信息
from core.document_processing.get_document_statistics import process_word_statistics, WordStatisticType
# 添加、删除编号
from core.tasks.rename_files import batch_add_prefix_numbers, batch_remove_prefix_numbers

"""格式转换"""
# 文档类型转换
from core.document_processing.convert_document import convert_document

"""更多功能"""
# 打开网页
from core.webpages import characters, switch_case

BUTTON_GROUPS = [
    {
        "name": "文档处理",
        "button": [
            {"text": "比较 Word",
             "command": compare_documents_with_ui,
             "tip": "比较文档的两个版本",
             "placeholder": False},
            {"text": "高亮修订",
             "command": highlight_document_revisions,
             "tip": "高亮 Word 中的修订内容",
             "placeholder": False},
            {"text": "取消隐藏",
             "command": execute_unhide_workflow,
             "tip": "取消 Word 中的隐藏",
             "placeholder": False}
        ]
    },
    {
        "name": "信息统计",
        "button": [
            {"text": "统计页数",
             "command": process_page_count_collection,
             "tip": "统计文件的页数",
             "placeholder": False},
            {"text": "统计字数",
             "command": process_word_statistics,
             "command_kwargs": {"statistic_type": WordStatisticType.WORDS},
             "tip": "统计 Word 文档的字数",
             "placeholder": False},
            {"text": "统计字符数",
             "command": process_word_statistics,
             "command_kwargs": {"statistic_type": WordStatisticType.CHARACTERS_NO_SPACES},
             "tip": "统计 Word 文档的字符数(不计空格)",
             "placeholder": False},
            {"text": "占位符",
             "placeholder": True},
            {"text": "添加编号",
             "command": batch_add_prefix_numbers,
             "tip": "为选择目录下所有文件添加编号",
             "placeholder": False},
            {"text": "删除编号",
             "command": batch_remove_prefix_numbers,
             "tip": "删除选择目录下所有文件开头的编号",
             "placeholder": False}
        ]
    },
    {
        "name": "格式转换",
        "button": [
            {"text": "存为低版本",
             "command": convert_document,
             "command_kwargs": {"conversion_type": "docx_to_doc"},
             "tip": "批量将 docx 存为 doc",
             "placeholder": False},
            {"text": "存为高版本",
             "command": convert_document,
             "command_kwargs": {"conversion_type": "doc_to_docx"},
             "tip": "批量将 doc 存为 docx",
             "placeholder": False},
            {"text": "存为 PDF",
             "command": convert_document,
             "command_kwargs": {"conversion_type": "to_pdf"},
             "tip": "批量将 docx 存为 pdf",
             "placeholder": False}
        ]
    },
    {
        "name": "网址",
        "button": [
            {"text": "特殊字符表",
             "command": characters,
             "tip": "方便的复制特殊字符",
             "placeholder": False},
            {"text": "切换大小写",
             "command": switch_case,
             "tip": "切换英文字母大小写，或跳转至网页翻译",
             "placeholder": False}
        ]
    }
]
