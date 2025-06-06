# Word 文档比较功能
from core.document_processing.compare_word_documents import compare_documents_with_ui
# 高亮修订内容
from core.document_processing.highlight_revisions import highlight_document_revisions

# 统计文件页数
from core.document_processing.count_file_pages import process_page_count_collection
# 统计 Word 文档信息
from core.document_processing.get_document_statistics import process_word_statistics, WordStatisticType

# 文档类型转换
from core.document_processing.convert_document import convert_document

# 打开特殊字符表
from core.webpages import characters, switch_case

BUTTON_GROUPS = [
    {
        "name": "文档处理",
        "button": [
            {"text": "比较 Word",
             "command": compare_documents_with_ui,
             "tip": "比较文档的两个版本"},
            {"text": "高亮修订",
             "command": highlight_document_revisions,
             "tip": "高亮 Word 中的修订内容"},

        ]
    },
    {
        "name": "格式转换",
        "button": [
            {"text": "存为低版本",
             "command": convert_document,
             "command_kwargs": {"conversion_type": "docx_to_doc"},
             "tip": "批量将 docx 存为 doc"},
            {"text": "存为高版本",
             "command": convert_document,
             "command_kwargs": {"conversion_type": "doc_to_docx"},
             "tip": "批量将 doc 存为 docx"},
            {"text": "存为 PDF",
             "command": convert_document,
             "command_kwargs": {"conversion_type": "to_pdf"},
             "tip": "批量将 docx 存为 pdf"}
        ]
    },
    {
        "name": "文档统计",
        "button": [
            {"text": "统计页数",
             "command": process_page_count_collection,
             "tip": "统计文件的页数"},
            {"text": "统计字数",
             "command": process_word_statistics,
             "command_kwargs": {"statistic_type": WordStatisticType.WORDS},
             "tip": "统计 Word 文档的字数"},
            {"text": "统计字符数",
             "command": process_word_statistics,
             "command_kwargs": {"statistic_type": WordStatisticType.CHARACTERS_NO_SPACES},
             "tip": "统计 Word 文档的字符数(不计空格)"}
        ]
    },
    {
        "name": "网页",
        "button": [
            {"text": "特殊字符表",
             "command": characters,
             "tip": "方便的复制特殊字符"},
            {"text": "切换大小写",
             "command": switch_case,
             "tip": "切换英文字母大小写，或跳转至网页翻译"}
        ]
    },
    {
        "name": "测试",
        "button": [
            {"text": "测试",
             "command": "pymupdf_test",
             "tip": "<UNK>"},
        ]
    }
]
