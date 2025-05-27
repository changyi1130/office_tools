from config.functions import *

BUTTON_GROUPS  = [
    {
        "name": "统计信息",
        "button": [
            {"text": "提取页数", "command": count_file_pages},
            {"text": "添加编号", "command": process_add_index},
            {"text": "删除编号", "command": process_del_index}
        ]
    },
    {
        "name": "文档处理",
        "button": [
            {"text": "比较 Word", "command": select_and_compare_doc},
            {"text": "高亮修订内容", "command": process_highlight_revisions},
            {"text": "高亮术语", "command": process_highlight_term}
        ]
    },
    {
        "name": "转存文件",
        "button": [
            {"text": "存为低版本\n(doc)", "command": convert_doc},
            {"text": "存为高版本\n(docx)", "command": convert_docx},
            {"text": "存为 PDF\n(pdf)", "command": convert_to_pdf}
        ]
    },
    {
        "name": "打开网页",
        "button": [
            {"text": "特殊字符表", "command": open_character},
        ]
    }
]