# 提取文件页数（PDF 和 Word）
from functional_module.count_file_pages import count_file_pages
# 检查使用的字体
from functional_module.get_used_fonts import process_check_text

# 比较 Word
from functional_module.compare_documents import select_and_compare_doc
# 高亮修订内容
from functional_module.highlight_revisions import process_highlight_revisions
# 高亮 Word 中的术语
from functional_module.highlight_terms_in_word import process_highlight_term
# 标记高频词
from functional_module.text_segmentation import text_segmentation

# 转高、低版本
from functional_module.convert_doc import convert_doc, convert_docx, convert_to_pdf

# 打开特殊字符表
from functional_module.open_web import open_character

# 提取 PDF 页数
from functional_module.count_pdf_pages import process_pdf_pages
# 提取 Word 页数
from functional_module.statistics_of_word_information import process_word_statistics