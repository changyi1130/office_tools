# OfficeTools API 文档

## 项目概述

OfficeTools 是一个基于 Python 的办公文档处理工具集，提供文档处理、信息统计、格式转换等多种功能，通过图形界面和命令行方式使用。

## 核心模块

### 1. 文档处理模块 (core/document_processing)

#### 1.1 文档比较 (compare_word_documents.py)

**compare_word_documents(original_path, modified_path)**
- 描述：比较两个 Word 文档并生成比较结果
- 参数：
  - original_path (Path): 原始版本路径
  - modified_path (Path): 更新版本路径
- 返回：Path - 生成的比较文档路径
- 异常：DocumentComparisonError - 文档比较失败时抛出

**compare_documents_with_ui(progress_callback)**
- 描述：带UI的文档比较流程
- 参数：
  - progress_callback (Callable): 进度更新回调函数

#### 1.2 文档转换 (convert_document.py)

**convert_document(conversion_type, progress_callback)**
- 描述：通用文档转换函数
- 参数：
  - conversion_type (str): 转换类型
    - doc_to_docx: DOC转DOCX
    - docx_to_doc: DOCX转DOC
    - to_pdf: 转PDF
    - word_to_text: 转纯文本
  - progress_callback (Callable): 进度更新回调函数

#### 1.3 文档统计 (get_document_statistics.py)

**get_document_statistics(document, statistic_type, include_notes)**
- 描述：获取 Word 文档的统计信息
- 参数：
  - document (win32.CDispatch): Word 文档对象
  - statistic_type (WordStatisticType): 统计类型
    - WORDS: 字数
    - LINES: 行数
    - PAGES: 页数
    - CHARACTERS_NO_SPACES: 字符数(不计空格)
    - PARAGRAPHS: 段落数
    - CHARACTERS_WITH_SPACES: 字符数(计空格)
    - FAR_EAST_CHARACTERS: 中文字符和朝鲜语单词
  - include_notes (bool): 是否包含页眉、页脚和尾注
- 返回：int - 统计结果
- 异常：DocumentProcessingError - 统计失败时抛出

**process_word_statistics(statistic_type, include_notes, update_info)**
- 描述：主处理函数：收集 Word 文档统计信息
- 参数：
  - statistic_type (WordStatisticType): 统计类型
  - include_notes (bool): 是否包含页眉页脚
  - update_info (Callable): 状态更新回调

#### 1.4 文件页数统计 (count_file_pages.py)

**process_page_count_collection(progress_callback)**
- 描述：带UI的页数统计流程
- 参数：
  - progress_callback (Callable): 进度更新回调函数

支持的文件类型：
- PDF: 使用 pymupdf 统计页数
- Word: 使用 Word 统计页数
- PPT: 统计可见幻灯片数
- Excel: 不统计页数，返回0
- Image: 返回1页

#### 1.5 Excel到Word导出 (excel_to_word_export.py)

**execute_excel_vba_macro_on_files_simple(progress_callback)**
- 描述：执行 Excel VBA 宏，将 Excel 内容导出到 Word
- 参数：
  - progress_callback (Callable): 进度更新回调函数

#### 1.6 隐藏/取消隐藏内容

**execute_hide_non_chinese_workflow(progress_callback)**
- 描述：隐藏Word文档中不包含中文的段落
- 参数：
  - progress_callback (Callable): 进度更新回调函数

**execute_unhide_workflow(progress_callback)**
- 描述：取消 Word 中所有隐藏内容
- 参数：
  - progress_callback (Callable): 进度更新回调函数

**highlight_document_revisions(progress_callback)**
- 描述：高亮显示Word文档中的修订内容
- 参数：
  - progress_callback (Callable): 状态更新回调函数

### 2. 任务模块 (core/tasks)

#### 2.1 文件重命名 (rename_files.py)

**batch_add_prefix_numbers(progress_callback)**
- 描述：为目录下所有文件添加编号前缀
- 参数：
  - progress_callback (Callable): 进度更新回调函数

**batch_remove_prefix_numbers(progress_callback)**
- 描述：删除目录下所有文件的编号前缀
- 参数：
  - progress_callback (Callable): 进度更新回调函数

#### 2.2 目录树读取 (read_dirtree.py)

**read_dirtree(update_info)**
- 描述：扫描选定目录，将文件列表生成结构化 Excel 报告
- 参数：
  - update_info (Callable): 状态更新回调

### 3. 工具模块 (core/utils)

#### 3.1 Office应用管理器

**WordAppManager**
- 上下文管理器，用于安全地创建和清理 Word 应用实例
- 自动处理文档关闭和应用退出

**ExcelAppManager**
- 上下文管理器，用于安全地创建和清理 Excel 应用实例
- 自动处理工作簿关闭和应用退出

#### 3.2 文件对话框 (open_file_dialog.py)

**open_file_dialog(window_title, file_filter, multi_select)**
- 描述：通用文件选择对话框
- 参数：
  - window_title (str): 窗口标题
  - file_filter (list): 文件类型过滤器
  - multi_select (bool): 是否允许多选
- 返回：
  - 单选: str - 文件路径
  - 多选: list[str] - 文件路径列表
  - 取消: None

#### 3.3 VBA宏执行 (run_vba_macro.py)

**run_vba_macro(file_path, macro_name)**
- 描述：打开 Office 文件并执行指定的 VBA 宏
- 参数：
  - file_path (str): 需要处理的目标文件路径
  - macro_name (str): 要执行的宏的完整名称

**execute_vba_on_document(doc, macro_name)**
- 描述：对已打开的 Word 文档对象执行 VBA 宏
- 参数：
  - doc (win32.CDispatch): 已打开的 Word.Document 对象
  - macro_name (str): 要执行的宏的完整名称

#### 3.4 Excel报告生成 (write_report_to_excel.py)

**write_report_to_excel(report_data, output_path, merge_first_column, column_headers)**
- 描述：将报告数据写入 Excel 文件
- 参数：
  - report_data (Union[List[List], Dict[str, List]]): 报告数据
  - output_path (Path): Excel 文件输出路径
  - merge_first_column (bool): 是否合并第一列重复的单元格
  - column_headers (list[str]): 列表标题行
- 异常：
  - ValueError: 当数据与列标题不匹配时
  - PermissionError: 当文件被占用或无写入权限时
  - TypeError: 当数据类型不支持时

### 4. 服务模块 (core/services)

#### 4.1 日志服务 (logger_service.py)

**setup_logger(name, log_file, level)**
- 描述：配置并返回一个 logger 实例
- 参数：
  - name (str): logger 名称
  - log_file (str): 日志文件名
  - level (int): 日志级别，默认 logging.INFO
- 返回：logging.Logger - 配置好的 logger 实例

### 5. 网页工具 (core/webpages.py)

**characters()**
- 描述：打开字符转换网页

**switch_case()**
- 描述：打开大小写转换网页

### 6. 异常类 (core/utils/exceptions.py)

**DocumentComparisonError**
- 描述：Word 文档比较操作失败时的自定义异常

**DocumentConversionError**
- 描述：转换文档格式失败时的自定义异常

**DocumentProcessingError**
- 描述：处理文档失败时的自定义异常

## 使用示例

```python
# 文档比较
from pathlib import Path
from core.document_processing.compare_word_documents import compare_word_documents

result = compare_word_documents(
    original_path=Path("original.docx"),
    modified_path=Path("modified.docx")
)

# 文档转换
from core.document_processing.convert_document import convert_document

convert_document(
    conversion_type="docx_to_doc",
    progress_callback=lambda **kwargs: print(kwargs)
)

# 统计Word文档字数
from core.document_processing.get_document_statistics import process_word_statistics, WordStatisticType

process_word_statistics(
    statistic_type=WordStatisticType.WORDS,
    include_notes=True,
    update_info=lambda msg: print(msg)
)
```

## 注意事项

1. 所有 Office 文档操作都需要安装 Microsoft Office
2. 使用 VBA 宏功能需要在 Office 中启用宏安全设置
3. 日志文件默认保存在用户文档目录下的 OfficeTools/logs 文件夹
4. 批量处理大量文件时，建议使用进度回调函数监控处理进度
