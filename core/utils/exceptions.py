# core/utils/exceptions.py

class DocumentComparisonError(Exception):
    """Word 文档比较操作失败时的自定义异常"""

    def __init__(self, message="文档比较过程中发生错误"):
        self.message = message
        super().__init__(self.message)


class DocumentConversionError(Exception):
    """转换文档格式失败时的自定义异常"""

    def __init__(self, message="文档转换发生错误"):
        self.message = message
        super().__init__(self.message)


class DocumentProcessingError(Exception):
    """处理文档失败时的自定义异常"""

    def __init__(self, message="处理文档发生错误"):
        self.message = message
        super().__init__(self.message)
