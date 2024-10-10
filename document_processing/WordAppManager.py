import win32com.client

class WordAppManager:
    """管理 Word 对象"""
    def __enter__(self):
        self.word_app = win32com.client.DispatchEx('Word.Application')
        return self.word_app

    def __exit__(self, exc_type, exc_value, traceback):
        if self.word_app:
            self.word_app.Quit()