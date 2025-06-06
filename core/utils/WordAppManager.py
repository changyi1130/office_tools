import win32com.client as win32


class WordAppManager:
    """管理 Word 对象"""

    def __enter__(self):
        # self.word_app = win32com.client.DispatchEx('Word.Application')
        self.word_app = win32.gencache.EnsureDispatch('Word.Application')
        self.word_app.Visible = False
        return self.word_app

    def __exit__(self, exc_type, exc_value, traceback):
        try:
            for i in range(self.word_app.Documents.Count):
                try:
                    self.word_app.Documents(1).Close(SaveChanges=False)
                except Exception as e:
                    print(f"关闭文档失败: {e}")
        except Exception as e:
            print(f"循环出现错误：{e}")

        # 退出 Word 应用
        try:
            self.word_app.Quit()
        except Exception as e:
            print(f"退出 Word 失败: {e}")

        del self.word_app
