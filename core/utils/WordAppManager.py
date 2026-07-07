from core.services.logger_service import setup_logger

logger = setup_logger('WordAppManager', 'word_app_manager.log')


class WordAppManager:
    """管理 Word 对象"""

    def __enter__(self):
        import win32com.client as win32
        self.word_app = win32.DispatchEx('Word.Application')
        # self.word_app = win32.gencache.EnsureDispatch('Word.Application')
        self.word_app.Visible = True
        self.word_app.AutomationSecurity = 1 # 设置宏安全级别（1=msoAutomationSecurityLow）
        logger.debug("Word 应用已启动")
        return self.word_app

    def __exit__(self, exc_type, exc_value, traceback):
        try:
            for i in range(self.word_app.Documents.Count):
                try:
                    self.word_app.Documents(1).Close(SaveChanges=False)
                except Exception as e:
                    logger.error(f"关闭文档失败: {e}")
        except Exception as e:
            logger.error(f"循环出现错误：{e}")

        # 退出 Word 应用
        try:
            self.word_app.Quit()
            logger.debug("Word 应用已退出")
        except Exception as e:
            logger.error(f"退出 Word 失败: {e}")

        del self.word_app
