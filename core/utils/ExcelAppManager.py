import win32com.client as win32
from core.services.logger_service import setup_logger

logger = setup_logger('ExcelAppManager', 'excel_app_manager.log')


class ExcelAppManager:
    """上下文管理器，用于安全地创建和清理 Excel 应用实例"""

    def __enter__(self):
        self.excel_app = win32.DispatchEx('Excel.Application')
        self.excel_app.Visible = True
        self.excel_app.AutomationSecurity = 1  # 设置宏安全级别（1=msoAutomationSecurityLow）
        logger.debug("Excel 应用已启动")
        return self.excel_app

    def __exit__(self, exc_type, exc_value, traceback):
        # 1. 关闭所有打开的工作簿
        try:
            # 从后向前关闭，避免索引变化
            for i in range(self.excel_app.Workbooks.Count, 0, -1):
                wb = self.excel_app.Workbooks(i)
                try:
                    wb.Close(SaveChanges=False)
                except Exception as e:
                    logger.error(f"关闭工作簿 '{wb.Name}' 失败: {e}")
        except AttributeError as e:
            # 可能在 Excel 已意外关闭时发生
            logger.error(f"访问 Workbooks 集合时出错（Excel 可能已关闭）：{e}")
        except Exception as e:
            logger.error(f"关闭工作簿时发生未知错误：{e}")

        # 2. 退出 Excel 应用
        try:
            # 先检查是否还能访问 Quit 方法，避免重复退出错误
            if hasattr(self.excel_app, 'Quit'):
                self.excel_app.Quit()
                logger.debug("Excel 应用已退出")
        except Exception as e:
            logger.error(f"退出 Excel 应用失败: {e}")
        finally:
            # 3. 确保删除 COM 对象引用，释放资源
            try:
                del self.excel_app
            except Exception as e:
                logger.debug(f"删除 Excel 应用引用时出错: {e}")