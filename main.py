from gui.main_window import MainWindow
from core.services.logger_service import setup_logger

# 初始化日志记录器
logger = setup_logger(__name__, 'app_main.log')

if __name__ == "__main__":
    logger.info("程序启动")
    print("程序开始")

    try:
        app = MainWindow()
        app.mainloop()
    except Exception as e:
        logger.error(f"程序运行时发生错误: {e}", exc_info=True)
        raise
    finally:
        logger.info("程序正常退出")
        print("程序退出")