import logging
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path

def setup_logger(name: str, log_file: str, level=logging.INFO) -> logging.Logger:
    """配置并返回一个 logger 实例"""
    # 创建日志目录
    log_dir = Path(os.path.expanduser('~\\Documents')) / 'OfficeTools' / 'logs'
    log_dir.mkdir(parents=True, exist_ok=True)

    # 创建 logger
    logger = logging.getLogger(name)
    logger.setLevel(level)

    # 创建日志处理器
    file_handler = RotatingFileHandler(
        log_dir / log_file,
        maxBytes=1024 * 1024 * 5,  # 5MB
        backupCount=3,
        encoding='utf-8'
    )

    # 设置日志格式
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)

    # 添加处理器到 logger
    logger.addHandler(file_handler)

    return logger