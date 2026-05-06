import os
from pathlib import Path


class Setting:
    """应用配置管理类，集中管理所有配置项"""

    def __init__(self):
        # 应用版本
        self.version = "0.1.9"

        # 日志配置
        self.log_dir = Path(os.path.expanduser(r'~\Documents')) / 'OfficeTools' / 'logs'
        self.log_dir.mkdir(parents=True, exist_ok=True)

        # VBA模板路径
        self.word_vba_template = "core/vba_libs/word_vba.dotm"
        self.excel_vba_template = "core/vba_libs/excel_vba.xlam"
        self.excel_vba_macro_template = "core/vba_libs/ExportToWordForWordCount.vb"

        # Office应用安全设置
        self.automation_security = 1  # msoAutomationSecurityLow

        # 文件处理配置
        self.default_encoding = 'utf-8'
        self.log_max_bytes = 1024 * 1024 * 5  # 5MB
        self.log_backup_count = 3

        # UI配置
        self.window_width = 520
        self.window_height = 700
        self.progress_window_width = 600
        self.progress_window_height = 400