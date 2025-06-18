import dataclasses
from pathlib import Path
from typing import Optional


@dataclasses.dataclass
class CountResult:
    """页数统计结果封装类"""

    file_path: Path
    page_count: Optional[int] = None
    error: Optional[str] = None

    def to_row(self) -> list:
        """将结果转换为 Excel 行数据"""
        return [
            self.file_path.name,
            f"错误：{self.error}" if self.error else self.page_count
        ]

    def __str__(self) -> str:
        """用于文本报告的字符串"""
        return f"{self.file_path.name}\t{self.error or self.page_count}"
