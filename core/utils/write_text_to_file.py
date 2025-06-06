import os
from pathlib import Path
from typing import Union, List, Optional, Tuple
import logging

logger = logging.getLogger(__name__)


def write_text_to_file(
        contents: Union[str, List[str]],
        target_path: Union[str, Path],
        mode: str = 'w',
        encoding: str = 'utf-8',
        newline: Optional[str] = None,
        create_parents: bool = True,
        overwrite: bool = True
) -> Tuple[bool, str]:
    """
    高级文本文件写入工具

    :param contents: 写入内容（字符串或列表）
    :param target_path: 目标文件路径（支持str/Path）
    :param mode: 写入模式（'w'覆盖/'a'追加）
    :param encoding: 文件编码格式
    :param newline: 换行符（None使用系统默认）
    :param create_parents: 是否自动创建父目录
    :param overwrite: 当文件存在时是否覆盖（仅模式'w'有效）
    :return: (成功状态, 操作信息)
    """
    try:
        # 类型标准化处理
        target_path = Path(target_path)
        contents = _normalize_contents(contents)

        # 路径处理
        if create_parents:
            target_path.parent.mkdir(parents=True, exist_ok=True)

        # 存在性检查
        if not overwrite and target_path.exists() and mode == 'w':
            return False, f"文件已存在: {target_path}"

        # 写入操作
        with target_path.open(mode, encoding=encoding, newline=newline) as f:
            if isinstance(contents, list):
                f.writelines([f"{line}\n" for line in contents])
            else:
                f.write(contents)

        return True, f"成功写入: {target_path}"

    except Exception as e:
        logger.error(f"写入失败: {str(e)}", exc_info=True)
        return False, f"写入失败: {str(e)}"


def _normalize_contents(contents: Union[str, List[str]]) -> Union[str, List[str]]:
    """内容标准化处理"""
    if isinstance(contents, list):
        return [str(item).rstrip('\n') for item in contents]
    return str(contents)
