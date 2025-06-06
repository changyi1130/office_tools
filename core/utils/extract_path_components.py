from pathlib import Path
from typing import Literal, Union


def extract_path_components(
        file_path: str,
        component: Literal['name', 'directory', 'full_name', 'without_ext', 'ext', 'ext_without_dot'] = 'full_name'
) -> Union[str, Path]:
    """
    提取文件路径的指定组件

    参数:
        file_path: 文件路径字符串
        component: 要提取的路径组件类型，可选值:
            'name'         - 仅文件名（不含扩展名）
            'directory'    - 文件所在目录路径
            'full_name'    - 完整文件名（含扩展名）
            'without_ext'  - 文件路径不含扩展名
            'ext'          - 文件扩展名（带点）
            'ext_without_dot' - 文件扩展名（不带点）

    返回:
        请求的路径组件（字符串或Path对象）

    异常:
        ValueError: 当提供无效的组件类型时
        FileNotFoundError: 当路径不存在时（可选）
    """
    # 创建Path对象（自动处理路径分隔符）
    path_obj = Path(file_path)

    # 根据请求的组件返回相应部分
    if component == 'name':
        return path_obj.stem
    elif component == 'directory':
        return path_obj.parent
    elif component == 'full_name':
        return path_obj.name
    elif component == 'without_ext':
        return str(path_obj.with_suffix(''))
    elif component == 'ext':
        return path_obj.suffix
    elif component == 'ext_without_dot':
        return path_obj.suffix[1:] if path_obj.suffix else ''
    else:
        raise ValueError(f"无效的组件类型: '{component}'. 可用选项: 'name', 'directory', 'full_name', "
                         "'without_ext', 'ext', 'ext_without_dot'")
