from pathlib import Path
import pandas as pd
from typing import Union, List, Dict, Optional


def write_report_to_excel(
        report_data: Union[List[List], Dict[str, List]],
        output_path: Path,
        merge_first_column: bool = False,
        column_headers: Optional[list[str]] = None
) -> None:
    """将报告写数据写入 Excel 文件

    支持两种数据格式：
    1. 二维列表：必须提供 column_headers
    2. 字典：键作为列标题，值作为列数据，无需 column_headers

    Args:
        report_data: 报告数据，可以是二维列表或字典格式
        output_path: Excel 文件输出路径
        merge_first_column: 是否合并第一列重复的单元格，默认为True
        column_headers: 列表标题行（仅当 report_data 为列表时必需）

    Raises:
        ValueError: 当数据与列标题不匹配时
        PermissionError: 当文件被占用或无写入权限时
        typeError: 当数据类型不支持时
    """

    # 根据数据类型处理数据
    if isinstance(report_data, dict):
        # 处理字典类型数据
        df = pd.DataFrame(report_data)

        # 如果传入了 column_headers，可以用于重命名列
        if column_headers:
            # 检查列数是否匹配
            if len(column_headers) != len(df.columns):
                raise ValueError(
                    f"列数不匹配：字典有 {len(df.columns)} 列，"
                    f"但传入了 {len(column_headers)} 个列标题"
                )
            df.columns = column_headers

    elif isinstance(report_data, list):
        # 处理列表格式数据
        if not column_headers:
            raise ValueError("列表类型必须提供 column_headers 参数")

        # 检查数据是否为空
        if not report_data:
            df = pd.DataFrame(columns=column_headers)
        else:
            # 检查列数是否匹配
            expected_columns = len(column_headers)
            # 获取第一行的列数作为参考
            actual_columns = len(report_data[0])
            if actual_columns != expected_columns:
                raise ValueError(
                    f"列数不匹配：字典有 {expected_columns} 列，"
                    f"但传入了 {actual_columns} 个列标题"
                )

            # 创建 DataFrame
            df = pd.DataFrame(report_data, columns=column_headers)

    else:
        raise TypeError(
            f"不支持的数据类型：{type(report_data)}。"
            f"请使用 List[List] 或 Dict[str, List]"
        )

    # 写入 Excel
    try:
        # 确保输出目录存在
        # output_path.parent.mkdir(parents=True, exist_ok=True)

        # 判断第一列是否需要合并
        if merge_first_column and len(df) > 0 and len(df.columns) > 0:
            # 将第前两列设置为索引（仅当存在多级索引时，前一列才会合并单元格）
            df_with_index = df.copy()
            df_with_index = df_with_index.set_index([df_with_index.columns[0], df_with_index.columns[1]])

            # 写入 Excel，合并索引单元格
            with pd.ExcelWriter(output_path) as writer:
                df_with_index.to_excel(writer,
                                       index=True,
                                       merge_cells=True,
                                       index_label=None)
        else:
            # 正常写入，不合并单元格
            df.to_excel(output_path, index=False)

    except PermissionError as e:
        raise PermissionError(
            f"无法写入文件 {output_path}。请检查文件是否被占用或有无写入权限。"
        ) from e
    except Exception as e:
        # 重新抛出其他异常
        raise e