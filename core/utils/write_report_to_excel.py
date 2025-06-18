from pathlib import Path

import pandas as pd


def write_report_to_excel(
        report_data: list[list],
        column_headers: list[str],
        output_path: Path
) -> None:
    """将报告写数据写入 Excel 文件

    Args:
    report_data: 二维列表格式的报告数据（无标题行）
    column_headers: 列表标题行
    output_path: Excel 文件输出路径

    Raises:
        ValueError: 当数据与列标题不匹配时
        PermissionError: 当文件被占用或无写入权限时
    """

    # 创建 DataFrame 并写入 Excel
    try:
        df = pd.DataFrame(report_data, columns=column_headers)
        df.to_excel(output_path, index=False)
    except ValueError as e:
        # 列数不匹配
        expected_columns = len(column_headers)
        actual_columns = len(report_data[0]) if report_data else 0

        if actual_columns != expected_columns:
            error_msg = (f"列数不匹配：标题有 {expected_columns} 列，"
                         f"数据有 {actual_columns} 列")
            raise ValueError(error_msg) from e
        raise  # 重新抛出其他 ValueError
