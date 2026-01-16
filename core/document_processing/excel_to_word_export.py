import os
from typing import Callable

import win32com.client as win32

from core.utils.extract_path_components import extract_path_components
from core.utils.open_file_dialog import open_file_dialog
from core.utils.path_helper import get_resource_path


def add_and_run_vba_macro(excel_app, excel_file):
    """
    向Excel工作簿添加VBA宏并执行
    """
    # 读取VBA代码
    word_vba_template_path = "core/vba_libs/ExportToWordForWordCount.txt"

    # 获取 word vba 资源路径
    vba_code_path = get_resource_path(word_vba_template_path)

    with open(vba_code_path, 'r', encoding='utf-8') as f:
        vba_code = f.read()

    try:
        # 打开Excel工作簿
        wb = excel_app.Workbooks.Open(excel_file)

        try:
            # 获取VBA项目
            vb_project = wb.VBProject
        except Exception as e:
            # 可能需要启用对VBA项目的访问
            print(f"无法访问VBA项目: {e}")
            print(
                "请在Excel中启用：文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 宏设置 -> 勾选'信任对VBA项目对象模型的访问'")
            return False

        # 创建新模块
        new_module = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule

        # 添加VBA代码到模块
        new_module.CodeModule.AddFromString(vba_code)

        # 运行宏
        excel_app.Run("ExportToWordForWordCount")

        # 等待宏执行完成
        import time
        time.sleep(2)  # 等待Word文档创建

        # 删除临时模块
        vb_project.VBComponents.Remove(new_module)

        # 关闭工作簿，不保存更改（不保存VBA模块）
        wb.Close(SaveChanges=False)

        return True

    except Exception as e:
        print(f"执行VBA宏时出错: {e}")
        return False


def execute_excel_vba_macro(update_info: Callable[[str], None]):
    """
    批量执行Excel VBA宏导出到Word
    """
    # 1. 选择多个Excel文件
    excel_files = open_file_dialog(
        window_title="请选择要处理的Excel文件",
        file_filter=[('Excel文件', '*.xls*'), ('所有文件', '*')],
        multi_select=True
    )

    if not excel_files:
        update_info("未选择任何文件")
        return

    if isinstance(excel_files, str):
        excel_files = [excel_files]

    # 2. 处理每个Excel文件
    for excel_file in excel_files:
        try:
            print_excel_file_name = extract_path_components(file_path=excel_file, component='name')
            print_excel_file_name = print_excel_file_name[0:9] + '……'
            update_info(f"正在处理: {print_excel_file_name}")

            # 创建Excel应用实例
            excel_app = win32.DispatchEx('Excel.Application')
            excel_app.Visible = False  # 后台运行
            excel_app.DisplayAlerts = False  # 不显示警告

            try:
                # 添加并运行VBA宏
                success = add_and_run_vba_macro(excel_app, excel_file)

                if success:
                    update_info(f"✓ 完成: {print_excel_file_name}")
                else:
                    update_info(f"✗ 失败: {print_excel_file_name}")

            finally:
                # 确保Excel进程被关闭
                excel_app.Quit()
                del excel_app

        except Exception as e:
            update_info(f"处理文件 '{excel_file}' 时出错: {e}")

    update_info("所有文件处理完成！")