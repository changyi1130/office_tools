import os
from typing import Callable

from core.utils.ExcelAppManager import ExcelAppManager
from core.utils.open_file_dialog import open_file_dialog
from core.utils.path_helper import get_resource_path
from core.services.logger_service import setup_logger

logger = setup_logger(__name__, 'excel_to_word_export.log')


def add_and_run_vba_macro(excel_app, excel_file, vba_code) -> bool:
    """
    向 Excel 工作簿添加 VBA 宏并执行

    参数:
        excel_app: Excel 应用实例
        excel_file: Excel 文件路径
        vba_code: VBA 代码字符串

    返回:
        bool: 操作是否成功
    """
    logger.info(f"开始处理 Excel 文件：{excel_file}")
    logger.debug(f"VBA 代码长度: {len(vba_code)} 字符")

    try:
        # 打开 Excel 工作簿
        wb = excel_app.Workbooks.Open(excel_file)
        logger.info("Excel 工作簿打开成功")

        try:
            # 获取 VBA 项目
            vb_project = wb.VBProject
        except Exception as e:
            # 可能需要启用对VBA项目的访问
            logger.error(f"无法访问 VBA 项目: {e}")
            logger.error("请在 Excel 中勾选「信任对 VBA 项目对象模型的访问」")
            return False

        # 创建新模块
        new_module = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        logger.debug("VBA 新模块已创建")

        # 添加 VBA 代码到模块
        new_module.CodeModule.AddFromString(vba_code)
        logger.debug("VBA 代码已添加到模块")

        # 运行宏
        excel_app.Run("ExportToWordForWordCount")
        logger.info("VBA 宏已执行")

        # 等待宏执行完成
        import time
        time.sleep(2)  # 等待 Word 文档创建

        # 删除临时模块
        vb_project.VBComponents.Remove(new_module)
        logger.debug("VBA 模块删除成功")

        # 关闭工作簿，不保存更改（不保存VBA模块）
        wb.Close(SaveChanges=False)
        logger.info("Excel 工作簿已关闭")

        return True

    except Exception as e:
        logger.error(f"执行 VBA 宏时出错: {e}", exc_info=True)
        return False


def execute_excel_vba_macro_on_files_simple(progress_callback: Callable[..., None] = None) -> None:
    """
    执行 Excel VBA 宏
    
    参数:
        progress_callback: 进度更新回调函数
    """
    logger.info("开始批量处理 Excel 文件")

    # 1. 选择多个 Excel 文件
    excel_files = open_file_dialog(
        window_title="请选择要导出文本的 Excel 文件",
        file_filter=[('Excel文件', '*.xls*'), ('所有文件', '*')],
        multi_select=True
    )

    if not excel_files:
        logger.info("未选择任何文件")
        if progress_callback:
            progress_callback(message="未选择任何文件")
        return

    # 检查输入的 excel_files 是否为字符串类型
    if isinstance(excel_files, str):
        # 如果是字符串，则将其转换为包含该字符串的列表
        excel_files = [excel_files]

    # 2. 读取 VBA 代码（一次性读取，提高效率）
    excel_export_to_word = "core/vba_libs/ExportToWordForWordCount.vb"
    vba_code_path = get_resource_path(excel_export_to_word)

    try:
        with open(vba_code_path, 'r', encoding='utf-8') as f:
            vba_code = f.read()
        logger.debug("VBA 代码模板读取成功")
    except Exception as e:
        logger.error(f"读取 VBA 模板失败: {e}", exc_info=True)
        if progress_callback:
            progress_callback(message=f"读取 VBA 模板失败: {e}")
        return

    # 3. 处理每个 Excel 文件
    success_count = 0
    fail_count = 0

    # 使用 ExcelAppManager 上下文管理器创建 Excel 应用实例
    with ExcelAppManager() as excel_app:
        excel_app.Visible = False  # 后台运行
        excel_app.DisplayAlerts = False  # 不显示警告
        logger.debug("Excel 应用实例创建成功")

        total = len(excel_files)
        for idx, excel_file in enumerate(excel_files):
            try:
                logger.info(f"开始处理文件: {excel_file}")
                if progress_callback:
                    progress_callback(message=f"正在处理: {os.path.basename(excel_file)}")

                try:
                    # 调用 add_and_run_vba_macro 函数处理文件
                    if add_and_run_vba_macro(excel_app, excel_file, vba_code):
                        success_count += 1
                        if progress_callback:
                            progress_callback(idx + 1, total, message=f"完成：{os.path.basename(excel_file)}")
                    else:
                        fail_count += 1
                        if progress_callback:
                            progress_callback(message=f"失败：{os.path.basename(excel_file)}")

                except Exception as e:
                    logger.error(f"处理文件 {excel_file} 时出错: {e}", exc_info=True)
                    fail_count += 1
                    if progress_callback:
                        progress_callback(message=f"失败：{os.path.basename(excel_file)} - {str(e)}")

            except Exception as e:
                logger.error(f"处理文件 '{excel_file}' 时出错: {e}", exc_info=True)
                fail_count += 1
                if progress_callback:
                    progress_callback(message=f"失败：{os.path.basename(excel_file)} - {str(e)}")

    logger.info(f"批量处理完成: 成功 {success_count}, 失败 {fail_count}")
    if progress_callback:
        progress_callback(total, total, f"所有文件处理完成！成功: {success_count}, 失败 {fail_count}")
