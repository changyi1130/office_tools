import os
from pathlib import Path

from core.utils.ExcelAppManager import ExcelAppManager
from core.utils.WordAppManager import WordAppManager
from core.utils.path_helper import get_resource_path


def run_vba_macro(file_path, macro_name):
    """
    打开Office文件并执行指定的VBA宏。

    :param file_path: 需要处理的目标文件路径（.docx, .xlsx等）
    :param macro_name: 要执行的宏的完整名称，例如 "MyTemplateModule.MyMacro"
    """
    # 包含宏的模板或加载项文件路径
    word_vba_template_path = "core/vba_libs/word_vba.dotm"
    excel_vba_template_path = "core/vba_libs/excel_vba.xlam"

    file_ext = Path(file_path).suffix.lower()

    try:
        if file_ext in ['.doc', '.docx']:
            # 使用 WordAppManager 上下文管理器
            with WordAppManager() as word_app:
                # 通过管理器获得的应用实例打开文档
                doc = word_app.Documents.Open(str(file_path))

                # 获取 word vba 资源路径
                vba_template_path = get_resource_path(word_vba_template_path)

                # 如果需要，加载 VBA 模板
                if vba_template_path and os.path.exists(vba_template_path):
                    # 注意：AddIns 可能只在特定上下文中有效，视你的 VBA 模板类型而定
                    word_app.Application.AddIns.Add(vba_template_path).Installed = True
                    print("成功加载 Word VBA")
                    # 执行宏
                    word_app.Application.Run(macro_name)
                    # 保存并关闭文档
                    doc.Save()
                    doc.Close()
                    # 注意：WordAppManager的__exit__方法应会负责退出应用
                    print(f"成功在 Word 中执行宏 {macro_name}")

        elif file_ext in ['.xls', '.xlsx']:
            # 使用 ExcelAppManager 上下文管理器
            with ExcelAppManager() as excel_app:
                # 通过管理器获得的应用实例打开文档
                wb = excel_app.Workbooks.Open(str(file_path))

                # 获取 excel vba 资源路径
                vba_template_path = get_resource_path(excel_vba_template_path)

                # 如果需要，加载 VBA 模板
                if vba_template_path and os.path.exists(vba_template_path):
                    print("成功加载 Excel VBA")
                    # 注意：AddIns 可能只在特定上下文中有效，视你的 VBA 模板类型而定
                    excel_app.Application.AddIns.Add(vba_template_path).Installed = True

                    # 执行宏
                    excel_app.Application.Run(macro_name)
                    # 保存并关闭文档
                    wb.Save()
                    wb.Close()
                    # 注意：ExcelAppManager的__exit__方法应会负责退出应用
                    print(f"成功在Word中执行宏 {macro_name}")

        else:
            raise ValueError(f"不支持的文件格式: {file_ext}")

    except Exception as e:
        # 更详细的错误处理
        print(f"处理文件 {file_path} 时出错:")
        print(f"错误类型: {type(e).__name__}")
        print(f"错误详情: {e}")
        # 可以根据需要重新抛出异常或进行其他处理
        raise


def execute_vba_on_document(doc, macro_name):
    """
    对已打开的 Word 文档对象执行 VBA 宏。
    注意：调用者需自行管理文档的打开和关闭。
    本函数仅负责加载模板并执行宏。

    :param doc: 已打开的 Word.Document 对象
    :param macro_name: 要执行的宏的完整名称，例如 "Normal.Module1.MyMacro"
    """
    # 包含宏的模板文件路径（与原始函数保持一致）
    word_vba_template_path = "core/vba_libs/word_vba.dotm"

    try:
        # 获取 word vba 资源路径（与原始函数保持一致）
        vba_template_path = get_resource_path(word_vba_template_path)

        # 加载 VBA 模板（与原始函数完全一致的逻辑）
        if vba_template_path and os.path.exists(vba_template_path):
            # 保持你原有的 AddIns 加载方式
            doc.Application.AddIns.Add(vba_template_path).Installed = True
            print("✅ 成功加载 Word VBA 模板")

            # 执行宏
            doc.Application.Run(macro_name)

            print(f"✅ 成功在文档中执行宏: {macro_name}")

        else:
            raise FileNotFoundError(f"VBA模板文件未找到: {vba_template_path}")

    except Exception as e:
        # 改进错误信息，包含文档名
        doc_name = doc.Name if hasattr(doc, 'Name') else '未知文档'
        print(f"❌ 处理文档 '{doc_name}' 时出错:")
        print(f"   错误类型: {type(e).__name__}")
        print(f"   错误详情: {e}")
        raise