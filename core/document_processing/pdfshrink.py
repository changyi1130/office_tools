"""压缩 PDF 大小，以便上传至术语工具中"""
import os
import shutil
import sys
from typing import Callable

import pikepdf
import pymupdf
from PIL import Image
from io import BytesIO

from core.utils.extract_path_components import extract_path_components
from core.utils.open_file_dialog import open_file_dialog


import subprocess
from pathlib import Path

import subprocess
import sys
import re
from pathlib import Path


def find_ghostscript_windows():
    """
    专门为 Windows 系统查找 Ghostscript 可执行文件。
    返回完整路径，如果未找到则返回 None。
    """
    # 1. 定义主要的安装目录（64位系统典型路径）
    possible_base_dirs = [
        Path("C:/Program Files/gs"),  # 默认64位安装路径
        Path("C:/Program Files (x86)/gs")  # 32位程序在64位系统上的路径
    ]

    # 2. 用于匹配版本号文件夹的正则表达式，例如：gs10.06.0, gs10.01.1
    # 注意：Windows路径分隔符在正则中是普通的`\`，需要转义，但我们用`/`更安全
    version_pattern = re.compile(r'^gs\d+\.\d+\.\d+$')

    # 3. 按顺序尝试的二进制文件名（推荐使用控制台版本，无图形窗口）
    possible_exe_names = ["gswin64c.exe", "gswin32c.exe"]

    for base_dir in possible_base_dirs:
        if not base_dir.exists():
            continue  # 如果基础目录不存在，跳过

        print(f"正在扫描目录: {base_dir}")

        # 扫描该目录下所有符合版本号模式的子文件夹
        for subdir in base_dir.iterdir():
            if subdir.is_dir() and version_pattern.match(subdir.name):
                # 对于每个找到的版本目录，检查其下的 bin 文件夹
                bin_dir = subdir / "bin"
                if bin_dir.exists():
                    # 在 bin 目录下查找可执行文件
                    for exe_name in possible_exe_names:
                        exe_path = bin_dir / exe_name
                        if exe_path.exists():
                            print(f"✓ 找到 Ghostscript: {exe_path}")
                            return exe_path  # 返回第一个找到的

    # 4. 如果以上都没找到，尝试一个备选方案：查询 Windows 注册表
    # Ghostscript 安装时通常会在注册表中记录信息
    try:
        import winreg
        # 尝试从注册表获取安装路径
        reg_path = r"SOFTWARE\GPL Ghostscript"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
            # 获取最新的版本号
            latest_version = winreg.EnumKey(key, 0)
            with winreg.OpenKey(key, latest_version) as subkey:
                gs_path_str, _ = winreg.QueryValueEx(subkey, "GS_DLL")
                # GS_DLL 值通常是类似 "C:\Program Files\gs\gs10.06.0\bin\gsdll64.dll"
                # 我们需要将其转换为可执行文件路径
                gs_path = Path(gs_path_str)
                # 将 dll 文件名替换为 exe 文件名
                possible_exe_path = gs_path.parent / "gswin64c.exe"
                if possible_exe_path.exists():
                    print(f"✓ 通过注册表找到 Ghostscript: {possible_exe_path}")
                    return possible_exe_path
    except Exception as e:
        print(f"注册表查询失败（可能未安装或权限问题）: {e}")
        pass  # 注册表查询失败是正常的，继续后续逻辑

    # 5. 如果所有方法都失败
    print("✗ 未在常规位置找到 Ghostscript。")
    return None


def show_windows_installation_guide():
    """显示 Windows 专属安装指引"""
    print("\n" + "=" * 70)
    print("请为 Windows 系统安装 Ghostscript：")
    print("1. 访问官方下载页面：https://ghostscript.com/releases/gsdnld.html")
    print("2. 下载最新的 'AGPL Release' 版本（例如：gs1006w64.exe）")
    print("3. 运行安装程序，建议使用默认安装路径")
    print("=" * 70 + "\n")


def compress_pdf_with_ghostscript(input_pdf: str, output_pdf: str, level: str = "ebook"):
    """
    使用 Ghostscript 压缩 PDF (Windows 专用版)
    """
    # 1. 查找 Ghostscript 可执行文件
    gs_path = find_ghostscript_windows()
    if not gs_path:
        show_windows_installation_guide()
        sys.exit(1)

    # 2. 压缩级别映射
    level_map = {
        "screen": "/screen",
        "ebook": "/ebook",
        "printer": "/printer",
        "prepress": "/prepress",
    }
    if level not in level_map:
        raise ValueError(f'不支持的压缩级别: {level}。请选择: {list(level_map.keys())}')

    # 3. 处理文件路径
    input_path = Path(input_pdf).resolve()
    output_path = Path(output_pdf).resolve()

    if not input_path.exists():
        raise FileNotFoundError(f'输入文件不存在: {input_path}')

    # 确保输出目录存在
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # 4. 构建命令
    cmd = [
        str(gs_path),  # Ghostscript 可执行文件完整路径
        "-sDEVICE=pdfwrite",  # 指定输出设备为 PDF 写入器
        "-dCompatibilityLevel=1.4",  # 设置 PDF 兼容版本（1.4 = Acrobat 5.0）
        f"-dPDFSETTINGS={level_map[level]}",  # 预设配置（最关键的压缩级别）
        "-dNOPAUSE",  # 禁用每页处理后的暂停提示
        "-dQUIET",  # 减少控制台输出
        "-dBATCH",  # 处理结束后自动退出
        "-dDetectDuplicateImages=true",  # 检测并合并重复图像

        # 强制降低所有图像的分辨率（DPI 值越低，文件越小）
        "-dColorImageResolution=145",  # 彩色图像分辨率（默认 150）
        "-dGrayImageResolution=145",  # 灰度图像分辨率
        "-dMonoImageResolution=150",  # 黑白图像分辨率（文本可保持较高）

        # 启用下采样（必须与上面分辨率参数一起使用）
        "-dDownsampleColorImages=true",
        "-dDownsampleGrayImages=true",
        "-dDownsampleMonoImages=true",

        # 强制使用有损的 JPEG 压缩（大幅减小体积）
        "-dAutoFilterColorImages=false",  # 禁用自动过滤器
        "-dAutoFilterGrayImages=false",
        "-dColorImageFilter=/DCTEncode",  # 彩色图像使用 JPEG 编码
        "-dGrayImageFilter=/DCTEncode",  # 灰度图像使用 JPEG 编码

        "-dCompressFonts=true",  # 压缩字体数据流
        "-dSubsetFonts=true",  # 字体子集化（只嵌入使用的字符）
        f"-sOutputFile={output_path}",  # 输出文件路径
        str(input_path),  # 输入文件路径
    ]

    print(f"正在使用 Ghostscript (v{gs_path.parent.parent.name}) 压缩 PDF...")
    print(f"预设级别: '{level}'")
    print(f"输入文件: {input_path}")
    print(f"输出文件: {output_path}")

    # 5. 执行命令
    try:
        # 设置编码为当前活动代码页，避免中文路径问题
        result = subprocess.run(cmd, check=True, capture_output=True, text=True, encoding='cp936')
        print("\n✓ PDF 压缩成功！")

        # 比较文件大小
        if output_path.exists():
            orig_size = input_path.stat().st_size / 1024
            comp_size = output_path.stat().st_size / 1024
            ratio = (1 - comp_size / orig_size) * 100

            print("-" * 40)
            print(f"原始大小: {orig_size:,.2f} KB")
            print(f"压缩后:   {comp_size:,.2f} KB")
            print(f"减少体积: {orig_size - comp_size:,.2f} KB")
            print(f"压缩率:   {ratio:+.1f}%")
            print("-" * 40)

    except subprocess.CalledProcessError as e:
        print(f"\n✗ Ghostscript 执行失败 (返回码: {e.returncode})")
        if e.stdout:
            print("标准输出:", e.stdout[-500:])  # 只显示最后500个字符
        if e.stderr:
            print("错误输出:", e.stderr[-500:])
        raise
    except FileNotFoundError:
        print(f"\n✗ 找不到可执行文件: {gs_path}")
        print("请检查 Ghostscript 是否正确安装。")
        raise

def run_compress_pdf(update_info: Callable) -> None:
    """执行 compress_pdf"""
    update_info("请选择要压缩的 PDF 文件")

    # 选择文件
    file_paths = open_file_dialog(
        window_title="选择文件",
        file_filter=[("PDF文件", "*.pdf")],
        multi_select=True
    )

    if not file_paths:
        update_info("已取消选择文件")
        return

    # 压缩文件
    update_info("开始压缩选择的 PDF 文件，请稍等……")

    current_count = 0
    total_count = len(file_paths)
    for file_path in file_paths:
        """处理所有文件"""
        current_count += 1
        update_info(f"正在处理：{current_count} / {total_count}")

        output_path = (
            extract_path_components(file_path=file_path, component='without_ext') +
            "-压缩" +
            extract_path_components(file_path=file_path, component='ext'))

        compress_pdf_with_ghostscript(file_path, output_path)

    update_info(f"压缩已完成")