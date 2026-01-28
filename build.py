#!/usr/bin/env python3
"""
项目打包构建脚本 (build.py)
使用方法: python build.py
"""

import os
import sys
import shutil
import subprocess
import datetime
from pathlib import Path

from config.setting import Setting

setting = Setting()

# ---------- 用户配置区 ----------
# 请根据你的实际情况修改以下变量
PROJECT_NAME = "集装箱" + " V" + setting.version  # 你的项目/工具名称
MAIN_ENTRY = "main.py"        # 主程序入口文件
ICON_PATH = None              # 图标文件路径，例如 "assets/icon.ico" (可选)

# 需要打包的资源文件列表 (源路径 : 打包后的目标路径)
# 格式: (项目中的相对路径, 在打包临时目录中的目标文件夹)
RESOURCE_FILES = [
    ("core/vba_libs/word_vba.dotm", "core/vba_libs"),
    # 你可以继续添加其他资源，如图片、配置文件等
    # ("assets/config.json", "assets"),
    ("core/vba_libs/ExportToWordForWordCount.vb", "core/vba_libs"),
]

# ---------- 脚本主逻辑 ----------
def main():
    print(f"🔨 开始构建 {PROJECT_NAME}...")
    project_root = Path(__file__).parent

    # 1. 检查主入口文件是否存在
    main_file = project_root / MAIN_ENTRY
    if not main_file.exists():
        print(f"❌ 错误: 未找到主入口文件 {MAIN_ENTRY}")
        sys.exit(1)

    # 2. 检查所有资源文件是否存在
    print("📂 检查资源文件...")
    missing_files = []
    for src, _ in RESOURCE_FILES:
        if not (project_root / src).exists():
            missing_files.append(src)
    if missing_files:
        print("❌ 以下资源文件未找到，请检查:")
        for f in missing_files:
            print(f"    - {f}")
        sys.exit(1)
    print("✅ 所有资源文件检查通过。")

    # 3. 构建 PyInstaller 命令
    # 确定操作系统特定的路径分隔符
    separator = ";" if os.name == "nt" else ":"

    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    dist_dir = f"dist_{timestamp}"
    cmd = [
        "pyinstaller",
        "--onefile",  # 打包成单个exe
        "--windowed",  # 不显示控制台窗口 (对于GUI应用)
        "--clean",     # 清理临时文件
        f"--name={PROJECT_NAME}",
        # f"--distpath={dist_dir}",
    ]

    # 添加图标 (如果提供)
    if ICON_PATH and (project_root / ICON_PATH).exists():
        cmd.append(f"--icon={ICON_PATH}")

    # 添加资源文件
    for src, dst in RESOURCE_FILES:
        cmd.append(f"--add-data={src}{separator}{dst}")

    # 添加主程序文件
    cmd.append(str(main_file))

    # 4. 清理旧的构建目录 (可选)
    if (project_root / "dist").exists():
        shutil.rmtree(project_root / "dist")
        print("🧹 已清理旧的 dist 目录。")
    if (project_root / "build").exists():
        shutil.rmtree(project_root / "build")
        print("🧹 已清理旧的 build 目录。")
    spec_file = project_root / f"{PROJECT_NAME}.spec"
    if spec_file.exists():
        spec_file.unlink()
        print("🧹 已清理旧的 .spec 文件。")

    # 5. 显示命令并执行
    print("🚀 执行打包命令:")
    print(f"    {' '.join(cmd)}")
    print("-" * 50)

    try:
        # 使用 subprocess 运行命令，并实时输出信息
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print(result.stdout)
        if result.stderr:
            print("警告信息:", result.stderr)
    except subprocess.CalledProcessError as e:
        print(f"❌ 打包过程出错 (返回码: {e.returncode}):")
        print(e.stderr)
        sys.exit(1)
    except FileNotFoundError:
        print("❌ 未找到 'pyinstaller' 命令。请确保已安装 PyInstaller:")
        print("    运行: pip install pyinstaller")
        sys.exit(1)

    # 6. 打包完成后的信息
    print("-" * 50)
    print(f"🎉 构建成功完成!")
    final_exe = project_root / "dist" / f"{PROJECT_NAME}.exe"
    if final_exe.exists():
        print(f"✅ 可执行文件位于: {final_exe}")
        print(f"    文件大小: {final_exe.stat().st_size / (1024*1024):.2f} MB")
    else:
        print(f"⚠️  未在预期位置找到可执行文件，请检查 build 目录。")

    # 7. 快速验证提示
    print("\n📋 【验证步骤】")
    print("1. 请将生成的 exe 文件复制到一个新的空文件夹。")
    print("2. 在该文件夹中运行它，测试所有功能。")
    print("3. 确认 Word/Excel 宏调用和资源加载正常。")

if __name__ == "__main__":
    main()