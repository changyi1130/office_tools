"""
文件整理工具：按原文目录结构组织译文文件。
通过文件名前缀编号（如 001--）将译文文件复制到与原文相同的相对路径下。
只复制，不改名，不删除，不覆盖已存在的目标文件。
"""

import datetime
import re
import shutil
from pathlib import Path
from collections import defaultdict
from tkinter import filedialog
from typing import Dict, Set, List, Tuple, Callable, Optional


# ---------------------- 核心逻辑 ----------------------

def extract_number(filename: str) -> str | None:
    """
    从文件名开头提取数字编号，如 "001--xxx" -> "001"
    编号必须紧跟 '--'，且不包含其他字符。
    若无匹配返回 None。
    """
    m = re.match(r'^(\d+)--', filename)
    return m.group(1) if m else None


def scan_original_structure(root: Path) -> Tuple[Dict[str, str], Dict[str, Set[str]], list]:
    """
    扫描原文目录结构，返回：
      1) number_to_dir: 编号 -> 相对目录（首次遇到优先）
      2) dir_to_numbers: 相对目录 -> 该目录下所有编号的集合
      3) conflicts: 同编号不同目录的冲突记录
    """
    number_to_dir: Dict[str, str] = {}
    dir_to_numbers: Dict[str, Set[str]] = defaultdict(set)
    conflicts = []

    for file_path in root.rglob('*'):
        if not file_path.is_file():
            continue
        rel_dir = file_path.parent.relative_to(root)
        number = extract_number(file_path.name)
        if number is None:
            continue

        dir_to_numbers[str(rel_dir)].add(number)

        if number in number_to_dir:
            existing_dir = number_to_dir[number]
            if str(rel_dir) != existing_dir:
                conflicts.append((number, existing_dir, str(rel_dir)))
        else:
            number_to_dir[number] = str(rel_dir)

    return number_to_dir, dict(dir_to_numbers), conflicts


def collect_translation_files(trans_dir: Path) -> List[Path]:
    """收集译文目录下所有文件路径（递归）"""
    files = []
    for file_path in trans_dir.rglob('*'):
        if file_path.is_file():
            files.append(file_path)
    return files


def organize(
    trans_files: List[Path],
    number_to_dir: Dict[str, str],
    output_root: Path,
    conflicts_log: list,
    unmatched_log: list,
    progress_callback: Callable[..., None]
) -> int:
    """
    执行文件整理，将译文文件复制到输出目录的正确位置。
    返回成功复制的文件数。
    """
    copied = 0
    output_root.mkdir(parents=True, exist_ok=True)
    total = len(trans_files)

    for idx, src in enumerate(trans_files, start=1):
        number = extract_number(src.name)
        if number is None or number not in number_to_dir:
            unmatched_log.append((str(src), number))
            progress_callback(current=idx, total=total, message=f"未匹配: {src.name}")
            continue

        target_subdir = number_to_dir[number]
        target_dir = output_root / target_subdir
        target_path = target_dir / src.name

        if target_path.exists():
            conflicts_log.append((str(src), str(target_path)))
            progress_callback(current=idx, total=total, message=f"跳过(已存在): {src.name}")
            continue

        target_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, target_path)
        copied += 1
        progress_callback(current=idx, total=total, message=f"已复制: {src.name}")

    return copied


def build_output_number_map(output_root: Path) -> Dict[str, Set[str]]:
    """扫描输出目录，提取每个子目录下的文件编号集合"""
    mapping = defaultdict(set)
    if not output_root.exists():
        return dict(mapping)

    for file_path in output_root.rglob('*'):
        if not file_path.is_file():
            continue
        rel_dir = file_path.parent.relative_to(output_root)
        number = extract_number(file_path.name)
        if number:
            mapping[str(rel_dir)].add(number)
    return dict(mapping)


def compare_completeness(
    original_dir_numbers: Dict[str, Set[str]],
    output_dir_numbers: Dict[str, Set[str]]
) -> str:
    """按目录对比原文编号与输出（译文）编号，生成可读的完整性报告"""
    lines = []
    all_dirs = sorted(set(original_dir_numbers.keys()) | set(output_dir_numbers.keys()))
    if not all_dirs:
        return "（无目录信息）\n"

    lines.append("目录完整性对比（原文 vs 输出译文编号）：")
    lines.append("-" * 60)

    for d in all_dirs:
        orig_nums = original_dir_numbers.get(d, set())
        out_nums = output_dir_numbers.get(d, set())
        missing = orig_nums - out_nums
        extra = out_nums - orig_nums
        total_orig = len(orig_nums)
        matched = len(orig_nums & out_nums)

        status = "✓ 完整" if not missing else f"✗ 缺 {len(missing)} 个"
        lines.append(f"\n📁 目录: {d if d != '.' else '（根目录）'}")
        lines.append(f"   原文编号数: {total_orig}  输出编号数: {len(out_nums)}  匹配: {matched}  [{status}]")
        if missing:
            lines.append(f"   缺失编号: {', '.join(sorted(missing))}")
        if extra:
            lines.append(f"   多余编号（原文中无对应）: {', '.join(sorted(extra))}")

    return "\n".join(lines)


def generate_report_text(
    output_root: Path,
    start_time: datetime.datetime,
    original_root: Path,
    trans_root: Path,
    number_to_dir: Dict[str, str],
    original_dir_numbers: Dict[str, Set[str]],
    num_orig_files: int,
    conflicts_log: list,
    unmatched_log: list,
    mapping_conflicts: list,
    num_trans_files: int,
    copied: int
) -> str:
    """生成整理报告文本"""
    output_dir_numbers = build_output_number_map(output_root)
    completeness_report = compare_completeness(original_dir_numbers, output_dir_numbers)

    num_orig_dirs = len(original_dir_numbers)
    num_output_dirs = len(output_dir_numbers)
    all_original_numbers = set(number_to_dir.keys())
    translated_numbers_in_output = set()
    for nums in output_dir_numbers.values():
        translated_numbers_in_output.update(nums)

    missing_numbers_global = all_original_numbers - translated_numbers_in_output
    extra_numbers_global = translated_numbers_in_output - all_original_numbers

    lines = []
    lines.append("=" * 60)
    lines.append("              文件整理报告")
    lines.append("=" * 60)
    lines.append(f"生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"处理耗时: {datetime.datetime.now() - start_time}")
    lines.append(f"\n原文根目录: {original_root}")
    lines.append(f"译文来源目录: {trans_root}")
    lines.append(f"输出目录: {output_root}")
    lines.append("\n--- 原文结构扫描 ---")
    lines.append(f"原文文件总数: {num_orig_files}")
    lines.append(f"原文唯一编号数: {len(all_original_numbers)}")
    lines.append(f"原文包含目录数: {num_orig_dirs}")

    if mapping_conflicts:
        lines.append("\n⚠️ 编号映射冲突（同编号出现在不同目录，保留首次）:")
        for num, d1, d2 in mapping_conflicts:
            lines.append(f"   编号 {num}: 已映射至 '{d1}'，忽略 '{d2}'")

    lines.append("\n--- 译文文件扫描 ---")
    lines.append(f"译文文件总数: {num_trans_files}")

    lines.append("\n--- 整理结果 ---")
    lines.append(f"成功复制文件: {copied}")
    lines.append(f"未匹配（无原文编号）: {len(unmatched_log)}")
    lines.append(f"目标冲突（跳过）: {len(conflicts_log)}")

    if unmatched_log:
        lines.append("\n未匹配文件列表（无法对应到原文编号）:")
        for fpath, num in unmatched_log[:50]:  # 最多显示前50个
            lines.append(f"   {fpath}  提取编号: {num if num else '无'}")
        if len(unmatched_log) > 50:
            lines.append(f"   ... 还有 {len(unmatched_log) - 50} 个未列出")

    if conflicts_log:
        lines.append("\n目标冲突文件列表（跳过复制，最多显示前20个）:")
        for src, dest in conflicts_log[:20]:
            lines.append(f"   源: {src}")
            lines.append(f"   目标: {dest}")
        if len(conflicts_log) > 20:
            lines.append(f"   ... 还有 {len(conflicts_log) - 20} 个未列出")

    lines.append("\n--- 全局编号差异 ---")
    lines.append(f"原文有但译文完全缺失的编号（{len(missing_numbers_global)} 个）: "
                 f"{', '.join(sorted(missing_numbers_global)) if missing_numbers_global else '无'}")
    lines.append(f"译文有但原文不存在的编号（{len(extra_numbers_global)} 个）: "
                 f"{', '.join(sorted(extra_numbers_global)) if extra_numbers_global else '无'}")

    lines.append("\n" + completeness_report)
    lines.append("\n" + "=" * 60)
    lines.append("报告结束")

    return "\n".join(lines)


def select_directory(title: str = "选择目录") -> Optional[Path]:
    """打开目录选择对话框"""
    dir_path = filedialog.askdirectory(title=title, mustexist=True)
    return Path(dir_path) if dir_path else None


# ---------------------- 主入口函数 ----------------------

def organize_translation_files(progress_callback: Callable[..., None]) -> None:
    """
    按原文目录结构整理译文文件的主入口函数。
    用户依次选择：原文目录 → 译文目录 → 输出目录。
    """
    # 1. 选择原文目录
    progress_callback(message="请选择【原文目录】...")
    original_root = select_directory("选择原文根目录")
    if not original_root:
        progress_callback(message="操作已取消：未选择原文目录")
        return
    progress_callback(message=f"已选择原文目录: {original_root}")

    # 2. 选择译文目录
    progress_callback(message="请选择【译文目录】...")
    trans_root = select_directory("选择待整理的译文目录")
    if not trans_root:
        progress_callback(message="操作已取消：未选择译文目录")
        return
    progress_callback(message=f"已选择译文目录: {trans_root}")

    # 3. 选择输出目录
    progress_callback(message="请选择【输出目录】...")
    output_root = select_directory("选择输出目录（整理后的文件将保存至此）")
    if not output_root:
        progress_callback(message="操作已取消：未选择输出目录")
        return
    progress_callback(message=f"已选择输出目录: {output_root}")

    start_time = datetime.datetime.now()
    progress_callback(message="正在扫描原文结构...")

    # 4. 扫描原文结构
    number_to_dir, orig_dir_numbers, mapping_conflicts = scan_original_structure(original_root)
    num_orig_files = sum(1 for _ in original_root.rglob('*') if _.is_file())
    progress_callback(message=f"原文扫描完成：共 {num_orig_files} 个文件，{len(number_to_dir)} 个编号")

    # 5. 收集译文文件
    progress_callback(message="正在收集译文文件...")
    trans_files = collect_translation_files(trans_root)
    progress_callback(message=f"译文文件收集完成：共 {len(trans_files)} 个文件")

    # 6. 执行整理（复制文件）
    progress_callback(message="开始整理文件（复制中）...")
    conflicts_log = []
    unmatched_log = []

    copied = organize(
        trans_files,
        number_to_dir,
        output_root,
        conflicts_log,
        unmatched_log,
        progress_callback
    )

    progress_callback(message=f"整理完成！成功复制 {copied} 个文件，"
                              f"跳过 {len(conflicts_log)} 个冲突，"
                              f"未匹配 {len(unmatched_log)} 个")

    # 7. 生成报告
    progress_callback(message="正在生成报告...")
    report = generate_report_text(
        output_root=output_root,
        start_time=start_time,
        original_root=original_root,
        trans_root=trans_root,
        number_to_dir=number_to_dir,
        original_dir_numbers=orig_dir_numbers,
        num_orig_files=num_orig_files,
        conflicts_log=conflicts_log,
        unmatched_log=unmatched_log,
        mapping_conflicts=mapping_conflicts,
        num_trans_files=len(trans_files),
        copied=copied
    )

    # 保存报告到输出目录
    report_path = output_root / "整理报告.txt"
    report_path.write_text(report, encoding='utf-8')
    progress_callback(message=f"报告已保存至: {report_path}")
