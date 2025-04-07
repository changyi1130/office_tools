"""
批量修改文件名
添加前缀编号
删除前缀编号
添加后缀
删除后缀
终稿命名
"""

import glob
import os
import re
import ctypes
from natsort import os_sorted
from tkinter import filedialog

from other_functions.extract_file_name import extract_file_name
from other_functions.write_text import write_text

def is_hidden(file):
    """检查文件是否为隐藏"""
    attrs = ctypes.windll.kernel32.GetFileAttributesW(file)
    return attrs != -1 and attrs & 2  # FILE_ATTRIBUTE_HIDDEN = 0x00000002

def get_dirs(root_dir):
    """提取目录"""
    path = root_dir + "\\*"

    dirs = []

    # 递归目录
    for dir in glob.glob(path):
        if os.path.isdir(dir):
            dirs.append(dir)
            dirs.extend(get_dirs(dir))

    # Windows 顺序
    dirs = os_sorted(dirs)

    return dirs

def get_files(root_dir):
    """提取文件"""
    path = root_dir + "\\*"

    all_files = glob.glob(path)

    files = [dir for dir in all_files if not os.path.isdir(dir) and not is_hidden(dir)]

    # Windows 顺序
    files = os_sorted(files)

    return files

def add_index(old_files):
    """添加编号"""

    new_files = []
    index = 1

    for file in old_files:
        dir = extract_file_name(file, "directory")
        filename = extract_file_name(file, "full_name")
        number_prefix = f"{index:03}--"  # f-string 格式化
        index += 1
        new_files.append(dir + "//" + number_prefix + filename)
    
    return new_files

def del_index(old_files):
    """删除编号"""

    new_files = []
    pattern = '^[0-9]{3}--$'

    for file in old_files:
        dir = extract_file_name(file, "directory")
        filename = extract_file_name(file, "full_name")
        index = filename[:5]
        if re.fullmatch(pattern, index):
            new_files.append(dir + "//" + filename[5:])
    
    return new_files

def rename(old_files, new_files):
    """重命名"""
    files_number = len(old_files)
    if files_number == len(new_files):
        for i in range(files_number):
            os.rename(old_files[i], new_files[i])
    
        # 修改后的文件名与原文件数量一致
        return True
    else:
        # 数量不一致（一般有编号就都有编号，不应该不一致）
        return False
        
def process_add_index(update_info):
    """执行添加编号"""

    # 选择目录
    directory_path = filedialog.askdirectory(title="选择目录")  # 打开目录选择对话框
    path = os.path.normpath(directory_path)

    if directory_path == '':
        update_info("未选择目录")
        return
    
    update_info("正在添加编号，请稍后……")

    files_list = []
    dirs_list = [path]

    # 分析路径
    dirs_list += get_dirs(path)

    # 获取所有文件名
    for dir in dirs_list:
        files_list += get_files(dir)
    
    # 创建新文件名
    new_files_list = add_index(files_list)

    # 重命名（添加编号）
    rename(files_list, new_files_list)

    update_info("已为目录下所有文件添加编号")

    print("文件添加编号执行完成")

def process_del_index(update_info):
    """执行删除编号"""

    # 选择目录
    directory_path = filedialog.askdirectory(title="选择目录")  # 打开目录选择对话框
    path = os.path.normpath(directory_path)

    if directory_path == '':
        update_info("未选择目录")
        return
    
    update_info("正在删除编号，请稍后……")

    files_list = []
    dirs_list = [path]

    # get_files(path, files_list)
    # 分析路径
    dirs_list += get_dirs(path)

    # 获取所有文件名
    for dir in dirs_list:
        files_list += get_files(dir)
    
    # 创建新文件名
    new_files_list = del_index(files_list)

    # 重命名（删除编号）
    if not rename(files_list, new_files_list):
        update_info("部分文件无编号，故未修改")
    else:
        update_info("已为目录下所有文件删除编号")

    print("文件删除编号执行完成")

def add_suffix(old_files):
    """添加后缀"""
    
    new_files = []
    
    list_suffix = ("-译前", "-AI", "-QC", "-译后", "-有标红")
    add_suffix_num = "0123"

    suffix = ""
    for num in range(len(add_suffix_num)):
        suffix += list_suffix[add_suffix_num[num]]
    
    print(suffix)
        
    # for file in old_files:
    #     dir = extract_file_name(file, "direction")
    #     filename = extract_file_name(file, "full_name")

        