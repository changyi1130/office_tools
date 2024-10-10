def extract_file_name(file_path, type='full_name'):
    """
    传入文件地址，返回文件名、路径、扩展名等内容。
    
    :param file_path: 路径信息
    :param type: 返回类型
    """
    position_point = file_path.rfind('.') # 最后一个间隔点位置
    position_backslash = file_path.rfind('\\') # 最后一个分隔符位置

    if type == 'name':
        # 仅名称（无地址和扩展名）
        return file_path[position_backslash+1:position_point]
    elif type == 'directory':
        # 仅地址
        return file_path[:position_backslash]
    elif type == 'full_name':
        # 全名（名称和扩展名）
        return file_path[position_backslash+1:]
    elif type == 'except':
        # 除扩展名（地址和名称）
        return file_path[:position_point]
    elif type == 'ext':
        # 仅扩展名（有间隔点）
        return file_path[position_point:]
    elif type == 'ext_not_point':
        # 仅扩展名（无间隔点）
        return file_path[position_point+1:]
    elif type == 'path_not_slash':
        # 无分隔符地址
        return file_path.replace('\\', '')
    elif type == 'full_directory':
        # Windows 完整地址（直接返回）
        return file_path
    else:
        raise ValueError(f"提供的类型无效：{type}")