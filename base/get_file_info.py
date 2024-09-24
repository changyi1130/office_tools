def get_file_info(file_path, type='all_name'):
    """
    获取文件信息
    
    :param file_path: 文件全路径
    :param type: 返回类型
    """
    position_point = file_path.rfind('.')
    position_backslash = file_path.rfind('\\')

    if type == 'name':
        # 文件名称
        return file_path[position_backslash+1:position_point]
    elif type == 'directory':
        # 文件路径
        return file_path[:position_backslash]
    elif type == 'all_name':
        # 文件全名
        return file_path[position_backslash+1:]
    elif type == 'ext':
        # 带点扩展名
        return file_path[position_point:]
    elif type == 'ext_not_point':
        # 仅扩展名
        return file_path[position_point+1:]
    elif type == 'path_not_slash':
        # 无分隔符路径
        return file_path.replace('\\', '')
    elif type == 'full_directory':
        # Windows 完整路径
        return file_path