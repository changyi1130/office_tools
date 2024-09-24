import os

def write_to_txt(texts, directory, filename):
    """
    写入文本至记事本

    :param texts: 待写入的文本列表
    :param directory: 文件夹路径
    :param filename: 文件名
    """

    txt_file = os.path.join(directory, filename)
    with open(txt_file, 'w', encoding='utf-8') as f:
        if isinstance(texts, list):
            for text in texts:
                f.write(text + '\n')
        else:
            f.write(texts)