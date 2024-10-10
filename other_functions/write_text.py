import os

def write_text(texts, directory, filename):
    """
    写入文本至记事本

    :param texts: 待写入的文本，建议是列表或字符串
    :param directory: 记事本的位置
    :param filename: 记事本名称
    """

    txt_file = os.path.join(directory, filename)
    with open(txt_file, 'w', encoding='utf-8') as f:
        if isinstance(texts, list):
            for text in texts:
                f.write(text + '\n')
        else:
            f.write(texts)