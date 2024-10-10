import jieba
import jieba.posseg as pseg
import win32com.client
import tkinter as tk
import tkinter.messagebox

from collections import Counter

from document_processing.WordAppManager import WordAppManager
from file_processing.open_single_file import open_single_file
from functional_module.highlight_terms_in_word import highlight_terms_in_word
from other_functions.contains_chinese import contains_chinese
from other_functions.extract_file_name import extract_file_name 
from other_functions.write_text import write_text

# 词频和长度
user_count = None
user_length = None


def input_count_and_length(root):
    """用户输入词频和长度"""
    input_window = tk.Toplevel(root)
    input_window.title("请输入词频和长度")

    # 设置窗口大小并禁止调整
    width = 300
    height = 200
    screenwidth = input_window.winfo_screenwidth()
    screenheight = input_window.winfo_screenheight()
    input_window.geometry('%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2))
    input_window.resizable(False, False)

    count_label = tk.Label(input_window, text="词频（大于等于）：") 
    count_label.pack(pady=5, side='top')
    count_entry = tk.Entry(input_window)
    count_entry.pack(pady=5)
    count_entry.insert(0, '3')

    length_label = tk.Label(input_window, text="长度（大于等于）：")
    length_label.pack(pady=5, side='top')
    length_entry = tk.Entry(input_window)
    length_entry.pack(pady=5)
    length_entry.insert(0, '2')

    # 提交按钮
    submit_button = tk.Button(input_window, text="提交",
                              command=lambda: confirm(count_entry.get(),
                                                      length_entry.get(),
                                                      input_window))
    submit_button.pack(pady=10)

    # 关闭窗口事件
    input_window.protocol("WM_DELETE_WINDOW", lambda: on_closing(input_window))

    # 阻塞窗口，等待用户输入
    input_window.wait_window()

def confirm(count_entry, length_entry, input_window):
    """检查输入是否合法"""
    global user_count, user_length
    try:
        user_count = int(count_entry)
        user_length = int(length_entry)
        
        # 如输入值过小，则抛出异常，重新输入
        if user_count <= 1 and user_length <= 1:
            raise ValueError
        
        input_window.destroy()

    except ValueError:
        tkinter.messagebox.showerror("错误", "请输入有效的整数")
        # 清空输入框以便用户重新输入
        count_entry.delete(0, tk.END)
        length_entry.delete(0, tk.END)
        count_entry.insert(0, "10")  # 可以重新设置默认值
        length_entry.insert(0, "5")  # 可以重新设置默认值

def on_closing(input_window):
    """窗口关闭事件"""
    global user_count
    user_count = None
    input_window.destroy()

def text_segmentation(root, update_info):
    """分析高频词"""

    # 获取用户输入，如果用户关闭窗口则退出
    input_count_and_length(root)
    global user_count
    if user_count is None:
        return

    file_path = open_single_file('选择文档')

    # 检查是否选择了文件
    if file_path is None:
        update_info("已取消")
        return None
    
    update_info('正在读取文档内容...')
    
    try:
        word_app = win32com.client.DispatchEx('Word.Application')
        doc = word_app.Documents.Open(file_path)
        text = doc.Content.Text
    
    except Exception as e:
        print(f"错误：{e}")
        return

    # 加载用户自定义词典
    try:
        # 暂时还没有这个文件
        jieba.load_userdict('D:\\临时\\工具箱-对照表高亮术语测试\\custom_dict.txt')
        update_info("已加载自定义词典")

    except FileNotFoundError as e:
        update_info("自定义词典文件未找到")

    update_info("正在统计词频...")

    # 使用 jieba 进行中文分词
    words = jieba.cut(text)

    # directory = extract_file_name(file_path, 'directory')
    # write_text(texts=words, directory=directory, filename='读取内容')

    # 统计词频
    word_counts = Counter(words)

    # 仅保留出现次数大于3的词，长度大于2的词
    # 输出词和频率
    # filtered_words = [word + '\t' + str(count) 
    #                   for word, count in word_counts.items() 
    #                   if count >= user_count and len(word) >= user_length and contains_chinese(word)]
    
    # 仅输出词
    word_list = [word 
                 for word, count in word_counts.items() 
                 if count >= user_count and len(word) >= user_length and contains_chinese(word)]

    # 获取工作路径
    directory = extract_file_name(file_path, 'directory')
    
    # 将结果写入 txt 文件
    txt_name = extract_file_name(file_path, 'except') + '-词频统计.txt'
    write_text(word_list, directory, txt_name)

    result = highlight_terms_in_word(doc, word_list, update_info)

    # 另存处理完的 Word 文档
    save_name = extract_file_name(file_path, 'except') + '-标记高频词' + extract_file_name(file_path, 'ext')
    doc.SaveAs2(save_name)

    doc.Close(SaveChanges=False)

    update_info(f"已完成，共标记了 {result} 个词，保存在文件目录下")

    print("分析高频词完成")