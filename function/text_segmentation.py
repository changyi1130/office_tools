import jieba
import jieba.posseg as pseg
import win32com.client as win32
import tkinter as tk
import tkinter.messagebox

from collections import Counter

from base.get_path import get_path
from base.write_to_txt import write_to_txt
from base.get_file_info import get_file_info 
from base.contains_chinese import contains_chinese
from function.highlight_terms_in_word import highlight_terms_in_word

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
    global user_count, user_length
    try:
        user_count = int(count_entry)
        user_length = int(length_entry)
        input_window.destroy()
    except ValueError:
        tkinter.messagebox.showerror("输入错误", "请输入有效的整数")
        # 清空输入框以便用户重新输入
        count_entry.delete(0, tk.END)
        length_entry.delete(0, tk.END)
        count_entry.insert(0, "10")  # 可以重新设置默认值
        length_entry.insert(0, "5")  # 可以重新设置默认值

def on_closing(input_window):
    input_window.destroy()

def text_segmentation(root, window):
    # 获取用户输入，如果用户关闭窗口则退出
    input_count_and_length(root)

    file_filter = [('Word 文档', '*.doc*')]
    file_path = get_path('打开文件', file_filter=file_filter)

    if not file_path:
        return None
    
    window.update_info('正在读取文档内容...')
    
    # 读取 Word 文档内容
    word_app = win32.gencache.EnsureDispatch('Word.Application')
    word_app.Visible = False
    doc = word_app.Documents.Open(file_path)
    text = doc.Content.Text

    # 加载用户自定义词典
    try:
        # 暂时还没有这个文件
        jieba.load_userdict('D:\\临时\\工具箱-对照表高亮术语测试\\custom_dict.txt')
        window.update_info("已加载自定义词典")
        print("已加载自定义词典")
    except FileNotFoundError as e:
        window.update_info("自定义词典文件未找到")
        print(f"{e} - 自定义词典文件未找到")

    window.update_info("正在统计词频...")

    # 使用 jieba 进行中文分词
    words = jieba.cut(text)

    directory = get_file_info(file_path, 'directory')
    write_to_txt(texts=words, directory=directory, filename='读取内容')

    # 统计词频
    word_counts = Counter(words)

    # 仅保留出现次数大于3的词，长度大于2的词
    if user_count > 1 and user_length > 1:
        window.update_info("正在筛选词频...")
        print(user_count, user_length)
    else:
        tkinter.messagebox.showerror("输入错误", "请输入有效的整数，且词频和长度均大于1")
        exit()

    # filtered_words = {word: count for word, count in word_counts.items() if count > 3 and len(word) > 1}
    filtered_words = [word + '\t' + str(count) 
                      for word, count in word_counts.items() 
                      if count >= user_count and len(word) >= user_length and contains_chinese(word)]
    word_list = [word 
                 for word, count in word_counts.items() 
                 if count >= user_count and len(word) >= user_length and contains_chinese(word)]

    # 获取工作路径
    directory = get_file_info(file_path, 'directory')
    
    # 将结果写入 txt 文件
    write_to_txt(filtered_words, directory, '词频统计.txt')

    print("开始高亮显示高频词...")
    result = highlight_terms_in_word(doc, word_list, window.update_info)
    # 询问是否在 Word 中标记
    # if tkinter.messagebox.askyesno('提示', '是否在 Word 中标记高频词？'):
    #     # 在 word 文档中高亮显示词频大于3的词
    #     result = highlight_terms_in_word(doc, word_list, window.update_info)
    # else:
    #     result = 0

    # 保存修改后的 Word 文档
    print("保存修改后的 Word 文档...")
    save_path = directory + '\\highlighted.docx'
    doc.SaveAs(save_path)
    doc.Close(False)  # 关闭文档，不保存更改

    word_app.Quit()  # 退出 Word 应用程序

    if result > 0:
        window.update_info(f"已完成。共标记了 {result} 个词")
    else:
        window.update_info(f"「词频统计」文件保存在文件目录下")