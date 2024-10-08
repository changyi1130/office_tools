import tkinter as tk
import tkinter.messagebox
import functools
import time

# 比较 Word
from function.compare_documents import run_compare
# 提取 PDF 页数
from function.compute_pdf_pages import process_pdf_pages
# 提取 Word 页数
from function.compute_word_statistics import process_word_statistics
# 提取文件页数（PDF 和 Word）
from function.count_file_pages import count_file_pages

# 打开特殊字符表
from function.open_web import open_character

# 检查使用的字体
from function.get_used_fonts import process_fonts
# 高亮 Word 中的术语
from function.highlight_terms_in_word import run_highlight_term
# 标记高频词
from function.text_segmentation import text_segmentation
# 高亮修订内容
from function.highlight_revisions import run_highlight_revisions

# 转高、低版本
from function.convert_doc import convert_doc, convert_docx, convert_to_pdf

# 创建主应用程序类
class Main_App:
    def __init__(self, root):
        self.root = root
        self.root.title("集装箱")
        
        # 设置窗口大小并禁止调整
        width = 505
        height = 600
        screenwidth = self.root.winfo_screenwidth()
        screenheight = self.root.winfo_screenheight()
        self.root.geometry('%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2))
        self.root.resizable(False, False)

        # 设置背景色
        # root.configure(bg='white')
        
        # 创建菜单栏
        self.create_menu_bar()

        # 提示信息标签
        self.label_info = tk.Label(self.root,
                                   text="集装箱",
                                   font=('微软雅黑', 10))
        # self.label_info.configure(bg='white')
        self.label_info.pack(side='bottom', pady=(10, 20))
        self.create_separator('bottom')
    
    def create_menu_bar(self):
        """
        创建菜单栏
        """
        menu_bar = tk.Menu(self.root)
        # menu_bar.configure(bg='white')

        about_menu = tk.Menu(menu_bar, tearoff=0)
        about_menu.add_command(label="名称：集装箱")
        about_menu.add_command(label="作者：长意")

        date = str(time.strftime('%Y%m%d', time.localtime()))
        about_menu.add_command(label="版本：" + date)
        menu_bar.add_cascade(label="关于", menu=about_menu)
        self.root.config(menu=menu_bar)

    def create_label(self, label_text, label_font=('微软雅黑', 14, 'bold')):
        """
        创建分类标签
        
        :param label_text: 标签文本
        :param label_font: 标签字体
        """
        # 分类标签
        label = tk.Label(self.root, text=label_text, font=label_font)
        # label.configure(bg='white')
        label.pack(padx=10, pady=5)
    
    def create_button_frame(self):
        """
        创建按钮框架
        """
        button_frame = tk.Frame(self.root)
        # button_frame.configure(bg='white')
        button_frame.pack(padx=10, pady=5, fill='x')

        return button_frame

    def create_button_cb(self, button_frame, button_text, button_command):
        """
        创建按钮，共 5 列，传入 callback

        :param button_text: 按钮文本
        :param button_command: 按钮命令
        """
        new_button = tk.Button(button_frame,
                               text=button_text,
                               command=functools.partial(button_command,
                                                    callback=self.update_info),
                               width=11,
                               height=2,
                               wraplength=80,
                               relief='groove')
        # new_button.configure(bg='white')
        new_button.pack(side=tk.LEFT, padx=5)

    def create_button_root(self, button_frame, button_text, button_command, root, window):
        """
        创建按钮，共 5 列，传入 root 和 窗体

        :param button_text: 按钮文本
        :param button_command: 按钮命令
        """
        new_button = tk.Button(button_frame,
                               text=button_text,
                               command=functools.partial(button_command,
                                                         root=root,
                                                         window=window),
                               width=11,
                               height=2,
                               wraplength=80,
                               relief='groove')
        # new_button.configure(bg='white')
        new_button.pack(side=tk.LEFT, padx=5)
    
    def create_separator(self, side='top'):
        """
        创建分隔符
        """
        separator = tk.Frame(self.root, height=2, bd=1, relief='sunken')
        # separator.configure(bg='white')
        separator.pack(side=side, fill='x', padx=10, pady=(15, 5))
    
    def update_info(self, info_text):
        """
        更新提示信息

        :param info_text: 提示信息文本
        """
        self.label_info.config(text=info_text)
        self.label_info.update()
    
# 主程序
if __name__ == "__main__":
    root = tk.Tk()
    app = Main_App(root)

    # 添加功能
    # 第一部分
    app.create_label('统计信息')

    button_frame_1_1 = app.create_button_frame()
    app.create_button_cb(button_frame_1_1, '比较 Word', run_compare)
    app.create_button_cb(button_frame_1_1, '提取 PDF\n页数', process_pdf_pages)
    app.create_button_cb(button_frame_1_1, '提取 Word\n页数',
                      functools.partial(process_word_statistics,
                                        statistic_type=2,
                                        include_head_and_foot=False,
                                        callback=app.update_info))
    app.create_button_cb(button_frame_1_1, '提取文件\n页数', count_file_pages)

    # 第二部分
    app.create_separator()
    app.create_label('打开网页')
    button_frame_2_1 = app.create_button_frame()
    app.create_button_cb(button_frame_2_1, '特殊字符表', open_character)

    # 第三部分
    app.create_separator()
    app.create_label('其他功能')
    button_frame_3_1 = app.create_button_frame()
    app.create_button_cb(button_frame_3_1, '检查字体\n(测试)', process_fonts)
    app.create_button_cb(button_frame_3_1, '高亮术语\n(测试)', run_highlight_term)
    app.create_button_root(button_frame_3_1, '标记高频词\n(测试)', text_segmentation, root, app)
    app.create_button_cb(button_frame_3_1, '高亮修订内容\n(测试)', run_highlight_revisions)

    # 转存文件
    app.create_separator()
    app.create_label('转存文件')
    button_frame_4_1 = app.create_button_frame()
    app.create_button_cb(button_frame_4_1, '存为低版本\n(doc)', convert_docx)
    app.create_button_cb(button_frame_4_1, '存为高版本\n(docx)', convert_doc)
    app.create_button_cb(button_frame_4_1, '存为 PDF\n(pdf)', convert_to_pdf)

    # 执行
    print("程序开始")
    root.mainloop()

    print("程序结束")