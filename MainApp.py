import tkinter as tk
import tkinter.messagebox
import functools

# 导入功能
from import_functions import *

class MainApp:
    def __init__(self, root):
        self.root = root
        self.version = 'V0.0.8'
        self.setup_window()
        self.create_menu_bar()
        self.label_info = self.create_info_label()
        self.build_ui()

    def setup_window(self):
        """设置主窗口基本信息"""
        self.root.title("集装箱 " + self.version)
        width, height = 505, 600
        screenwidth = self.root.winfo_screenwidth()
        screenheight = self.root.winfo_screenheight()
        self.root.geometry('%dx%d+%d+%d' % (
            width, height, (screenwidth - width) / 2, (screenheight - height) / 2))
        self.root.resizable(False, False)
    
    def create_menu_bar(self):
        """创建菜单栏"""
        menu_bar = tk.Menu(self.root)
        # menu_bar.configure(bg='white')
        about_menu = tk.Menu(menu_bar, tearoff=0)

        about_menu.add_command(label="名称：集装箱")
        about_menu.add_command(label="作者：长意")
        about_menu.add_command(label="版本：" + self.version)

        menu_bar.add_cascade(label="关于", menu=about_menu)
        self.root.config(menu=menu_bar)

    def create_info_label(self):
        """创建信息提示标签"""
        label = tk.Label(self.root, 
                         text="集装箱 " + self.version, 
                         wraplength = 480,
                         font=('微软雅黑', 9))
        label.pack(side='bottom', pady=(10, 20))
        return label
    
    def update_info(self, info_text):
        """更新提示信息"""
        self.label_info.config(text=info_text)
        self.label_info.update()

    def create_label(self, label_text, label_font=('微软雅黑', 14, 'bold')):
        """创建分类标签"""
        label = tk.Label(self.root, text=label_text, font=label_font)
        label.pack(padx=10, pady=5)
    
    def create_button_frame(self):
        """创建按钮框架"""
        button_frame = tk.Frame(self.root)
        button_frame.pack(padx=10, pady=5, fill='x')
        return button_frame

    def create_button(self, button_frame, button_text, button_command, *args):
        """创建按钮"""
        new_button = tk.Button(button_frame,
                               text=button_text,
                               command=lambda: button_command(*args),
                               width=11,
                               height=2,
                               wraplength=80,
                               relief='groove')
        new_button.pack(side=tk.LEFT, padx=5)

    def create_separator(self, side='top'):
        """创建分隔符"""
        separator = tk.Frame(self.root, height=2, bd=1, relief='sunken')
        # separator.configure(bg='white')
        separator.pack(side=side, fill='x', padx=10, pady=(15, 5))
    
    def build_ui(self):
        """构建用户界面"""
        # 第一部分
        self.create_label('统计信息')
        frame_1_1 = self.create_button_frame()
        self.create_button(frame_1_1, '提取页数', count_file_pages, self.update_info)
        self.create_button(frame_1_1, '检查字体\n(测试)', process_check_text, self.update_info)
        self.create_button(frame_1_1, '添加编号\n(测试)', process_add_index, self.update_info)
        self.create_button(frame_1_1, '删除编号\n(测试)', process_del_index, self.update_info)
        self.create_separator()

        # 第二部分
        self.create_label('文档处理')
        frame_3_1 = self.create_button_frame()
        self.create_button(frame_3_1, '比较 Word', select_and_compare_doc, self.update_info)
        self.create_button(frame_3_1, '高亮修订内容', process_highlight_revisions, self.update_info)
        self.create_button(frame_3_1, '高亮术语\n(测试)', process_highlight_term, self.update_info)
        self.create_button(frame_3_1, '分析词频\n(测试)', text_segmentation, self.root, self.update_info)

        # 第三部分
        self.create_separator()
        self.create_label('转存文件')
        frame_4_1 = self.create_button_frame()
        self.create_button(frame_4_1, '存为低版本\n(doc)', convert_docx, self.update_info)
        self.create_button(frame_4_1, '存为高版本\n(docx)', convert_doc, self.update_info)
        self.create_button(frame_4_1, '存为 PDF\n(pdf)', convert_to_pdf, self.update_info)
        self.create_separator()
    
        # 第四部分
        self.create_label('打开网页')
        frame_2_1 = self.create_button_frame()
        self.create_button(frame_2_1, '特殊字符表', open_character)

# 主程序
if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)

    # 执行
    print("程序开始")
    root.mainloop()

    print("程序结束")