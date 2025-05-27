import tkinter as tk
from tkinter import ttk, messagebox

from gui.progress_window import ProgressWindow
from config.buttons import BUTTON_GROUPS
from config.styles import StyleManager


class MainWindow(tk.Tk):
    """ 主窗口 """
    def __init__(self):
        super().__init__()
        self.version = "V0.0.10"

        self._setup_window()

        # 按钮样式
        StyleManager.setup_style(self)

        # 导入按钮信息
        self.button_groups = BUTTON_GROUPS
        self._create_buttons()

        self._create_info_label()

    def _setup_window(self):
        self.title("集装箱")
        width, height = 505, 600
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        self.geometry('%dx%d+%d+%d' % (
            width, height, (screenwidth - width) / 2, (screenheight - height) / 2))
        self.resizable(False, False)

    def _create_buttons(self):
        """ 创建功能按钮 """
        # 主容器
        main_frame = ttk.Frame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # 动态生成按钮组
        for group in self.button_groups:
            # 分组容器
            group_frame = ttk.Labelframe(main_frame, text=group["name"])
            group_frame.pack(fill="x", expand=False, padx=5, pady=5)

            # 内部容器管理换行
            btn_container = ttk.Frame(group_frame)
            btn_container.pack(fill="x", expand=True)

            cols = 4 # 每行按钮个数
            for idx, btn_info in enumerate(group["button"]):
                row = idx // cols
                col = idx % cols

                btn = ttk.Button(
                    btn_container,
                    text=btn_info["text"],
                    command=btn_info["command"],
                )

                # 按钮间距
                btn.grid(
                    row=row,
                    column=col,
                    padx=5,
                    pady=5,
                    sticky="ew",    # 水平拉伸
                    ipadx=2,        # 内边距
                    ipady=2
                )

            # 设置按钮容器自动换行
            # btn_container.grid_columnconfigure(0, weight=1)
            # for col in range(cols):
            #     btn_container.grid_columnconfigure(col, weight=1)

    def _create_info_label(self):
        """ 创建信息提示标签 """
        # 添加一条分隔符
        self.label_info_before_separator = ttk.Separator(self, orient='horizontal')
        self.label_info_before_separator.pack(fill="x", padx=30, pady=(15, 5))

        # 信息标签
        self.label = ttk.Label(
            text="集装箱 " + self.version,
            wraplength=400
        )
        self.label.pack(side="bottom", pady=(10, 20))

        # 进度条
        self.progress = ttk.Progressbar(self)
        self.progress.pack(padx=50, pady=10, fill="x")
        self.progress.pack_forget()

    def update_info(self, info_text, progress):
        """ 更新提示信息 """
        self.label.text = info_text
        self.progress.value = progress

        self.label.update()
        self.progress.update()