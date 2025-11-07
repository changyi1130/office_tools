import tkinter as tk
from tkinter import ttk

from config.buttons import BUTTON_GROUPS
from config.styles import StyleManager
from gui.tooltip import ToolTip


class MainWindow(tk.Tk):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.version = "V0.0.12"

        self._setup_window()

        # 按钮样式
        StyleManager.setup_style(self)

        # 导入按钮信息
        self.button_groups = BUTTON_GROUPS
        self._create_buttons()

        self._create_label_info()

    def _setup_window(self):
        """窗口基础设置"""
        self.title("集装箱")
        width, height = 505, 700
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        self.geometry(f"{width}x{height}+{(screen_width - width) // 2}+{(screen_height - height) // 2}")
        self.resizable(False, False)  # 禁止修改窗口大小

    def _create_buttons(self):
        """创建功能按钮"""
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

            cols = 4  # 每行按钮个数
            for idx, btn_info in enumerate(group["button"]):
                row = idx // cols
                col = idx % cols

                # 检查是否为占位符
                if btn_info.get("placeholder", False):
                    # 创建透明占位 Frame
                    placeholder = ttk.Frame(btn_container)
                    placeholder.grid(
                        row=row,
                        column=col,
                        padx=5,
                        pady=5,
                        sticky="ew"
                    )
                else:
                    # 创建正常按钮
                    btn = ttk.Button(
                        btn_container,
                        text=btn_info["text"],
                        command=self._create_button_command(btn_info)
                    )
                    ToolTip(btn, btn_info["tip"])

                    # 按钮间距
                    btn.grid(
                        row=row,
                        column=col,
                        padx=5,
                        pady=5,
                        sticky="ew",  # 水平拉伸
                        ipadx=2,  # 内边距
                        ipady=2
                    )

    def _create_label_info(self):
        """创建信息提示标签"""
        # 添加一条分隔符
        self.label_info_before_separator = ttk.Separator(self, orient='horizontal')
        self.label_info_before_separator.pack(fill="x", padx=30, pady=(15, 5))

        # 信息标签
        self.label = ttk.Label(
            text="集装箱 " + self.version,
            justify='center',
            wraplength=400
        )
        self.label.pack(side="bottom", pady=(10, 20))

    def update_info(self, info_text):
        """更新提示信息"""
        self.label.config(text=info_text)
        self.label.update()

    def _create_button_command(self, btn_info):
        """绑定功能"""

        def wrapper():

            try:
                # 获取函数参数
                kwargs = btn_info.get("command_kwargs", {})

                # 添加 update_info 参数（如有）
                if "update_info" in btn_info["command"].__code__.co_varnames:
                    kwargs["update_info"] = self.update_info

                # 调用命令函数
                btn_info["command"](**kwargs)
            except Exception as e:
                # 统一错误处理
                self.update_info(f"执行错误：{str(e)}")

        return wrapper
