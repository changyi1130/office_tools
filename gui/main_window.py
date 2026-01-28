import tkinter as tk
from tkinter import ttk
import threading
from queue import Queue

from config.buttons import BUTTON_GROUPS
from config.styles import StyleManager
from gui.tooltip import ToolTip
from gui.progress_window import ProgressWindow
from config.setting import Setting


class MainWindow(tk.Tk):
    """主窗口"""

    def __init__(self):
        super().__init__()
        setting = Setting()
        self.version = setting.version

        self._setup_window()

        # 按钮样式
        StyleManager.setup_style(self)

        # 导入按钮信息
        self.button_groups = BUTTON_GROUPS
        self._create_buttons()

        self._create_label_info()

        # 用于线程间通信的队列
        self.message_queue = Queue()

        # 定期检查消息队列
        self._check_message_queue()

        # 任务执行状态
        self.is_task_running = False

    def _setup_window(self):
        """窗口基础设置"""
        self.title("集装箱" + " V" + self.version)
        width, height = 520, 700
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        self.geometry(f"{width}x{height}+{(screen_width - width) // 2}+{(screen_height - height) // 2}")
        self.resizable(False, False)  # 禁止修改窗口大小

    def _create_buttons(self):
        """创建功能按钮"""
        # 主容器
        main_frame = ttk.Frame(self)
        main_frame.pack(fill="both", expand=True, padx=15, pady=(15, 10))

        # 动态生成按钮组
        for group in self.button_groups:
            # 分组容器
            group_frame = ttk.Labelframe(main_frame, text=group["name"])
            group_frame.pack(fill="x", expand=False, padx=5, pady=8)

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
        self.label_info_before_separator.pack(fill="x", padx=30, pady=(10, 8))

        # 信息标签
        self.label = ttk.Label(
            text="集装箱 " + self.version,
            justify='center',
            wraplength=450,
            font=("微软雅黑", 9)
        )
        self.label.pack(side="bottom", pady=(8, 15))

    def update_info(self, info_text):
        """更新提示信息"""
        self.label.config(text=info_text)
        self.label.update()

        # 同时将消息放入队列
        self.message_queue.put(("info", info_text))

    def _check_message_queue(self):
        """检查消息队列并更新界面"""
        try:
            while True:
                msg_type, msg_content = self.message_queue.get_nowait()

                if msg_type == "info":
                    # 更新主窗口信息标签
                    self.label.config(text=msg_content)
                    self.label.update()

                elif msg_type == "progress":
                    # 更新进度窗口
                    if hasattr(self, 'progress_window') and self.progress_window:
                        current, total, message = msg_content
                        # 如果提供了 total，更新进度窗口的总数
                        if total is not None and total > 0:
                            self.progress_window.total = total
                            self.progress_window.progress['maximum'] = total
                        # 如果提供了 current，更新进度
                        if current is not None:
                            self.progress_window.update(current, message)
                        # 只更新消息，不更新进度
                        elif message:
                            self.progress_window.add_message(message)

                elif msg_type == "complete":
                    # 标记进度窗口完成
                    if hasattr(self, 'progress_window') and self.progress_window:
                        self.progress_window.complete()
                    # 重置任务执行状态
                    self.is_task_running = False

                elif msg_type == "error":
                    # 显示错误信息
                    if hasattr(self, 'progress_window') and self.progress_window:
                        self.progress_window.add_message(f"错误: {msg_content}")
                        self.progress_window.complete()
                    else:
                        self.label.config(text=f"错误: {msg_content}")
                        self.label.update()
                    # 重置任务执行状态
                    self.is_task_running = False
        except:
            pass

        # 继续检查消息队列
        self.after(100, self._check_message_queue)

    def _create_button_command(self, btn_info):
        """绑定功能"""

        def wrapper():
            try:
                # 检查是否有任务正在执行
                if self.is_task_running:
                    self.message_queue.put(("error", "已有任务正在执行，请等待当前任务完成后再试。"))
                    return

                # 设置任务执行状态
                self.is_task_running = True

                # 创建进度窗口
                self.progress_window = ProgressWindow(self, btn_info["text"], 0)
                self.progress_window.add_message("准备开始处理...")

                # 在新线程中执行命令
                thread = threading.Thread(
                    target=self._execute_command,
                    args=(btn_info,),
                    daemon=True
                )
                thread.start()

            except Exception as e:
                # 统一错误处理
                self.message_queue.put(("error", str(e)))
                self.is_task_running = False

        return wrapper

    def _execute_command(self, btn_info):
        """在新线程中执行命令"""
        try:
            # 获取函数参数
            kwargs = btn_info.get("command_kwargs", {})

            # 添加进度更新回调函数
            def progress_callback(current=None, total=None, message=""):
                """进度更新回调函数
                
                参数:
                    current: 当前进度（可选）
                    total: 总数（可选）
                    message: 消息文本
                
                用法:
                    # 只更新消息，不更新进度
                    progress_callback(message="处理中...")
                    
                    # 更新进度和消息
                    progress_callback(1, 10, "处理第 1 个文件")
                """
                self.message_queue.put(("progress", (current, total, message)))

            kwargs["progress_callback"] = progress_callback

            # 调用命令函数
            btn_info["command"](**kwargs)

            # 标记任务完成
            self.message_queue.put(("complete", None))

        except Exception as e:
            # 统一错误处理
            self.message_queue.put(("error", str(e)))
