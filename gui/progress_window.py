import tkinter as tk
from tkinter import ttk


class ProgressWindow:
    """进度窗口"""

    def __init__(self, parent, title, total):
        self.win = tk.Toplevel(parent)
        self.win.title(title)
        self.win.geometry("600x400")
        self.win.resizable(False, False)

        # 设置窗口为模态窗口
        self.win.transient(parent)  # 设置为主窗口的子窗口
        self.win.grab_set()  # 捕获所有事件，阻止主窗口操作
        self.win.focus_set()  # 设置焦点
        
        # 窗口居中显示
        self._center_window()

        # 进度条容器
        progress_frame = ttk.Frame(self.win)
        progress_frame.pack(fill="x", padx=30, pady=(20, 10))
        
        # 进度百分比标签
        # self.percent_label = ttk.Label(progress_frame, text='0%', font=("微软雅黑", 11, "bold"))
        # self.percent_label.pack(side="right", padx=(10, 0))
        
        # 进度条
        self.progress = ttk.Progressbar(progress_frame, maximum=total, length=400)
        self.progress.pack(side="left", fill="x", expand=True)

        # 创建文本框和滚动条
        text_frame = ttk.Frame(self.win)
        text_frame.pack(fill="both", expand=True, padx=30, pady=10)

        self.text = tk.Text(text_frame, height=15, state="disabled", font=("微软雅黑", 9),
                           relief="solid", borderwidth=1, wrap="none")
        scrollbar = ttk.Scrollbar(text_frame, command=self.text.yview)
        self.text.configure(yscrollcommand=scrollbar.set)

        self.text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 关闭按钮
        button_frame = ttk.Frame(self.win)
        button_frame.pack(pady=(10, 20))
        self.close_button = ttk.Button(button_frame, text="关闭", command=self.close, state="disabled", width=12)
        self.close_button.pack()

        # 记录总数和当前数
        self.total = total
        self.current = 0

        # 记录是否完成
        self.completed = False
    
    def _center_window(self):
        """将窗口居中显示在屏幕上"""
        self.win.update_idletasks()
        width = self.win.winfo_width()
        height = self.win.winfo_height()
        x = (self.win.winfo_screenwidth() // 2) - (width // 2)
        y = (self.win.winfo_screenheight() // 2) - (height // 2)
        self.win.geometry(f'{width}x{height}+{x}+{y}')

    def update(self, value, message=""):
        """更新进度"""
        self.current = value
        self.progress['value'] = value
        # 修复进度计算，确保除数不为 0
        if self.total > 0:
            percent = int(value / self.total * 100)
        else:
            percent = 0
        # self.percent_label.config(text=f"{percent}%")

        if message:
            self.add_message(message)

        self.win.update_idletasks()  # 强制刷新界面

    def add_message(self, message):
        """添加消息到文本框"""
        self.text.config(state="normal")
        self.text.insert("end", message + "\n")
        self.text.see("end")  # 滚动到最新消息
        self.text.config(state="disabled")

    def complete(self):
        """标记任务完成"""
        self.completed = True
        self.close_button.config(state="normal")
        self.add_message("任务已结束")

    def close(self):
        """关闭窗口"""
        self.win.destroy()
