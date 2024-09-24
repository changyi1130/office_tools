import tkinter as tk
from tkinter import messagebox
from tkinter.ttk import Progressbar

class Progress_bar:
    def __init__(self, parent, total, title='当前进度'):
        self.total = total
        self.current = 0
        self.parent = parent
        self.title = title
        self.root = None
        self.progress_bar = None
        self.is_closed = False

    def start(self):
    # 创建并显示进度条窗口
        self.root = tk.Toplevel(self.parent)
        self.root.title(self.title)
        # self.root.geometry('300x100')
        self.root.resizable(False, False)

        # 提示信息
        self.label1 = tk.Label(self.root, text="正在处理中，请稍候...")
        self.label1.pack(pady=5)

        # 进度条
        self.progress_bar = Progressbar(self.root,
                                        orient='horizontal',
                                        length=280,
                                        mode='determinate')
        self.progress_bar.pack(padx=15, pady=5)

        # 进度信息
        self.label2 = tk.Label(self.root, text=f"当前进度：0 / {self.total}")
        self.label2.pack(pady=5)

        self.progress_bar['maximum'] = self.total
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing) # 当窗口关闭时的处理
        self.root.update_idletasks() # 更新界面
        # self.root.deiconify() # 显示窗口
        # self.root.mainloop()

    def update_progress(self, increment):
    # 更新进度条
        if self.is_closed:
            return

        self.current += increment
        self.progress_bar['value'] = self.current
        self.label2.config(text=f"当前进度：{self.current} / {self.total}")

        if self.current >= self.total:
            self.close()

    def close(self):
    # 关闭进度条窗口
        if not self.is_closed:
            self.is_closed = True
            self.root.destroy()
            self.parent.deiconify()  # 显示主窗口

    def on_closing(self):
    # 处理窗口关闭事件
        if messagebox.askokcancel("退出", "你确定要退出吗？"):
            self.close()