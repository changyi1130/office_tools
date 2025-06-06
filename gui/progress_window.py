import tkinter as tk
from tkinter import ttk

class ProgressWindow:
    """进度窗口"""
    def __init__(self, parent, title, total):
        self.win = tk.Toplevel(parent)
        self.win.title(title)

        self.progress = ttk.Progressbar(self.win, maximum=total)
        self.progress.pack(padx=20, pady=10)

        self.label = ttk.Label(self.win, text='0%')
        self.label.pack()

    def update(self, value):
        """更新进度"""
        self.progress['value'] = value
        self.label.config(text=f"{int(value/self.progress['value']*100)}%")
        self.win.update_idletasks() # 强制刷新界面

    def close(self):
        self.win.destroy()