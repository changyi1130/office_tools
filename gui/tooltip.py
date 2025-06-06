import tkinter as tk


class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        x = self.widget.winfo_rootx() + 90  # 偏移量调整位置
        y = self.widget.winfo_rooty() + 45
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)  # 去除边框
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip, text=self.text,
                         bg="#ffffe0", fg="black",
                         border=1, relief="solid")
        label.pack(ipadx=5, ipady=2)

    def hide_tooltip(self, event):
        if self.tooltip:
            self.tooltip.destroy()
