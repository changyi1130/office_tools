from tkinter import ttk


class StyleManager:
    """ 集中管理所有 TTK 样式 """
    @staticmethod
    def setup_style(root):
        style = ttk.Style(root)

        # 按钮样式
        style.configure(
            "TButton",
            antialias=True,
            padding=5,
            anchor="center",
            width=11,
            height=2
        )

        # Labelframe 样式
        style.configure(
            "TLabelframe",
            padding=(5, 5, 5, 10)
        )

        style.configure(
            "TLabelframe.Label",
            antialias=True,
            font=("微软雅黑", 12, "bold")
        )