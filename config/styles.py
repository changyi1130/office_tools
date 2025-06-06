from tkinter import ttk


class StyleManager:
    """集中管理所有 TTK 样式"""

    @staticmethod
    def setup_style(root):
        style = ttk.Style(root)

        # style.theme_use("clam")

        # style.configure(
        #     root,
        #     background="#f0f0f0",
        #     foreground="#000000",
        # )

        # 按钮样式
        style.configure(
            "TButton",
            padding=5,
            justify="center",
            width=12,
            # background="#EFF1F5",
            wraplength=80,  # 字符换行
        )

        # Labelframe 样式
        style.configure(
            "TLabelframe",
            padding=(5, 5, 5, 10)
        )

        style.configure(
            "TLabelframe.Label",
            relief="groove",
            font=("微软雅黑", 14, "bold")
        )
