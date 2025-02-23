from tkinter import ttk

class StyleManager:
    def __init__(self, colors):
        self.colors = colors
        
    def create_custom_style(self):
        """创建自定义样式"""
        style = ttk.Style()
        
        # 配置按钮样式
        style.configure(
            'Rounded.TButton',
            background=self.colors['button_bg'],
            foreground='white',
            font=('微软雅黑', 10),
            padding=5,
            relief='flat',
            borderwidth=0
        )
        style.map('Rounded.TButton',
            background=[('active', self.colors['button_hover']), ('pressed', self.colors['button_hover'])],
            relief=[('pressed', 'flat')]
        )
        
        # 配置Combobox样式
        style.configure(
            'Rounded.TCombobox',
            background=self.colors['frame_bg'],
            fieldbackground=self.colors['frame_bg'],
            foreground='black',
            arrowcolor=self.colors['text_color'],
            relief='flat',
            borderwidth=0
        )
        
        # 配置树形视图样式
        style.configure(
            "Rounded.Treeview",
            background=self.colors['frame_bg'],
            foreground="black",
            fieldbackground=self.colors['frame_bg'],
            font=('微软雅黑', 9),
            relief='flat',
            borderwidth=0,
            padding=5
        )
        style.map('Rounded.Treeview',
            background=[('selected', self.colors['tree_select'])],
            foreground=[('selected', 'black')]
        )
        
        # 配置树形视图标题样式
        style.configure(
            "Rounded.Treeview.Heading",
            background=self.colors['frame_bg'],
            foreground=self.colors['text_color'],
            font=('微软雅黑', 10, 'bold'),
            relief='flat',
            borderwidth=0
        )
        style.map('Rounded.Treeview.Heading',
            background=[('active', self.colors['tree_select'])]
        )
        
        # 配置进度条样式
        style.configure(
            'Rounded.Horizontal.TProgressbar',
            troughcolor=self.colors['frame_bg'],
            background=self.colors['text_color'],
            relief='flat',
            borderwidth=0
        ) 