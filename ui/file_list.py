from tkinter import ttk
import os

class FileListManager:
    def __init__(self, parent, columns, colors):
        self.parent = parent
        self.columns = columns
        self.colors = colors
        self.all_files = []
        self.setup_treeview()
        
    def setup_treeview(self):
        # 创建滚动条容器
        self.scroll_frame = ttk.Frame(self.parent)
        self.scroll_frame.pack(fill='both', expand=True)
        
        # 创建树形视图
        self.tree = ttk.Treeview(
            self.scroll_frame,
            columns=list(self.columns.keys()),
            show="headings",
            style="Rounded.Treeview"
        )
        
        # 设置滚动条
        self.setup_scrollbars()
        
        # 设置列
        self.setup_columns()
        
    def setup_scrollbars(self):
        # 添加垂直滚动条
        y_scrollbar = ttk.Scrollbar(
            self.scroll_frame, 
            orient='vertical', 
            command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=y_scrollbar.set)
        
        # 添加水平滚动条
        x_scrollbar = ttk.Scrollbar(
            self.scroll_frame, 
            orient='horizontal', 
            command=self.tree.xview
        )
        self.tree.configure(xscrollcommand=x_scrollbar.set)
        
        # 使用网格布局
        self.tree.grid(row=0, column=0, sticky='nsew')
        y_scrollbar.grid(row=0, column=1, sticky='ns')
        x_scrollbar.grid(row=1, column=0, sticky='ew')
        
        # 配置网格权重
        self.scroll_frame.grid_rowconfigure(0, weight=1)
        self.scroll_frame.grid_columnconfigure(0, weight=1) 