import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import glob
from datetime import datetime
import win32com.client
from win32com.shell import shell, shellcon
import pythoncom
from pathlib import Path
import threading
from queue import Queue
import time
import queue
import urllib.request
from win32gui import CreateRoundRectRgn, SetWindowRgn
from ui.styles import StyleManager
from ui.file_list import FileListManager
from core.file_search import FileSearcher
from core.config import ConfigManager

class FileOrganizer:
    def __init__(self, root):
        self.root = root
        self.root.title("✨文件小助手✨")
        self.root.geometry("1000x700")
        
        # 更新主题颜色为更和谐的配色
        self.colors = {
            'bg': '#FDF6F9',           # 更浅的背景粉色
            'frame_bg': '#FFFFFF',     # 纯白色背景
            'button_bg': '#F8BBD0',    # 柔和的粉色按钮
            'button_hover': '#F48FB1', # 按钮悬停颜色
            'text_color': '#EC407A',   # 文字主色
            'border_color': '#F8BBD0', # 边框颜色
            'tree_select': '#FCE4EC'   # 树形列表选中颜色
        }
        
        # 设置窗口样式
        self.root.configure(bg=self.colors['bg'])  # 使用背景色
        
        # 创建主容器框架
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 设置窗口圆角
        self.root.overrideredirect(True)  # 移除默认的窗口边框
        self.make_rounded()
        
        # 存储选择的目录
        self.selected_dirs = []
        
        # 支持的文件类型
        self.file_types = {
            "📝 Word文件": "*.doc*",
            "📊 Excel文件": "*.xls*",
            "📑 PPT文件": "*.ppt*"
        }
        
        # 创建自定义样式
        self.style_manager = StyleManager(self.colors)
        self.style_manager.create_custom_style()
        
        # 创建标题栏
        self.create_title_bar()
        
        self.setup_window()
        self.init_components()
        self.setup_ui()
        
        # 显示窗口
        self.root.after(100, lambda: self.root.attributes('-alpha', 1.0))
        
        # 添加窗口状态标记
        self.is_maximized = False
        self.normal_size = None
        
        # 添加排序状态记录
        self.sort_column = None  # 当前排序的列
        self.sort_reverse = False  # 排序方向
        
        # 添加文件列表存储
        self.all_files = []  # 存储所有文件的ID
        
        # 绑定窗口大小变化事件
        self.root.bind('<Configure>', lambda e: self.on_window_configure(e))
        
        # 初始化配置管理器
        self.config_manager = ConfigManager()
        
        # 设置窗口位置和大小
        size = self.config_manager.config['last_window_size']
        position = self.config_manager.config['last_window_position']
        self.root.geometry(f"{size}{position}")
        
        # 加载保存的目录
        self.load_saved_directories()
        
        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def setup_window(self):
        """设置窗口基本属性"""
        self.root.configure(bg=self.colors['bg'])  # 使用背景色
        self.root.overrideredirect(True)  # 移除默认的窗口边框
        self.make_rounded()
        
    def init_components(self):
        """初始化组件"""
        self.file_searcher = FileSearcher(self.handle_search_result)
        
    def handle_search_result(self, msg_type, data):
        """处理搜索结果的回调函数"""
        try:
            if msg_type == "file":
                # 添加文件到列表
                item_id = self.tree.insert("", tk.END, values=(
                    self.get_file_icon(data["type"]),  # 添加文件图标
                    data["name"],
                    data["type"],
                    data["size"],
                    data["created"],
                    data["modified"],
                    data["path"]
                ))
                self.all_files.append(item_id)
                
            elif msg_type == "done":
                self.completed_dirs += 1
                progress = f"正在搜索... ({self.completed_dirs}/{self.total_dirs})"
                self.progress_var.set(progress)
                
                if self.completed_dirs >= self.total_dirs:
                    self.progress_bar.stop()
                    self.progress_bar.pack_forget()
                    self.progress_var.set(f"搜索完成，共找到 {len(self.tree.get_children())} 个文件")
                    self.searching = False
                    # 启用文件类型选择
                    self.file_type_combo.configure(state="readonly")
                
            elif msg_type == "error":
                messagebox.showerror("错误", data)
                
        except Exception as e:
            messagebox.showerror("错误", f"处理搜索结果时出错: {str(e)}")

    def setup_ui(self):
        """设置用户界面"""
        # 创建主标题
        title_frame = ttk.Frame(self.main_frame)  # 改用main_frame
        title_frame.pack(fill=tk.X, pady=10)
        ttk.Label(
            title_frame, 
            text="🌸 文件整理小助手 🌸",
            font=('微软雅黑', 16, 'bold'),
            foreground=self.colors['text_color']
        ).pack()
        
        # 创建左侧面板
        left_frame = ttk.Frame(self.main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=20, pady=5)
        
        # 目录管理标题
        ttk.Label(
            left_frame,
            text="📁 目录管理",
            font=('微软雅黑', 12),
            foreground=self.colors['text_color']
        ).pack(pady=5)
        
        # 自定义按钮样式
        style = ttk.Style()
        style.configure(
            'Custom.TButton',
            background=self.colors['button_bg'],
            foreground='white',
            font=('微软雅黑', 10),
            padding=5
        )
        
        # 添加目录按钮
        ttk.Button(
            left_frame,
            text="➕ 添加目录",
            style='Rounded.TButton',
            command=self.add_directory
        ).pack(pady=5, fill=tk.X)
        
        # 目录列表
        self.dir_listbox = tk.Listbox(
            left_frame,
            width=40,
            height=10,
            font=('微软雅黑', 10),
            bg=self.colors['frame_bg'],
            selectmode=tk.SINGLE,
            relief='flat',
            borderwidth=0,
            highlightthickness=0,  # 移除焦点边框
            activestyle='none',  # 移除选中项的下划线
            selectbackground=self.colors['tree_select'],  # 设置选中项的背景色
            selectforeground='black'  # 设置选中项的文字颜色
        )
        self.dir_listbox.pack(pady=5, padx=2)
        
        # 删除选中目录按钮
        ttk.Button(
            left_frame,
            text="❌ 删除选中",
            style='Rounded.TButton',
            command=self.remove_directory
        ).pack(pady=5, fill=tk.X)
        
        # 创建右侧面板
        right_frame = ttk.Frame(self.main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        # 控制面板
        control_frame = ttk.Frame(right_frame)
        control_frame.pack(fill=tk.X, pady=10)
        
        # 搜索框
        search_frame = ttk.Frame(control_frame)
        search_frame.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(
            search_frame,
            text="🔍",
            font=('微软雅黑', 10),
            foreground=self.colors['text_color']
        ).pack(side=tk.LEFT)
        
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.on_search_change)  # 绑定变化事件
        
        self.search_entry = ttk.Entry(
            search_frame,
            textvariable=self.search_var,
            width=20,
            font=('微软雅黑', 10)
        )
        self.search_entry.pack(side=tk.LEFT, padx=5)
        
        # 文件类型过滤
        ttk.Label(
            control_frame,
            text="🔍 文件类型：",
            font=('微软雅黑', 10),
            foreground=self.colors['text_color']
        ).pack(side=tk.LEFT, padx=5)
        
        self.file_type_var = tk.StringVar(value="全部")
        file_type_options = ["✨ 全部"] + list(self.file_types.keys())
        self.file_type_combo = ttk.Combobox(
            control_frame,
            textvariable=self.file_type_var,
            values=file_type_options,
            state="readonly",
            font=('微软雅黑', 10),
            width=15,
            style='Rounded.TCombobox'
        )
        self.file_type_combo.pack(side=tk.LEFT)
        
        # 进度显示
        self.progress_var = tk.StringVar(value="💝 准备就绪")
        self.progress_label = ttk.Label(
            control_frame,
            textvariable=self.progress_var,
            font=('微软雅黑', 10),
            foreground=self.colors['text_color']
        )
        self.progress_label.pack(side=tk.LEFT, padx=5)
        
        self.progress_bar = ttk.Progressbar(
            control_frame,
            mode='indeterminate',
            length=150,
            style='Custom.Horizontal.TProgressbar'
        )
        
        # 创建滚动条容器
        scroll_frame = ttk.Frame(right_frame)
        scroll_frame.pack(fill=tk.BOTH, expand=True)
        
        # 文件列表
        self.tree = ttk.Treeview(
            scroll_frame,
            columns=("图标", "名称", "类型", "大小", "创建时间", "修改时间", "路径"),
            show="headings",
            style="Rounded.Treeview"
        )
        
        # 保存列配置为实例变量
        self.columns = {
            "图标": ("", 30),  # 新增图标列，宽度30像素，无标题
            "名称": ("📄 文件名", 200),
            "类型": ("📎 类型", 80),
            "大小": ("📦 大小", 100),
            "创建时间": ("📅 创建时间", 150),
            "修改时间": ("🕒 修改时间", 150),
            "路径": ("📂 路径", 300)
        }
        
        # 添加垂直滚动条
        y_scrollbar = ttk.Scrollbar(scroll_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=y_scrollbar.set)
        
        # 添加水平滚动条
        x_scrollbar = ttk.Scrollbar(scroll_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(xscrollcommand=x_scrollbar.set)
        
        # 使用网格布局
        self.tree.grid(row=0, column=0, sticky='nsew')
        y_scrollbar.grid(row=0, column=1, sticky='ns')
        x_scrollbar.grid(row=1, column=0, sticky='ew')
        
        # 配置网格权重
        scroll_frame.grid_rowconfigure(0, weight=1)
        scroll_frame.grid_columnconfigure(0, weight=1)
        
        # 设置列和绑定事件
        for col, (text, width) in self.columns.items():
            if col != "图标":  # 图标列不需要排序功能
                self.tree.heading(col, text=text, command=lambda c=col: self.sort_treeview(c))
            else:
                self.tree.heading(col, text="")  # 图标列不显示标题
            self.tree.column(col, width=width, minwidth=20, stretch=False)  # 禁用自动拉伸
        
        # 绑定列宽调整事件
        self.tree.bind('<Configure>', self.on_tree_configure)
        self.tree.bind('<Button-1>', self.on_click)
        self.tree.bind('<ButtonRelease-1>', self.on_release)  # 添加鼠标释放事件
        self.is_resizing = False
        self.last_x = 0
        
        # 保存初始列宽比例
        self.column_ratios = {}
        total_width = sum(width for _, width in self.columns.values())
        for col, (_, width) in self.columns.items():
            self.column_ratios[col] = width / total_width
        
        # 修改文件类型过滤事件绑定
        self.file_type_combo.bind("<<ComboboxSelected>>", lambda e: self.filter_files())
        
        # 绑定双击事件
        self.tree.bind('<Double-Button-1>', self.open_file)
        
    def load_saved_directories(self):
        """加载保存的目录并开始搜索"""
        directories = self.config_manager.get_directories()
        for directory in directories:
            if os.path.exists(directory):
                self.selected_dirs.append(directory)
                self.dir_listbox.insert(tk.END, directory)
        
        # 如果有保存的目录，开始搜索
        if self.selected_dirs:
            self.refresh_files()
    
    def add_directory(self):
        directory = filedialog.askdirectory()
        if directory and directory not in self.selected_dirs:
            self.selected_dirs.append(directory)
            self.dir_listbox.insert(tk.END, directory)
            self.config_manager.add_directory(directory)  # 保存到配置
            self.search_directory(directory)
    
    def remove_directory(self):
        selection = self.dir_listbox.curselection()
        if selection:
            index = selection[0]
            directory = self.selected_dirs[index]
            self.selected_dirs.pop(index)
            self.dir_listbox.delete(index)
            self.config_manager.remove_directory(directory)  # 从配置中移除
            self.refresh_files()
            
    def get_file_size(self, size_bytes):
        """将文件大小转换为人类可读格式"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024:
                return f"{size_bytes:.2f}{unit}"
            size_bytes /= 1024
        return f"{size_bytes:.2f}TB"
    
    def sort_treeview(self, col):
        """根据列头排序"""
        # 如果点击的是当前排序列，则反转排序方向
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = False
        
        # 获取所有项目
        items = [(self.tree.set(item, col), item) for item in self.tree.get_children("")]
        
        # 根据不同列类型进行排序
        if col in ["创建时间", "修改时间"]:
            items.sort(key=lambda x: datetime.strptime(x[0], "%Y-%m-%d %H:%M"), reverse=self.sort_reverse)
        elif col == "大小":
            def convert_size(size_str):
                units = {'B': 1, 'KB': 1024, 'MB': 1024**2, 'GB': 1024**3, 'TB': 1024**4}
                number = float(size_str[:-2])
                unit = size_str[-2:]
                return number * units[unit]
            items.sort(key=lambda x: convert_size(x[0]), reverse=self.sort_reverse)
        else:
            items.sort(key=lambda x: x[0].lower(), reverse=self.sort_reverse)
        
        # 重新插入排序后的项目
        for index, (val, item) in enumerate(items):
            self.tree.move(item, "", index)
        
        # 更新列头显示
        for column in ["名称", "类型", "大小", "创建时间", "修改时间", "路径"]:
            if column == col:
                # 显示排序指示器
                indicator = " ▼" if self.sort_reverse else " ▲"
                text = self.columns[column][0] + indicator
            else:
                # 移除其他列的排序指示器
                text = self.columns[column][0]
            self.tree.heading(column, text=text)

    def filter_files(self):
        """根据选择的文件类型筛选当前列表"""
        print("\n=== 开始筛选文件 ===")
        
        # 获取选择的文件类型
        selected_type = self.file_type_var.get()
        print(f"选择的文件类型: {selected_type}")
        
        if selected_type == "✨ 全部":
            selected_type = "全部"
        
        print(f"总文件数: {len(self.all_files)}")
        
        # 先将所有项目重新附加到树中
        for item in self.all_files:
            try:
                self.tree.reattach(item, "", "end")
            except:
                try:
                    self.tree.move(item, "", "end")
                except:
                    print(f"无法重新附加项目: {item}")
        
        # 如果不是"全部"，则隐藏不匹配的文件
        if selected_type != "全部":
            print(f"\n开始按类型 {selected_type} 筛选...")
            hidden_count = 0
            shown_count = 0
            
            for item in self.all_files:
                try:
                    values = self.tree.item(item)['values']
                    if not values:
                        print(f"警告: 项目 {item} 没有值")
                        continue
                    
                    file_type = values[2].lower()  # 获取文件类型（第三列）
                    print(f"检查文件: {values[1]}, 类型: {file_type}")
                    
                    should_show = False
                    # 根据文件扩展名判断类型
                    if selected_type == "📝 Word文件" and file_type in [".doc", ".docx"]:
                        should_show = True
                    elif selected_type == "📊 Excel文件" and file_type in [".xls", ".xlsx"]:
                        should_show = True
                    elif selected_type == "📑 PPT文件" and file_type in [".ppt", ".pptx"]:
                        should_show = True
                    
                    if should_show:
                        shown_count += 1
                        print(f"  显示文件: {values[1]}")
                    else:
                        hidden_count += 1
                        print(f"  隐藏文件: {values[1]}")
                        self.tree.detach(item)
                except Exception as e:
                    print(f"处理项目时出错: {e}")
            
            print(f"\n筛选结果:")
            print(f"显示的文件数: {shown_count}")
            print(f"隐藏的文件数: {hidden_count}")
        
        print("\n=== 筛选完成 ===")

    def process_search_results(self):
        """处理搜索结果队列"""
        try:
            while True:
                if not hasattr(self, 'search_queue') or not hasattr(self, 'searching'):
                    break
                    
                if not self.searching:
                    break
                    
                try:
                    msg_type, data = self.search_queue.get_nowait()
                    
                    if msg_type == "file":
                        # 添加文件到列表
                        stats = os.stat(data)
                        file_type = os.path.splitext(data)[1]
                        file_info = {
                            "path": data,
                            "name": os.path.basename(data),
                            "type": file_type,
                            "size": self.get_file_size(stats.st_size),
                            "created": datetime.fromtimestamp(stats.st_ctime).strftime("%Y-%m-%d %H:%M"),
                            "modified": datetime.fromtimestamp(stats.st_mtime).strftime("%Y-%m-%d %H:%M")
                        }
                        item_id = self.tree.insert("", tk.END, values=(
                            self.get_file_icon(file_type),  # 添加文件图标
                            file_info["name"],
                            file_info["type"],
                            file_info["size"],
                            file_info["created"],
                            file_info["modified"],
                            file_info["path"]
                        ))
                        self.all_files.append(item_id)
                        
                    elif msg_type == "done":
                        self.completed_dirs += 1
                        progress = f"正在搜索... ({self.completed_dirs}/{self.total_dirs})"
                        self.progress_var.set(progress)
                        
                    elif msg_type == "error":
                        messagebox.showerror("错误", data)
                        
                except queue.Empty:
                    pass
                    
                if self.completed_dirs >= self.total_dirs:
                    self.progress_bar.stop()
                    self.progress_bar.pack_forget()
                    self.progress_var.set(f"搜索完成，共找到 {len(self.tree.get_children())} 个文件")
                    self.searching = False
                    # 启用文件类型选择
                    self.file_type_combo.configure(state="readonly")
                    break
                    
                self.root.update()
                time.sleep(0.01)
                
        except Exception as e:
            messagebox.showerror("错误", f"处理搜索结果时出错: {str(e)}")
        finally:
            self.searching = False
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            # 确保控件被重新启用
            self.file_type_combo.configure(state="readonly")

    def refresh_files(self):
        """清空列表并重新搜索所有目录"""
        # 清空现有项目和存储
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.all_files.clear()
        
        if not self.selected_dirs:
            return
            
        # 准备搜索
        self.searching = True
        self.completed_dirs = 0
        self.total_dirs = len(self.selected_dirs)
        self.search_queue = Queue()
        
        # 禁用文件类型选择
        self.file_type_combo.configure(state="disabled")
        
        # 显示进度条
        self.progress_var.set("正在搜索...")
        self.progress_bar.pack(side=tk.LEFT, padx=5)
        self.progress_bar.start(10)
        
        # 搜索所有支持的文件类型
        patterns = ["doc", "docx", "xls", "xlsx", "ppt", "pptx"]
        
        # 为每个目录启动搜索线程
        search_threads = []
        for directory in self.selected_dirs:
            thread = threading.Thread(
                target=self.search_files_thread,
                args=(directory, patterns, self.search_queue),
                daemon=True
            )
            search_threads.append(thread)
            thread.start()
        
        # 启动结果处理
        process_thread = threading.Thread(
            target=self.process_search_results,
            daemon=True
        )
        process_thread.start()

    def search_files_thread(self, directory, patterns, search_queue):
        """在线程中执行文件搜索"""
        try:
            for pattern in patterns:
                for file in Path(directory).rglob(f"*.{pattern}"):
                    try:
                        file_path = str(file)
                        stats = os.stat(file_path)
                        search_queue.put(("file", file_path))
                    except Exception as e:
                        print(f"处理文件 {file_path} 时出错: {e}")
                        continue
            
            # 搜索完成
            search_queue.put(("done", directory))
            
        except Exception as e:
            search_queue.put(("error", f"搜索目录 {directory} 时出错: {str(e)}"))

    def make_rounded(self):
        """创建圆角窗口"""
        try:
            # 确保窗口尺寸已更新
            self.root.update_idletasks()
            
            # 获取实际窗口尺寸
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            
            # 重新创建圆角区域
            from win32gui import SetWindowRgn, CreateRoundRectRgn
            region = CreateRoundRectRgn(0, 0, width + 1, height + 1, 20, 20)  # 添加1像素避免边缘问题
            hwnd = self.root.winfo_id()
            SetWindowRgn(hwnd, region, True)
            
            # 刷新窗口
            self.root.update_idletasks()
            
        except Exception as e:
            print(f"设置圆角窗口失败: {e}")

    def create_title_bar(self):
        """创建自定义标题栏"""
        title_bar = ttk.Frame(self.main_frame)
        title_bar.pack(fill='x', pady=(0, 10))
        
        # 标题
        title_label = ttk.Label(
            title_bar,
            text="✨文件小助手✨",
            font=('微软雅黑', 12),
            foreground=self.colors['text_color']
        )
        title_label.pack(side='left', padx=10)
        
        # 关闭按钮
        close_button = ttk.Button(
            title_bar,
            text="✖",
            style='Rounded.TButton',
            command=self.root.quit,
            width=3
        )
        close_button.pack(side='right', padx=5)
        
        # 最大化/还原按钮
        self.maximize_button = ttk.Button(
            title_bar,
            text="□",
            style='Rounded.TButton',
            command=self.toggle_maximize,
            width=3
        )
        self.maximize_button.pack(side='right', padx=2)
        
        # 最小化按钮
        minimize_button = ttk.Button(
            title_bar,
            text="—",
            style='Rounded.TButton',
            command=self.root.iconify,
            width=3
        )
        minimize_button.pack(side='right', padx=2)
        
        # 绑定拖动事件
        title_bar.bind('<Button-1>', self.start_move)
        title_bar.bind('<B1-Motion>', self.do_move)
        # 双击标题栏最大化/还原
        title_bar.bind('<Double-Button-1>', lambda e: self.toggle_maximize())

    def start_move(self, event):
        """开始移动窗口"""
        self.x = event.x
        self.y = event.y

    def do_move(self, event):
        """移动窗口"""
        if not self.is_maximized:  # 只在非最大化状态下允许移动
            deltax = event.x - self.x
            deltay = event.y - self.y
            x = self.root.winfo_x() + deltax
            y = self.root.winfo_y() + deltay
            self.root.geometry(f"+{x}+{y}")

    def toggle_maximize(self):
        """切换最大化/还原状态"""
        if not self.is_maximized:
            # 保存当前位置和大小
            self.normal_size = self.root.geometry()
            # 获取屏幕尺寸
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            # 最大化窗口
            self.root.geometry(f"{screen_width}x{screen_height}+0+0")
            self.maximize_button.configure(text="❐")
            self.is_maximized = True
        else:
            # 还原窗口
            if self.normal_size:
                self.root.geometry(self.normal_size)
            self.maximize_button.configure(text="□")
            self.is_maximized = False
        
        # 等待窗口大小更新完成后再设置圆角
        self.root.after(100, self.make_rounded)

    def on_click(self, event):
        """记录鼠标点击位置"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "separator":
            self.is_resizing = True
            self.last_x = event.x
            # 获取当前调整的列
            self.current_column = self.tree.identify_column(event.x)
            # 将当前列转换为索引
            self.current_column_index = int(self.current_column[1]) - 1

    def on_release(self, event):
        """处理鼠标释放事件"""
        if self.is_resizing:
            self.is_resizing = False
            # 更新列宽比例
            total_width = sum(self.tree.column(col, 'width') for col in self.columns.keys())
            for col in self.columns.keys():
                self.column_ratios[col] = self.tree.column(col, 'width') / total_width

    def on_tree_configure(self, event):
        """处理列宽调整"""
        if self.is_resizing:
            # 手动调整列宽的逻辑
            delta = event.x - self.last_x
            if delta != 0:
                columns = list(self.columns.keys())
                current_width = self.tree.column(columns[self.current_column_index], 'width')
                
                # 调整当前列的宽度
                new_width = max(20, current_width + delta)
                self.tree.column(columns[self.current_column_index], width=new_width)
                
                # 调整后续列的宽度
                remaining_width = sum(self.tree.column(col, 'width') for col in columns[self.current_column_index + 1:])
                if remaining_width > 0:
                    # 计算每列应该分配的delta
                    for i in range(self.current_column_index + 1, len(columns)):
                        col = columns[i]
                        col_width = self.tree.column(col, 'width')
                        ratio = col_width / remaining_width
                        col_delta = int(-delta * ratio)  # 反向调整
                        new_col_width = max(20, col_width + col_delta)
                        self.tree.column(col, width=new_col_width)
                
                self.last_x = event.x
        else:
            # 窗口大小变化时的列宽调整
            tree_width = self.tree.winfo_width()
            if tree_width > 1:  # 避免初始化时的无效调整
                # 计算新的列宽
                available_width = tree_width - 20  # 预留滚动条空间
                for col in self.columns.keys():
                    new_width = max(20, int(available_width * self.column_ratios[col]))
                    self.tree.column(col, width=new_width)

    def get_file_icon(self, file_type):
        """根据文件类型返回对应的图标"""
        file_type = file_type.lower()
        if file_type in [".doc", ".docx"]:
            return "📝"
        elif file_type in [".xls", ".xlsx"]:
            return "📊"
        elif file_type in [".ppt", ".pptx"]:
            return "📑"
        return "📄"  # 默认文件图标

    def open_file(self, event):
        """双击打开文件"""
        # 获取点击的项目
        item = self.tree.identify('item', event.x, event.y)
        if not item:
            return
        
        try:
            # 获取文件路径（在最后一列）
            file_path = self.tree.item(item)['values'][-1]
            print(f"正在打开文件: {file_path}")
            
            # 使用系统默认程序打开文件
            if os.path.exists(file_path):
                os.startfile(file_path)
            else:
                messagebox.showerror("错误", f"文件不存在: {file_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件: {str(e)}")

    def on_window_configure(self, event):
        """处理窗口大小变化事件"""
        if event.widget == self.root:
            # 避免过于频繁的更新
            if hasattr(self, '_configure_timer'):
                self.root.after_cancel(self._configure_timer)
            self._configure_timer = self.root.after(100, self.make_rounded)

    def on_closing(self):
        """窗口关闭时保存配置"""
        self.config_manager.update_window_geometry(self.root.geometry())
        self.root.quit()

    def search_directory(self, directory):
        """搜索单个目录"""
        # 准备搜索
        self.searching = True
        self.completed_dirs = 0
        self.total_dirs = 1  # 只搜索一个目录
        
        # 禁用文件类型选择
        self.file_type_combo.configure(state="disabled")
        
        # 显示进度条
        self.progress_var.set("正在搜索...")
        self.progress_bar.pack(side=tk.LEFT, padx=5)
        self.progress_bar.start(10)
        
        # 搜索所有支持的文件类型
        patterns = ["doc", "docx", "xls", "xlsx", "ppt", "pptx"]
        
        # 启动搜索线程
        thread = threading.Thread(
            target=self.file_searcher.search_directory,
            args=(directory, patterns),
            daemon=True
        )
        thread.start()

    def on_search_change(self, *args):
        """处理搜索框内容变化"""
        search_text = self.search_var.get().lower()
        
        # 获取当前显示的所有项目
        all_items = self.tree.get_children("")
        
        if not search_text:
            # 如果搜索框为空，恢复所有项目
            for item in self.all_files:
                try:
                    self.tree.reattach(item, "", "end")
                except:
                    try:
                        self.tree.move(item, "", "end")
                    except:
                        pass
            return
        
        # 遍历所有项目进行搜索
        for item in all_items:
            values = self.tree.item(item)['values']
            if not values:
                continue
            
            # 获取文件名（第二列，因为第一列是图标）
            filename = values[1].lower()
            
            # 如果文件名不包含搜索文本，则隐藏该项目
            if search_text not in filename:
                self.tree.detach(item)

if __name__ == "__main__":
    root = tk.Tk()
          
    # 尝试加载Azure主题
    try:        
        root.tk.call('source', "azure.tcl")
        root.tk.call("set_theme", "light")
    except Exception as e:
        messagebox.showwarning("警告", f"加载主题失败: {e}\n将使用默认主题")
        root_style = ttk.Style(root)
        root_style.theme_use('clam')
    
    # 应用程序实例化
    app = FileOrganizer(root)
    root.mainloop() 