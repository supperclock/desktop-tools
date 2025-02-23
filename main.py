import tkinter as tk
from tkinter import ttk, messagebox
from file_organizer import FileOrganizer

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