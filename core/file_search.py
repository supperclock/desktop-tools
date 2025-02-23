import os
from datetime import datetime
from pathlib import Path
from typing import List, Callable

class FileSearcher:
    def __init__(self, callback: Callable):
        self.callback = callback
        self.searching = False
    
    def search_directory(self, directory: str, patterns: List[str]):
        """搜索目录中的文件"""
        try:
            for pattern in patterns:
                for file in Path(directory).rglob(f"*.{pattern}"):
                    try:
                        file_path = str(file)
                        stats = os.stat(file_path)
                        
                        # 获取文件信息
                        file_info = {
                            "path": file_path,
                            "name": os.path.basename(file_path),
                            "type": os.path.splitext(file_path)[1],
                            "size": self.get_file_size(stats.st_size),
                            "created": datetime.fromtimestamp(
                                stats.st_ctime
                            ).strftime("%Y-%m-%d %H:%M"),
                            "modified": datetime.fromtimestamp(
                                stats.st_mtime
                            ).strftime("%Y-%m-%d %H:%M")
                        }
                        
                        # 发送结果给回调函数
                        self.callback("file", file_info)
                    except Exception as e:
                        print(f"处理文件 {file_path} 时出错: {e}")
                        continue
            
            # 搜索完成
            self.callback("done", directory)
            
        except Exception as e:
            self.callback("error", f"搜索目录 {directory} 时出错: {str(e)}")
            
    @staticmethod
    def get_file_size(size_bytes: int) -> str:
        """将文件大小转换为人类可读格式"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024:
                return f"{size_bytes:.2f}{unit}"
            size_bytes /= 1024
        return f"{size_bytes:.2f}TB" 