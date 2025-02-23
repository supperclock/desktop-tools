import json
import os
from pathlib import Path

class ConfigManager:
    def __init__(self):
        # 确保配置目录存在
        self.config_dir = Path.home() / '.file_organizer'
        self.config_dir.mkdir(exist_ok=True)
        
        # 配置文件路径
        self.config_file = self.config_dir / 'config.json'
        
        # 默认配置
        self.default_config = {
            'directories': [],
            'last_window_size': '1000x700',
            'last_window_position': '+100+100'
        }
        
        # 加载配置
        self.config = self.load_config()
    
    def load_config(self):
        """加载配置文件"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return self.default_config.copy()
        except Exception as e:
            print(f"加载配置文件失败: {e}")
            return self.default_config.copy()
    
    def save_config(self):
        """保存配置文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"保存配置文件失败: {e}")
    
    def get_directories(self):
        """获取保存的目录列表"""
        return self.config.get('directories', [])
    
    def add_directory(self, directory):
        """添加目录到配置"""
        if directory not in self.config['directories']:
            self.config['directories'].append(directory)
            self.save_config()
    
    def remove_directory(self, directory):
        """从配置中移除目录"""
        if directory in self.config['directories']:
            self.config['directories'].remove(directory)
            self.save_config()
    
    def update_window_geometry(self, geometry):
        """更新窗口位置和大小"""
        try:
            # 处理几种可能的格式
            if '+' in geometry:
                size, *position = geometry.split('+')
                position = '+' + '+'.join(position)
            else:
                size = geometry
                position = '+100+100'  # 默认位置
            
            self.config['last_window_size'] = size
            self.config['last_window_position'] = position
            self.save_config()
        except Exception as e:
            print(f"更新窗口位置失败: {e}")
            # 使用默认值
            self.config['last_window_size'] = '1000x700'
            self.config['last_window_position'] = '+100+100' 