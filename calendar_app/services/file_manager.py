"""
文件管理服务 - 处理临时文件和资源清理
"""

import os
import shutil
from typing import Optional

from calendar_app.config.calendar_config import CalendarConfig


class FileManager:
    """文件管理服务"""
    
    def __init__(self, config: CalendarConfig = CalendarConfig):
        self.config = config
        self.temp_dir = config.TEMP_DIR
    
    def create_temp_dir(self) -> str:
        """
        创建临时目录
        
        Returns:
            str: 临时目录路径
        """
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
        os.makedirs(self.temp_dir, exist_ok=True)
        return self.temp_dir
    
    def get_temp_image_path(self, month: int, day: int) -> str:
        """
        获取临时图像文件路径
        
        Args:
            month: 月份
            day: 日期
            
        Returns:
            str: 文件路径
        """
        filename = f"day_{month:02d}_{day:02d}.png"
        return os.path.join(self.temp_dir, filename)
    
    def cleanup_temp_files(self) -> bool:
        """
        清理临时文件
        
        Returns:
            bool: 是否成功清理
        """
        try:
            if os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
            return True
        except Exception as e:
            print(f"清理临时文件失败: {e}")
            return False
    
    def get_output_filename(self, year: int) -> str:
        """
        获取输出文件名
        
        Args:
            year: 年份
            
        Returns:
            str: 输出文件名
        """
        return self.config.OUTPUT_FILENAME_PATTERN.format(year=year)
