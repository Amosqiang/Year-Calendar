"""
日历生成器 - 主应用类，协调各个服务完成日历生成
"""

from datetime import datetime

from calendar_app.config.calendar_config import CalendarConfig
from calendar_app.services.calendar_service import CalendarService
from calendar_app.services.file_manager import FileManager
from calendar_app.integration.excel_builder import ExcelBuilder


class CalendarGenerator:
    """日历生成器 - 协调各服务完成日历生成"""
    
    def __init__(self, config: CalendarConfig = None):
        """
        初始化生成器
        
        Args:
            config: 配置对象（可选，默认使用标准配置）
        """
        self.config = config or CalendarConfig
        self.calendar_service = CalendarService(self.config)
        self.file_manager = FileManager(self.config)
        self.excel_builder = ExcelBuilder(self.config)
    
    def generate(self, year: int = None, output_file: str = None) -> bool:
        """
        生成年日历
        
        Args:
            year: 年份（默认当前年份）
            output_file: 输出文件名（默认为yearly_calendar_{year}.xlsx）
            
        Returns:
            bool: 是否成功生成
        """
        try:
            # 处理默认参数
            if year is None:
                year = datetime.now().year
            
            if output_file is None:
                output_file = self.file_manager.get_output_filename(year)
            
            print(f"开始生成年日历...")
            print(f"  年份: {year}")
            print(f"  输出文件: {output_file}")
            
            # 第1阶段：数据准备
            print(f"\n[1/4] 生成日历数据...")
            calendar_data = self.calendar_service.generate_year_data(year)
            print(f"  ✓ 总格子数: {calendar_data.total_cells}")
            
            # 第2阶段：创建Excel工作簿
            print(f"\n[2/4] 创建Excel工作簿...")
            self.excel_builder.create_workbook()
            self.excel_builder.setup_layout()
            print(f"  ✓ 工作簿创建成功")
            
            # 第3阶段：填充数据和图像
            print(f"\n[3/4] 生成格子图像并填充数据...")
            self.excel_builder.fill_cells(calendar_data)
            print(f"  ✓ 数据填充完成")
            
            # 第4阶段：保存文件
            print(f"\n[4/4] 保存Excel文件...")
            if not self.excel_builder.save(output_file):
                return False
            print(f"  ✓ 文件保存成功")
            
            # 清理临时文件
            print(f"\n清理临时文件...")
            self.file_manager.cleanup_temp_files()
            print(f"  ✓ 清理完成")
            
            # 成功完成
            print(f"\n{'='*50}")
            print(f"✓ 年日历已成功生成")
            print(f"✓ 文件: {output_file}")
            print(f"✓ 年份: {year}")
            print(f"{'='*50}")
            
            return True
        
        except Exception as e:
            print(f"\n✗ 生成日历出错: {e}")
            # 清理临时文件
            self.file_manager.cleanup_temp_files()
            return False
