"""
数据模型 - 定义清晰的数据结构
"""

from dataclasses import dataclass
from typing import List, Optional


@dataclass
class CellInfo:
    """单个格子信息"""
    
    type: str  # 'month' 或 'day'
    row: int  # Excel行号
    col: int  # Excel列号
    
    # 对于月份格子
    month: Optional[int] = None
    
    # 对于日期格子
    day: Optional[int] = None
    weekday_char: Optional[str] = None
    is_weekend: bool = False
    
    # 格子尺寸（像素）
    width_px: int = 165
    height_px: int = 160


@dataclass
class MonthData:
    """单个月份数据"""
    
    year: int
    month: int
    days: int  # 该月天数
    cells: List[CellInfo]  # 该月所有格子（包括月份标签和日期）


@dataclass
class YearCalendarData:
    """完整年日历数据"""
    
    year: int
    months: List[MonthData]  # 12个月的数据
    
    @property
    def total_cells(self) -> int:
        """计算总格子数"""
        return sum(len(month.cells) for month in self.months)


@dataclass
class ImageGenerationRequest:
    """图像生成请求"""
    
    month: int  # 月份（1-12）
    day: int  # 日期（1-31）
    weekday_char: str  # 周几字符
    cell_width_px: int  # 格子宽度
    cell_height_px: int  # 格子高度
    is_weekend: bool = False  # 是否周末
