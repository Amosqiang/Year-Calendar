"""
日历数据服务 - 处理日历相关的业务逻辑
"""

import calendar as calendar_module
from typing import List

from calendar_app.config.calendar_config import CalendarConfig
from calendar_app.models.calendar_models import CellInfo, MonthData, YearCalendarData


class CalendarService:
    """日历数据生成服务"""
    
    def __init__(self, config: CalendarConfig = CalendarConfig):
        self.config = config
    
    def generate_year_data(self, year: int) -> YearCalendarData:
        """
        生成完整年份的日历数据
        
        Args:
            year: 年份
            
        Returns:
            YearCalendarData: 包含12个月数据的日历对象
        """
        months = []
        
        for month in range(1, 13):
            month_data = self._generate_month_data(year, month)
            months.append(month_data)
        
        return YearCalendarData(year=year, months=months)
    
    def _generate_month_data(self, year: int, month: int) -> MonthData:
        """
        生成单个月份的数据
        
        Args:
            year: 年份
            month: 月份（1-12）
            
        Returns:
            MonthData: 包含该月所有格子信息
        """
        # 获取该月的天数
        num_days = calendar_module.monthrange(year, month)[1]
        
        cells = []

        day_cell_width_px = self.config.get_day_cell_width_px()
        day_cell_height_px = self.config.get_day_cell_height_px()

        # 日期格子
        for day in range(1, num_days + 1):
            # 计算周几 (0=周一, 6=周日)
            weekday_index = calendar_module.weekday(year, month, day)
            weekday_char = self.config.WEEKDAY_NAMES[weekday_index]
            is_weekend = weekday_index >= 5
            
            row = month + (month - 1)
            day_cell = CellInfo(
                type='day',
                row=row,
                col=day,
                day=day,
                month=month,
                weekday_char=weekday_char,
                is_weekend=is_weekend,
                width_px=day_cell_width_px,
                height_px=day_cell_height_px,
            )
            cells.append(day_cell)
        
        return MonthData(
            year=year,
            month=month,
            days=num_days,
            cells=cells
        )
