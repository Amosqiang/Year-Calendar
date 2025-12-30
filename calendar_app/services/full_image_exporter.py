"""
年日历大图导出服务 - 导出一张完整日历图片
"""

import calendar as calendar_module
from typing import Dict, Tuple, List

from PIL import Image, ImageDraw

from calendar_app.config.calendar_config import CalendarConfig
from calendar_app.models.calendar_models import YearCalendarData, CellInfo, ImageGenerationRequest
from calendar_app.services.cell_image_service import CellImageService


class FullImageExporter:
    """年日历大图导出服务"""

    def __init__(self, config: CalendarConfig = CalendarConfig):
        self.config = config
        self.image_service = CellImageService(config)

    def render_year_image(self, calendar_data: YearCalendarData, output_file: str) -> str:
        """导出一张完整年日历图片"""
        day_map = self._build_day_map(calendar_data)

        col_count = self.config.DAYS_PER_MONTH_MAX
        row_heights = self._get_row_heights()
        total_width = col_count * self.config.get_day_cell_width_px()
        total_height = sum(row_heights)

        img = Image.new("RGB", (total_width, total_height), (255, 255, 255))
        draw = ImageDraw.Draw(img)

        weekend_color = self._hex_to_rgb(self.config.COLOR_WEEKEND_BG)
        weekday_color = self._hex_to_rgb(self.config.COLOR_WEEKDAY_BG)
        border_color = (0, 0, 0)

        y = 0
        for row_index, row_height in enumerate(row_heights, start=1):
            is_month_row = row_index % 2 == 1
            if is_month_row:
                month = (row_index + 1) // 2
                days_in_month = calendar_module.monthrange(calendar_data.year, month)[1]
                for col in range(1, col_count + 1):
                    x = (col - 1) * self.config.get_day_cell_width_px()
                    if col <= days_in_month:
                        cell_info = day_map[(month, col)]
                        bg_color = weekend_color if cell_info.is_weekend else weekday_color
                        draw.rectangle([x, y, x + self.config.get_day_cell_width_px(), y + row_height], fill=bg_color)
                        request = ImageGenerationRequest(
                            month=month,
                            day=col,
                            weekday_char=cell_info.weekday_char,
                            cell_width_px=self.config.get_day_cell_width_px(),
                            cell_height_px=row_height,
                            is_weekend=cell_info.is_weekend,
                        )
                        cell_img = self.image_service.create_image(request)
                        img.paste(cell_img, (x, y), cell_img)
                    else:
                        draw.rectangle([x, y, x + self.config.get_day_cell_width_px(), y + row_height], fill=weekday_color)

                    draw.rectangle(
                        [x, y, x + self.config.get_day_cell_width_px(), y + row_height],
                        outline=border_color,
                        width=1,
                    )
            y += row_height

        img.save(output_file)
        return output_file

    def _build_day_map(self, calendar_data: YearCalendarData) -> Dict[Tuple[int, int], CellInfo]:
        day_map = {}
        for month_data in calendar_data.months:
            for cell in month_data.cells:
                if cell.type == "day":
                    day_map[(cell.month, cell.day)] = cell
        return day_map

    def _get_row_heights(self) -> List[int]:
        row_height = self.config.get_day_cell_height_px()
        spacer_height = int(round(row_height * self.config.MONTH_SPACER_HEIGHT_RATIO))
        heights = []
        for month in range(1, 13):
            heights.append(row_height)
            if month < 12:
                heights.append(spacer_height)
        return heights

    @staticmethod
    def _hex_to_rgb(argb: str) -> Tuple[int, int, int]:
        value = argb[2:] if len(argb) == 8 else argb
        return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))
