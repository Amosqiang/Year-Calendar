"""
配置管理 - 集中管理所有常数和配置
"""


class CalendarConfig:
    """日历配置类"""
    
    # ===== Excel配置 =====
    PAPER_SIZE = "A3"  # 纸张大小
    ORIENTATION = "landscape"  # 横向
    
    # 页边距（厘米）
    MARGIN_LEFT = 0.4
    MARGIN_RIGHT = 0.4
    MARGIN_TOP = 0.4
    MARGIN_BOTTOM = 0.4
    
    # ===== 列行配置 =====
    DATE_COLUMN_WIDTH = 22  # 日期列宽（Excel单位）
    ROW_HEIGHT = 120  # 行高（Excel单位）
    DAYS_PER_MONTH_MAX = 31  # 每月最大天数
    MONTH_SPACER_HEIGHT_RATIO = 0.25  # 月份间隔占行高比例
    
    # ===== 像素尺寸 =====
    CELL_WIDTH_PX = 165  # 日期格子宽度（像素）
    CELL_HEIGHT_PX = 160  # 日期格子高度（像素）
    
    # ===== 颜色配置（ARGB格式） =====
    COLOR_MONTH_BG = "FF2E5AA2"  # 月份背景 - 深蓝
    COLOR_WEEKEND_BG = "FFE8F4FF"  # 周末背景 - 浅蓝
    COLOR_WEEKDAY_BG = "FFFFFFFF"  # 工作日背景 - 白色
    COLOR_TEXT_DATE = (0, 0, 0, 255)  # 日期文字 - 黑色 (RGBA)
    COLOR_TRIANGLE = (0, 0, 0, 255)  # 三角形 - 黑色 (RGBA)
    COLOR_TEXT_WEEKDAY = (255, 255, 255, 255)  # 周几文字 - 白色 (RGBA)
    
    # ===== 字体配置 =====
    FONT_PATH = "/System/Library/Fonts/Hiragino Sans GB.ttc"  # 字体路径
    FONT_FAMILY_NAME = "Hiragino Sans GB"  # 字体族名
    FONT_FALLBACK_PATHS = [
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
    ]
    FONT_MONTH_SIZE = 12  # 月份字号
    FONT_DATE_SIZE = None  # 日期字号（按比例计算）
    FONT_WEEKDAY_SIZE = None  # 周几字号（按比例计算）
    MONTH_LABEL_FONT_SIZE_RATIO = 0.208  # 月份标签字号占格子高度的比例
    MONTH_LABEL_STROKE_WIDTH = 2  # 月份标签描边，模拟超粗体
    MONTH_LABEL_ENGLISH_SIZE_RATIO = 0.385  # 英文月份字号比例（相对于中文月份字号）
    MONTH_LABEL_ENGLISH_STROKE_WIDTH = 1  # 英文月份描边，模拟特粗
    
    # ===== 排版配置 =====
    DATE_AREA_HEIGHT_RATIO = 0.45  # 日期区域占格子高度的比例
    DATE_AREA_WIDTH_RATIO = 0.45  # 日期区域占格子宽度的比例
    WEEKDAY_AREA_HEIGHT_RATIO = 0.35  # 周几区域占格子高度的比例
    WEEKDAY_AREA_WIDTH_RATIO = 0.45  # 周几区域占格子宽度的比例
    CONTENT_MARGIN_RATIO = 0.04  # 内容与边缘的间距比例
    TRIANGLE_MARGIN_RATIO = 0.0  # 三角形与边缘的间距比例
    
    # ===== 绘图配置 =====
    TRIANGLE_WIDTH_RATIO = 1.0  # 三角形宽度占周几区域宽度的比例
    TRIANGLE_HEIGHT_RATIO = 1.0  # 三角形高度占周几区域高度的比例
    LINE_WIDTH = 2  # 三角形线条宽度（像素）
    
    # ===== 字体大小比例 =====
    DATE_FONT_SIZE_RATIO = 1.0  # 日期字号占日期区域高度的比例
    WEEKDAY_FONT_SIZE_RATIO = 0.95  # 周几字号占周几区域高度的比例
    DATE_TEXT_AREA_RATIO = 0.025  # 日期文字占整张图片面积比例
    WEEKDAY_TEXT_AREA_RATIO = 0.01875  # 周几文字占整张图片面积比例
    WEEKDAY_TRIANGLE_TEXT_WIDTH_RATIO = 0.55  # 周几文字占三角形宽度比例
    WEEKDAY_TRIANGLE_TEXT_HEIGHT_RATIO = 0.55  # 周几文字占三角形高度比例

    # ===== 渲染配置 =====
    RENDER_SCALE = 4  # 先高分辨率绘制，提升清晰度
    
    # ===== 文件配置 =====
    TEMP_DIR = "./temp_calendar_images"  # 临时目录
    OUTPUT_FILENAME_PATTERN = "yearly_calendar_{year}.xlsx"  # 输出文件名模式
    OUTPUT_IMAGE_PATTERN = "yearly_calendar_{year}.png"  # 大图输出文件名模式

    
    # ===== 周几名称 =====
    WEEKDAY_NAMES = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']  # 中文周几
    MONTH_ENGLISH_NAMES = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]
    
    @classmethod
    def get_date_font_size(cls, cell_height_px):
        """根据格子高度计算日期字号"""
        date_area_height = int(cell_height_px * cls.DATE_AREA_HEIGHT_RATIO)
        return int(date_area_height * cls.DATE_FONT_SIZE_RATIO)
    
    @classmethod
    def get_weekday_font_size(cls, cell_height_px):
        """根据格子高度计算周几字号"""
        weekday_area_height = int(cell_height_px * cls.WEEKDAY_AREA_HEIGHT_RATIO)
        return int(weekday_area_height * cls.WEEKDAY_FONT_SIZE_RATIO)

    @classmethod
    def get_month_label_font_size(cls, cell_height_px):
        """根据格子高度计算月份标签字号"""
        return int(cell_height_px * cls.MONTH_LABEL_FONT_SIZE_RATIO)

    @staticmethod
    def excel_column_width_to_pixels(width: float) -> int:
        """将Excel列宽转换为像素（近似值）"""
        if width <= 0:
            return 0
        if width <= 1:
            return int(round(width * 12))
        return int(round(width * 7 + 5))

    @staticmethod
    def points_to_pixels(points: float) -> int:
        """将Excel行高（点）转换为像素（96 DPI）"""
        return int(round(points * 96 / 72))

    @staticmethod
    def pixels_to_points(pixels: int) -> float:
        """像素转点（96 DPI）"""
        return pixels * 72 / 96

    @staticmethod
    def mm_to_points(mm: float) -> float:
        """毫米转点（1英寸=25.4mm, 1点=1/72英寸）"""
        return mm * 72 / 25.4

    @staticmethod
    def mm_to_pixels(mm: float) -> int:
        """毫米转像素（96 DPI）"""
        return int(round(mm * 96 / 25.4))

    @staticmethod
    def pixels_to_excel_column_width(pixels: int) -> float:
        """将像素转换为Excel列宽（近似值）"""
        if pixels <= 12:
            return round(pixels / 12, 2)
        return round((pixels - 5) / 7, 2)

    @classmethod
    def get_date_column_width(cls) -> float:
        """获取日期列宽（Excel单位）"""
        return cls.DATE_COLUMN_WIDTH

    @classmethod
    def get_row_height_points(cls) -> float:
        """获取行高（点）"""
        return cls.ROW_HEIGHT

    @classmethod
    def get_day_cell_width_px(cls) -> int:
        """根据Excel列宽计算日期格子宽度（像素）"""
        return cls.excel_column_width_to_pixels(cls.get_date_column_width())

    @classmethod
    def get_day_cell_height_px(cls) -> int:
        """根据Excel行高计算日期格子高度（像素）"""
        return cls.points_to_pixels(cls.get_row_height_points())

    @classmethod
    def get_month_cell_width_px(cls) -> int:
        """根据月份列宽比例计算月份格子宽度（像素）"""
        return cls.get_day_cell_width_px()
