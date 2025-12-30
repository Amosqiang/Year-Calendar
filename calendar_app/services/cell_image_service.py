"""
格子图像生成服务 - 使用PIL绘制格子图像
"""

import io
import os

from PIL import Image, ImageDraw, ImageFont

from calendar_app.config.calendar_config import CalendarConfig
from calendar_app.models.calendar_models import ImageGenerationRequest

try:
    import cairosvg
except Exception:
    cairosvg = None


class CellImageService:
    """格子图像生成服务"""
    
    def __init__(self, config: CalendarConfig = CalendarConfig):
        self.config = config
        self.font_path = self._resolve_font_path()
    
    def _load_fonts(self):
        """加载字体"""
        try:
            self.date_font = None  # 延迟加载，因为字号需要根据格子大小计算
            self.weekday_font = None
        except Exception as e:
            print(f"字体加载失败: {e}")

    def _resolve_font_path(self):
        """选择可用字体路径"""
        paths = [self.config.FONT_PATH] + list(self.config.FONT_FALLBACK_PATHS)
        for path in paths:
            if path and os.path.exists(path):
                return path
        return None
    
    def create_image(self, request: ImageGenerationRequest) -> Image.Image:
        """
        为单个格子创建图像
        
        Args:
            request: 图像生成请求
            
        Returns:
            Image.Image: PIL Image对象（RGBA模式）
        """
        if cairosvg:
            try:
                return self._create_svg_image(request)
            except Exception as e:
                print(f"SVG渲染失败，回退到PIL绘制: {e}")
        return self._create_pil_image(request)

    def _create_svg_image(self, request: ImageGenerationRequest) -> Image.Image:
        """使用SVG矢量绘制并渲染为PNG"""
        width = request.cell_width_px
        height = request.cell_height_px

        date_font_size = self.config.get_date_font_size(height)
        weekday_font_size = self.config.get_weekday_font_size(height)
        month_font_size = self.config.get_month_label_font_size(height)

        date_area_width = width * self.config.DATE_AREA_WIDTH_RATIO
        date_area_height = height * self.config.DATE_AREA_HEIGHT_RATIO
        date_center_x = width - date_area_width / 2
        date_center_y = date_area_height / 2
        margin = min(width, height) * self.config.CONTENT_MARGIN_RATIO
        month_baseline = margin

        triangle_width = width * self.config.WEEKDAY_AREA_WIDTH_RATIO * self.config.TRIANGLE_WIDTH_RATIO
        triangle_height = height * self.config.WEEKDAY_AREA_HEIGHT_RATIO * self.config.TRIANGLE_HEIGHT_RATIO
        triangle_margin = min(width, height) * self.config.TRIANGLE_MARGIN_RATIO
        triangle_right = width - triangle_margin
        triangle_bottom = height - triangle_margin
        triangle_left = triangle_right - triangle_width
        triangle_top = triangle_bottom - triangle_height

        date_text = str(request.day)
        weekday_text = request.weekday_char
        month_text = f"{request.month:02d}" if request.day == 1 else ""
        month_en_text = self._get_month_english(request.month) if request.day == 1 else ""

        text_color = self._rgb_color(self.config.COLOR_TEXT_DATE)
        triangle_color = self._rgb_color(self.config.COLOR_TRIANGLE)
        weekday_text_color = self._rgb_color(self.config.COLOR_TEXT_WEEKDAY)
        line_width = self.config.LINE_WIDTH

        font_family = self.config.FONT_FAMILY_NAME or "sans-serif"
        svg = f"""<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
  <g shape-rendering="geometricPrecision">
    <polygon points="{triangle_right},{triangle_bottom} {triangle_left},{triangle_bottom} {triangle_right},{triangle_top}"
      fill="{triangle_color}" />
  </g>
  <text x="{date_center_x}" y="{date_center_y}" text-anchor="middle" dominant-baseline="middle"
    font-family="{font_family}, sans-serif" font-size="{date_font_size}" fill="{text_color}">{date_text}</text>
  <text x="{margin}" y="{month_baseline}" text-anchor="start" dominant-baseline="hanging"
    font-family="{font_family}, sans-serif" font-size="{month_font_size}" font-weight="900" fill="{text_color}">{month_text}</text>
  <text x="{margin}" y="{month_baseline + month_font_size * 1.05}" text-anchor="start" dominant-baseline="hanging"
    font-family="{font_family}, sans-serif" font-size="{month_font_size * self.config.MONTH_LABEL_ENGLISH_SIZE_RATIO}" font-weight="900" fill="{text_color}">{month_en_text}</text>
  <text x="{triangle_right - triangle_height * 0.38}" y="{triangle_bottom - triangle_height * 0.38}" text-anchor="middle" dominant-baseline="middle"
    font-family="{font_family}, sans-serif" font-size="{weekday_font_size}" fill="{weekday_text_color}">{weekday_text}</text>
</svg>"""

        png_bytes = cairosvg.svg2png(
            bytestring=svg.encode("utf-8"),
            output_width=width,
            output_height=height,
        )
        img = Image.open(io.BytesIO(png_bytes))
        return img.convert("RGBA")

    def _create_pil_image(self, request: ImageGenerationRequest) -> Image.Image:
        """PIL绘制回退方案"""
        scale = max(1, int(self.config.RENDER_SCALE))
        if scale == 1:
            return self._render_pil(request, scale)

        high_res_request = ImageGenerationRequest(
            month=request.month,
            day=request.day,
            weekday_char=request.weekday_char,
            cell_width_px=request.cell_width_px * scale,
            cell_height_px=request.cell_height_px * scale,
            is_weekend=request.is_weekend,
        )
        return self._render_pil(high_res_request, scale)

    def _render_pil(self, request: ImageGenerationRequest, scale: int) -> Image.Image:
        img = Image.new(
            'RGBA',
            (request.cell_width_px, request.cell_height_px),
            (255, 255, 255, 0),
        )
        draw = ImageDraw.Draw(img)

        line_width = max(1, int(self.config.LINE_WIDTH * scale))
        date_font = self._fit_font_for_date(request)
        self._draw_date(draw, request, date_font)
        if request.day == 1:
            month_font = self._get_month_label_font(request)
            self._draw_month_label(draw, request, month_font)
        self._draw_weekday_with_triangle(draw, request, line_width)
        return img

    @staticmethod
    def _rgb_color(color):
        """从RGBA/RGB元组转换为SVG颜色"""
        return f"rgb({color[0]}, {color[1]}, {color[2]})"

    def _get_triangle_bounds(self, request: ImageGenerationRequest, margin: int):
        triangle_width = int(request.cell_width_px * self.config.WEEKDAY_AREA_WIDTH_RATIO * self.config.TRIANGLE_WIDTH_RATIO)
        triangle_height = int(request.cell_height_px * self.config.WEEKDAY_AREA_HEIGHT_RATIO * self.config.TRIANGLE_HEIGHT_RATIO)

        triangle_right = request.cell_width_px - margin
        triangle_left = triangle_right - triangle_width
        triangle_bottom = request.cell_height_px - margin
        triangle_top = triangle_bottom - triangle_height
        triangle_center_x = (triangle_left + triangle_right) // 2
        return triangle_left, triangle_right, triangle_top, triangle_bottom, triangle_center_x

    def _fit_font_for_date(self, request: ImageGenerationRequest) -> ImageFont.FreeTypeFont:
        date_text = str(request.day)
        margin = int(min(request.cell_width_px, request.cell_height_px) * self.config.CONTENT_MARGIN_RATIO)

        date_area_width = int(request.cell_width_px * self.config.DATE_AREA_WIDTH_RATIO) - margin
        date_area_height = int(request.cell_height_px * self.config.DATE_AREA_HEIGHT_RATIO) - margin

        max_width = max(1, date_area_width)
        max_height = max(1, date_area_height)

        cell_area = request.cell_width_px * request.cell_height_px
        target_area = min(cell_area * self.config.DATE_TEXT_AREA_RATIO, max_width * max_height)

        return self._fit_font(date_text, max_width, max_height, target_area)

    def _fit_font_for_weekday(self, request: ImageGenerationRequest,
                              triangle_left: int, triangle_right: int,
                              triangle_top: int, triangle_bottom: int) -> ImageFont.FreeTypeFont:
        text = request.weekday_char
        inner_margin = max(1, int(min(request.cell_width_px, request.cell_height_px) * self.config.CONTENT_MARGIN_RATIO * 0.6))
        triangle_width = max(1, (triangle_right - triangle_left))
        triangle_height = max(1, (triangle_bottom - triangle_top))

        max_width = int(triangle_width * self.config.WEEKDAY_TRIANGLE_TEXT_WIDTH_RATIO)
        max_height = int(triangle_height * self.config.WEEKDAY_TRIANGLE_TEXT_HEIGHT_RATIO)
        max_width = max(1, max_width - inner_margin * 2)
        max_height = max(1, max_height - inner_margin * 2)

        cell_area = request.cell_width_px * request.cell_height_px
        max_area = max_width * max_height
        target_area = min(cell_area * self.config.WEEKDAY_TEXT_AREA_RATIO, max_area)

        return self._fit_font(text, max_width, max_height, target_area)

    def _fit_font(self, text: str, max_width: int, max_height: int, target_area: float) -> ImageFont.FreeTypeFont:
        if not self.font_path:
            return ImageFont.load_default()

        size_low = 4
        size_high = max(8, max_height)
        best_size = size_low

        probe_img = Image.new("RGB", (1, 1))
        probe_draw = ImageDraw.Draw(probe_img)

        while size_low <= size_high:
            size_mid = (size_low + size_high) // 2
            font = ImageFont.truetype(self.font_path, size_mid)
            bbox = probe_draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            area = text_width * text_height

            fits = text_width <= max_width and text_height <= max_height and area <= target_area
            if fits:
                best_size = size_mid
                size_low = size_mid + 1
            else:
                size_high = size_mid - 1

        return ImageFont.truetype(self.font_path, best_size)

    def _get_month_label_font(self, request: ImageGenerationRequest) -> ImageFont.FreeTypeFont:
        size = self.config.get_month_label_font_size(request.cell_height_px)
        if not self.font_path:
            return ImageFont.load_default()
        return ImageFont.truetype(self.font_path, size)

    def _get_month_english_font(self, request: ImageGenerationRequest) -> ImageFont.FreeTypeFont:
        base_size = self.config.get_month_label_font_size(request.cell_height_px)
        size = max(6, int(base_size * self.config.MONTH_LABEL_ENGLISH_SIZE_RATIO))
        if not self.font_path:
            return ImageFont.load_default()
        return ImageFont.truetype(self.font_path, size)

    def _get_month_english(self, month: int) -> str:
        if 1 <= month <= 12:
            return self.config.MONTH_ENGLISH_NAMES[month - 1]
        return ""
    
    def _draw_date(self, draw: ImageDraw.ImageDraw,
                   request: ImageGenerationRequest,
                   font: ImageFont.FreeTypeFont):
        """
        绘制日期（右上角）
        
        Args:
            draw: ImageDraw对象
            request: 请求对象
            font: 字体对象
        """
        # 格式化日期
        date_text = str(request.day)

        # 计算文本尺寸
        date_bbox = draw.textbbox((0, 0), date_text, font=font)
        date_text_width = date_bbox[2] - date_bbox[0]
        date_text_height = date_bbox[3] - date_bbox[1]

        margin = int(min(request.cell_width_px, request.cell_height_px) * self.config.CONTENT_MARGIN_RATIO)
        date_x = request.cell_width_px - margin - date_text_width
        date_y = margin

        # 绘制
        draw.text((date_x, date_y), date_text, 
                 fill=self.config.COLOR_TEXT_DATE, font=font)

    def _draw_month_label(self, draw: ImageDraw.ImageDraw,
                          request: ImageGenerationRequest,
                          font: ImageFont.FreeTypeFont):
        month_text = f"{request.month:02d}"
        margin = int(min(request.cell_width_px, request.cell_height_px) * self.config.CONTENT_MARGIN_RATIO)
        english_text = self._get_month_english(request.month)
        stroke_width = max(1, int(self.config.MONTH_LABEL_STROKE_WIDTH))
        draw.text(
            (margin, margin),
            month_text,
            fill=self.config.COLOR_TEXT_DATE,
            font=font,
            stroke_width=stroke_width,
            stroke_fill=self.config.COLOR_TEXT_DATE,
        )
        english_font = self._get_month_english_font(request)
        english_y = margin + int(font.size * 1.05)
        draw.text(
            (margin, english_y),
            english_text,
            fill=self.config.COLOR_TEXT_DATE,
            font=english_font,
            stroke_width=max(1, int(self.config.MONTH_LABEL_ENGLISH_STROKE_WIDTH)),
            stroke_fill=self.config.COLOR_TEXT_DATE,
        )
    
    def _draw_weekday_with_triangle(self, draw: ImageDraw.ImageDraw,
                                    request: ImageGenerationRequest,
                                    line_width: int):
        """
        绘制周几和三角形（右下角）
        
        Args:
            draw: ImageDraw对象
            request: 请求对象
            font: 字体对象
        """
        triangle_margin = int(min(request.cell_width_px, request.cell_height_px) * self.config.TRIANGLE_MARGIN_RATIO)
        triangle_left, triangle_right, triangle_top, triangle_bottom, _triangle_center_x = self._get_triangle_bounds(
            request, triangle_margin
        )
        triangle_height = triangle_bottom - triangle_top

        # 绘制直角三角形（直角在右下角）
        draw.polygon(
            [
                (triangle_right, triangle_bottom),
                (triangle_left, triangle_bottom),
                (triangle_right, triangle_top),
            ],
            fill=self.config.COLOR_TRIANGLE,
        )
        
        weekday_font = self._fit_font_for_weekday(request, triangle_left, triangle_right, triangle_top, triangle_bottom)
        weekday_bbox = draw.textbbox((0, 0), request.weekday_char, font=weekday_font)
        weekday_text_width = weekday_bbox[2] - weekday_bbox[0]
        weekday_text_height = weekday_bbox[3] - weekday_bbox[1]
        weekday_x = triangle_right - int(triangle_height * 0.38) - weekday_text_width // 2
        weekday_y = triangle_bottom - int(triangle_height * 0.38) - weekday_text_height // 2

        draw.text((weekday_x, weekday_y), request.weekday_char,
                 fill=self.config.COLOR_TEXT_WEEKDAY, font=weekday_font)
