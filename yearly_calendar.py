#!/usr/bin/env python3
"""
年日历生成工具 - 入口脚本

按照业内标准分层架构：
  1. 配置层（config）：集中管理所有常数和配置
  2. 模型层（models）：定义清晰的数据结构
  3. 业务逻辑层（services）：处理核心业务逻辑
  4. 集成层（integration）：与外部框架集成
  5. 应用层（app）：协调各层完成功能
"""

from calendar_app.app.calendar_generator import CalendarGenerator


def main():
    """主函数"""
    # 创建生成器并生成日历
    generator = CalendarGenerator()
    
    # 年份参数：需要生成指定年份时只改这里
    year = 2026  # 例如：2026；为 None 时生成当前年份

    # 生成日历（可指定年份与输出文件）
    # generator.generate(year=2025, output_file="my_calendar.xlsx")
    generator.generate(year=year)


if __name__ == "__main__":
    main()
