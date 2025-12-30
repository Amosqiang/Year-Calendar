# Yearly Calendar Generator

Generate a yearly calendar in Excel with custom day cells, weekday triangles, and month labels.

## Features
- Input a year and export an Excel file.
- 12 month rows with spacer rows between months.
- Day cells include date (top-right), weekday (bottom-right triangle), and month label on day 1.
- Weekend/weekday background colors and borders.

## Requirements
- Python 3.9+
- Dependencies:
  - openpyxl
  - pillow
  - cairosvg (optional, for SVG rendering; falls back to PIL if missing)

Install (recommended):
```bash
pip install -r requirements.txt
```

Install (manual):
```bash
pip install openpyxl pillow cairosvg
```

## Usage
Edit the year in `yearly_calendar.py`:
```python
year = 2026  # set your year, or None for current year
```

Run:
```bash
python yearly_calendar.py
```

Output:
- `yearly_calendar_{year}.xlsx` in the project root.

## Example
```bash
python yearly_calendar.py
```

You can change the year in `yearly_calendar.py`:
```python
year = 2026  # set your year, or None for current year
```

## Screenshot
Add a screenshot of the exported Excel if you want to showcase the output.

## Configuration
Most layout and style settings live in:
`calendar_app/config/calendar_config.py`

Common options:
- `DATE_COLUMN_WIDTH`, `ROW_HEIGHT`
- `MONTH_SPACER_HEIGHT_RATIO`
- `COLOR_WEEKDAY_BG`, `COLOR_WEEKEND_BG`
- `MONTH_LABEL_FONT_SIZE_RATIO`
- `WEEKDAY_TRIANGLE_TEXT_WIDTH_RATIO`

## Project Structure
```
calendar_app/
  app/            # application layer
  config/         # configuration
  integration/    # Excel integration (openpyxl)
  models/         # data models
  services/       # calendar logic and cell rendering
yearly_calendar.py
```

## Notes
- The font path is set to macOS system fonts by default. Update `FONT_PATH` if needed.
- SVG rendering is used when `cairosvg` is available; otherwise PIL rendering is used.
