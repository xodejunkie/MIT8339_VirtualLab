"""Excel styling helpers (colours, fonts, borders, chart setup) used by 03_excel_dashboard.py."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter

# Blue palette
PRIMARY = "1B2A4A"
SECONDARY = "2E5A88"
ACCENT_POSITIVE = "2E8B57"
ACCENT_NEGATIVE = "C0392B"
ACCENT_WARNING = "F2994A"
NEUTRAL_900 = "1A1A1A"
NEUTRAL_600 = "666666"
NEUTRAL_200 = "D9D9D9"
NEUTRAL_100 = "F0F2F5"
NEUTRAL_50 = "F7F8FA"

FONT_NAME = "Calibri"
HEADER_BOLD = True

FORMATS = {
    "percentage": "0.0%",
    "integer": "#,##0",
    "decimal_1": "0.0",
    "money": '"$"#,##0',
}

_thin = Side(style="thin", color=NEUTRAL_200)


def font_title():
    return Font(name=FONT_NAME, size=16, bold=True, color=PRIMARY)


def font_header():
    return Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")


def font_subheader():
    return Font(name=FONT_NAME, size=12, bold=True, color=PRIMARY)


def font_body():
    return Font(name=FONT_NAME, size=10, color=NEUTRAL_900)


def font_caption():
    return Font(name=FONT_NAME, size=9, italic=True, color=NEUTRAL_600)


def font_kpi():
    return Font(name=FONT_NAME, size=18, bold=True, color=PRIMARY)


def font_kpi_label():
    return Font(name=FONT_NAME, size=9, color=NEUTRAL_600)


def fill_header():
    return PatternFill("solid", fgColor=PRIMARY)


def fill_total():
    return PatternFill("solid", fgColor=NEUTRAL_100)


def fill_data_row(i: int):
    return PatternFill("solid", fgColor=NEUTRAL_50 if i % 2 else "FFFFFF")


def border_header():
    return Border(bottom=Side(style="medium", color=PRIMARY))


def align_text():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


def align_number():
    return Alignment(horizontal="right", vertical="center")


def setup_sheet(ws, title: str, last_col: int = 10):
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=min(last_col, 12))
    cell = ws.cell(row=1, column=2, value=title)
    cell.font = font_title()
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3


def style_header_row(ws, row: int, start_col: int, end_col: int):
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = font_header()
        cell.fill = fill_header()
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_header()
    ws.row_dimensions[row].height = 22


def style_data_row(ws, row: int, start_col: int, end_col: int, i: int):
    fill = fill_data_row(i)
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font_body()
        cell.alignment = Alignment(horizontal="left", vertical="center")


def style_total_row(ws, row: int, start_col: int, end_col: int):
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT_NAME, size=10, bold=True, color=NEUTRAL_900)
        cell.fill = fill_total()


def auto_fit_columns(ws, min_width=10, max_width=30, header_row=1, data_start_row=2):
    for col_cells in ws.iter_cols(min_row=header_row, max_row=min(data_start_row + 40, ws.max_row)):
        letter = get_column_letter(col_cells[0].column)
        if letter == "A":
            continue
        lengths = [len(str(c.value)) if c.value is not None else 0 for c in col_cells]
        width = min(max(max(lengths, default=min_width) + 2, min_width), max_width)
        ws.column_dimensions[letter].width = width


def create_bar_chart(chart_type="col", width=16, height=9.2):
    ch = BarChart()
    ch.type = "col" if chart_type == "col" else chart_type
    ch.style = 10
    ch.width = width
    ch.height = height
    ch.y_axis.majorGridlines = None
    return ch


def create_line_chart(width=16, height=9.2):
    ch = LineChart()
    ch.style = 10
    ch.width = width
    ch.height = height
    return ch


def setup_chart_titles(ch, title="", title_size=12):
    ch.title = title
    if ch.title:
        ch.title.txPr = None


def apply_chart_colors(ch, colors):
    for i, series in enumerate(ch.series):
        hx = colors[i % len(colors)]
        series.graphicalProperties.solidFill = hx


def make_chart_title(text):
    return text
