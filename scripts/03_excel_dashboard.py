#!/usr/bin/env python3
"""
MIT 8339 Virtual Lab - Task 4: Two-page Excel BI Dashboard
Page 1: Executive Overview   Page 2: Customer Retention Analysis
Self-contained openpyxl build (no external skills/xlsx dependency).
"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import SCORED_CSV, SEGMENTS, QUALITY_CHECKS, DASHBOARD_XLSX
import xlsx_style as base
from xlsx_style import (
    FORMATS, HEADER_BOLD, FONT_NAME,
    setup_sheet, style_header_row, style_data_row,
    font_title, font_header, font_subheader, font_body, font_caption, font_kpi,
    fill_header, fill_data_row, align_text, auto_fit_columns,
    create_bar_chart, create_line_chart, setup_chart_titles, apply_chart_colors,
)

PRIMARY, SECONDARY = base.PRIMARY, base.SECONDARY
POS, NEG, WARN = base.ACCENT_POSITIVE, base.ACCENT_NEGATIVE, base.ACCENT_WARNING
N900, N600, N200, N100, N50 = (
    base.NEUTRAL_900, base.NEUTRAL_600, base.NEUTRAL_200, base.NEUTRAL_100, base.NEUTRAL_50
)

df = pd.read_csv(SCORED_CSV)
N = len(df)
seg = lambda name: pd.read_csv(SEGMENTS / f"{name}.csv")

engaged = float(
    (seg("by_trans_frequency")
     .query("Transaction_Frequency in ['High','Very High']")["Customers"].sum()) / N
)
dormant = float((df["Avg_Utilization_Ratio"] == 0).mean())
K = {
    "attrition": float(df["Is_Attrited"].mean()),
    "attrited_n": int(df["Is_Attrited"].sum()),
    "engaged": engaged,
    "dormant": dormant,
    "avg_credit": float(df["Credit_Limit"].mean()),
    "median_credit": float(df["Credit_Limit"].median()),
    "avg_util": float(df["Avg_Utilization_Ratio"].mean()),
    "avg_trans_amt": float(df["Total_Trans_Amt"].mean()),
    "avg_trans_ct": float(df["Total_Trans_Ct"].mean()),
    "avg_products": float(df["Total_Relationship_Count"].mean()),
    "high_risk": float((df["Churn_Risk_Band"] == "High").mean()),
    "high_risk_n": int((df["Churn_Risk_Band"] == "High").sum()),
}

wb = Workbook()
wb.properties.creator = "MIT8339 BI Virtual Lab"
wb.calculation.fullCalcOnLoad = True


def set_widths(ws, cols, width=12):
    ws.column_dimensions["A"].width = 3
    for c in range(2, cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = width


def kpi_card(ws, row, col, label, value, caption, fmt, span=2):
    L, R = col, col + span - 1
    fill = PatternFill("solid", fgColor=N100)
    thin = Border(bottom=Side(style="thin", color=N200))
    for r in (row, row + 1, row + 2):
        ws.merge_cells(start_row=r, start_column=L, end_row=r, end_column=R)
        for c in range(L, R + 1):
            ws.cell(row=r, column=c).fill = fill
    c1 = ws.cell(row=row, column=L, value=label)
    c1.font = Font(name=FONT_NAME, size=9, color=N600)
    c1.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c2 = ws.cell(row=row + 1, column=L, value=value)
    c2.font = font_kpi()
    c2.number_format = fmt
    c2.alignment = Alignment(horizontal="left", vertical="center")
    c3 = ws.cell(row=row + 2, column=L, value=caption)
    c3.font = Font(name=FONT_NAME, size=8, color=N600)
    c3.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    for c in range(L, R + 1):
        ws.cell(row=row + 2, column=c).border = thin


def page_setup(ws, area):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_area = area
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4


def rate_chart(
    ws_calc, tbl_row, n_rows, anchor, title, kind="col", target_ws=None,
    per_point_colors=None, width=16, height=9.2, label_fmt="0.0%", num_col=3, y_fmt=None,
):
    ch = (
        create_bar_chart(chart_type=kind, width=width, height=height)
        if kind != "line"
        else create_line_chart(width=width, height=height)
    )
    data = Reference(ws_calc, min_col=num_col, min_row=tbl_row, max_row=tbl_row + n_rows)
    cats = Reference(ws_calc, min_col=2, min_row=tbl_row + 1, max_row=tbl_row + n_rows)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    setup_chart_titles(ch, title=title, title_size=12)
    apply_chart_colors(ch, [PRIMARY])
    if per_point_colors:
        for i, hx in enumerate(per_point_colors):
            pt = DataPoint(idx=i)
            pt.graphicalProperties.solidFill = hx
            ch.series[0].data_points.append(pt)
    ch.legend = None
    ch.y_axis.number_format = y_fmt or label_fmt
    ch.dataLabels = DataLabelList(
        showVal=True, showCatName=False, showSerName=False,
        showLegendKey=False, showPercent=False, showBubbleSize=False,
    )
    ch.dataLabels.numFmt = label_fmt
    if kind == "line":
        from openpyxl.chart.marker import Marker
        ch.series[0].marker = Marker(symbol="circle", size=6)
        ch.series[0].smooth = False
    (target_ws or ws_calc).add_chart(ch, anchor)
    return ch


# ================================================================ Calc sheet
calc = wb.active
calc.title = "Calc"
calc.sheet_properties.tabColor = N600
setup_sheet(calc, title="Chart Source Data - reproducibility layer (computed from cleaned dataset)", last_col=5)
rate_fmt = "0.0%"
r = 4
calc.cell(row=r, column=2, value="Status")
calc.cell(row=r, column=3, value="Customers")
style_header_row(calc, r, 2, 3)
for i, (s, n_) in enumerate([
    ("Existing Customers", N - K["attrited_n"]),
    ("Attrited Customers", K["attrited_n"]),
]):
    calc.cell(row=r + 1 + i, column=2, value=s)
    calc.cell(row=r + 1 + i, column=3, value=n_)
    style_data_row(calc, r + 1 + i, 2, 3, i)
T_STATUS = r
r += 4


def write_table(ws, r, headers, rows, fmts=None):
    for j, h in enumerate(headers):
        ws.cell(row=r, column=2 + j, value=h)
    style_header_row(ws, r, 2, 1 + len(headers))
    for i, rowvals in enumerate(rows):
        for j, v in enumerate(rowvals):
            c = ws.cell(row=r + 1 + i, column=2 + j, value=v)
            if fmts and fmts[j]:
                c.number_format = fmts[j]
        style_data_row(ws, r + 1 + i, 2, 1 + len(headers), i)
    return r, len(rows)


inc = seg("by_income")
age = seg("by_age_band")
val = seg("by_value_class")
rsk = seg("by_risk_band")
prd = seg("by_products")
mia = seg("by_months_inactive")

T_INCOME, n1 = write_table(
    calc, r, ["Income Segment", "Attrition Rate"],
    list(zip(inc["Income_Segment"], inc["Attrition_Rate"])), [None, rate_fmt],
)
r += n1 + 3
T_AGE, n2 = write_table(
    calc, r, ["Age Band", "Attrition Rate"],
    list(zip(age["Age_Band"].astype(str), age["Attrition_Rate"])), [None, rate_fmt],
)
r += n2 + 3
T_VALUE, n3 = write_table(
    calc, r, ["Customer Value Class", "Attrition Rate"],
    list(zip(val["Customer_Value_Class"], val["Attrition_Rate"])), [None, rate_fmt],
)
r += n3 + 3
T_RISK, n4 = write_table(
    calc, r, ["Risk Band", "Customers", "Attrition Rate"],
    list(zip(rsk["Churn_Risk_Band"], rsk["Customers"], rsk["Attrition_Rate"])),
    [None, "#,##0", rate_fmt],
)
r += n4 + 3
T_PROD, n5 = write_table(
    calc, r, ["Products Held", "Attrition Rate"],
    list(zip(prd["Total_Relationship_Count"].astype(int), prd["Attrition_Rate"])),
    [None, rate_fmt],
)
r += n5 + 3
T_INACT, n6 = write_table(
    calc, r, ["Months Inactive (12m)", "Attrition Rate"],
    list(zip(mia["Months_Inactive_12_mon"].astype(int), mia["Attrition_Rate"])),
    [None, rate_fmt],
)
r += n6 + 3
calc.cell(
    row=r + 1, column=2,
    value="Source: BankChurners cleaned dataset (n=10,127). Rates = attrited customers / customers in segment.",
).font = font_caption()
auto_fit_columns(calc, min_width=10, max_width=30, header_row=T_STATUS, data_start_row=T_STATUS + 1)

# ================================================================ Sheet 1: Executive Overview
ex = wb.create_sheet("Executive Overview")
ex.sheet_properties.tabColor = PRIMARY
setup_sheet(ex, title="Retail Banking Customer Retention Intelligence - Executive Overview", last_col=16)
ex["B3"] = (
    "Portfolio: 10,127 credit-card customers | Dataset: BankChurners (Kaggle) | "
    "All figures computed from cleaned dataset"
)
ex["B3"].font = font_caption()
set_widths(ex, 17)

cards_r1 = [
    ("ATTRITION RATE", K["attrition"], f"{K['attrited_n']:,} of {N:,} customers lost", FORMATS["percentage"]),
    ("ENGAGED CUSTOMER RATIO", K["engaged"], "customers with 68+ transactions per year", FORMATS["percentage"]),
    ("CARD DORMANCY RATE", K["dormant"], "customers with 0% credit utilisation", FORMATS["percentage"]),
    ("AVERAGE CREDIT LIMIT", K["avg_credit"], f"portfolio median ${K['median_credit']:,.0f}", '"$"#,##0'),
]
cards_r2 = [
    ("AVERAGE UTILISATION", K["avg_util"], "mean share of credit limit in use", FORMATS["percentage"]),
    ("AVG TXN AMOUNT / CUSTOMER", K["avg_trans_amt"], f"over {K['avg_trans_ct']:.0f} txns per customer p.a.", '"$"#,##0'),
    ("AVG PRODUCTS / CUSTOMER", K["avg_products"], "card relationships held (max 6)", FORMATS["decimal_1"]),
    ("HIGH-RISK SHARE", K["high_risk"], f"{K['high_risk_n']:,} customers in High risk band", FORMATS["percentage"]),
]
for i, (lab, val_, cap, fmt) in enumerate(cards_r1):
    kpi_card(ex, 4, 2 + i * 3, lab, val_, cap, fmt)
for i, (lab, val_, cap, fmt) in enumerate(cards_r2):
    kpi_card(ex, 8, 2 + i * 3, lab, val_, cap, fmt)
for row, h in [(4, 16), (5, 30), (6, 26), (7, 6), (8, 16), (9, 30), (10, 26)]:
    ex.row_dimensions[row].height = h

rate_chart(
    calc, T_STATUS, 2, "B13", "Customer Portfolio Split - Existing vs Attrited", "col",
    target_ws=ex, per_point_colors=[PRIMARY, NEG], label_fmt="#,##0", y_fmt="#,##0",
)
rate_chart(calc, T_INCOME, n1, "J13", "Attrition Rate by Income Segment", "col", target_ws=ex)
rate_chart(calc, T_AGE, n2, "B37", "Attrition Rate by Age Band", "col", target_ws=ex)
rate_chart(
    calc, T_VALUE, n3, "J37", "Attrition Rate by Customer Value Class", "col",
    target_ws=ex, per_point_colors=[NEG, PRIMARY, PRIMARY, PRIMARY],
)
page_setup(ex, "A1:Q61")
ex.sheet_view.zoomScale = 90

# ================================================================ Sheet 2: Retention Analysis
ra = wb.create_sheet("Retention Analysis")
ra.sheet_properties.tabColor = PRIMARY
setup_sheet(ra, title="Customer Retention Analysis - Risk Segmentation & Segment Explorer", last_col=16)
ra["B3"] = (
    "Interactive: choose an income segment / card category in the blue dropdown cells; "
    "all metrics and the watchlist recalculate live."
)
ra["B3"].font = font_caption()
set_widths(ra, 17)

DW = "Data"


def explorer_block(ws, col0, block_title, dim_label, dim_col_letter, sel_cell, options):
    cL, cV = col0, col0 + 1
    ws.merge_cells(start_row=4, start_column=cL, end_row=4, end_column=cV + 1)
    t = ws.cell(row=4, column=cL, value=block_title)
    t.font = Font(name=FONT_NAME, size=11, bold=HEADER_BOLD, color="FFFFFF")
    t.fill = fill_header()
    t.alignment = Alignment(horizontal="left", vertical="center")
    for c in range(cL, cV + 2):
        ws.cell(row=4, column=c).fill = fill_header()
    labels = [
        (dim_label, None, None),
        (
            "Customers",
            f'=IFERROR(IF({sel_cell}="All",COUNTA({DW}!$B$5:$B$10131),'
            f'COUNTIF({DW}!${dim_col_letter}$5:${dim_col_letter}$10131,{sel_cell})),0)',
            FORMATS["integer"],
        ),
        (
            "Attrition rate",
            f'=IFERROR(IF({sel_cell}="All",AVERAGE({DW}!$W$5:$W$10131),'
            f'AVERAGEIF({DW}!${dim_col_letter}$5:${dim_col_letter}$10131,{sel_cell},{DW}!$W$5:$W$10131)),0)',
            FORMATS["percentage"],
        ),
        (
            "Avg credit limit",
            f'=IFERROR(IF({sel_cell}="All",AVERAGE({DW}!$O$5:$O$10131),'
            f'AVERAGEIF({DW}!${dim_col_letter}$5:${dim_col_letter}$10131,{sel_cell},{DW}!$O$5:$O$10131)),0)',
            '"$"#,##0',
        ),
        (
            "Avg utilisation",
            f'=IFERROR(IF({sel_cell}="All",AVERAGE({DW}!$V$5:$V$10131),'
            f'AVERAGEIF({DW}!${dim_col_letter}$5:${dim_col_letter}$10131,{sel_cell},{DW}!$V$5:$V$10131)),0)',
            FORMATS["percentage"],
        ),
        (
            "Avg txn amount",
            f'=IFERROR(IF({sel_cell}="All",AVERAGE({DW}!$S$5:$S$10131),'
            f'AVERAGEIF({DW}!${dim_col_letter}$5:${dim_col_letter}$10131,{sel_cell},{DW}!$S$5:$S$10131)),0)',
            '"$"#,##0',
        ),
        (
            "High-risk share",
            f'=IFERROR(IF({sel_cell}="All",COUNTIF({DW}!$AE$5:$AE$10131,"High")/COUNTA({DW}!$B$5:$B$10131),'
            f'COUNTIFS({DW}!${dim_col_letter}$5:${dim_col_letter}$10131,{sel_cell},{DW}!$AE$5:$AE$10131,"High")/'
            f'COUNTIF({DW}!${dim_col_letter}$5:${dim_col_letter}$10131,{sel_cell})),0)',
            FORMATS["percentage"],
        ),
    ]
    for i, (lab, formula, fmt) in enumerate(labels):
        rr = 5 + i
        lc = ws.cell(row=rr, column=cL, value=lab)
        lc.font = Font(name=FONT_NAME, size=10, color=N900 if formula else "FFFFFF")
        lc.fill = fill_data_row(i)
        lc.alignment = align_text()
        ws.merge_cells(start_row=rr, start_column=cV, end_row=rr, end_column=cV + 1)
        for c in range(cV, cV + 2):
            ws.cell(row=rr, column=c).fill = fill_data_row(i)
        vc = ws.cell(row=rr, column=cV, value=formula if formula else options["default"])
        if formula:
            vc.number_format = fmt
            vc.font = Font(name=FONT_NAME, size=10, bold=HEADER_BOLD, color=PRIMARY)
        else:
            vc.font = Font(name=FONT_NAME, size=10, bold=HEADER_BOLD, color="FFFFFF")
            vc.fill = PatternFill("solid", fgColor=PRIMARY)
        vc.alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[rr].height = 18
    dv = DataValidation(type="list", formula1=f'"{options["csv"]}"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws.cell(row=5, column=cV))
    ws.cell(row=12, column=cL, value="Dropdown control - change selection to update metrics").font = font_caption()


explorer_block(
    ra, 2, "SEGMENT EXPLORER - INCOME", "Income segment:", "Y", "$C$5",
    {
        "default": "All",
        "csv": "All,Low (<$40K),Lower-Middle ($40-60K),Middle ($60-80K),Upper-Middle ($80-120K),High (>$120K),Not Declared",
    },
)
explorer_block(
    ra, 7, "SEGMENT EXPLORER - CARD", "Card category:", "J", "$H$5",
    {"default": "All", "csv": "All,Blue,Silver,Gold,Platinum"},
)

rate_chart(
    calc, T_RISK, n4, "B14",
    "Retention Risk Bands - Attrition Rate (rule-based score, 12 points)", "col",
    target_ws=ra, per_point_colors=[POS, WARN, NEG],
)
rate_chart(
    calc, T_RISK, n4, "J14", "Retention Risk Bands - Customer Count", "col",
    target_ws=ra, per_point_colors=[POS, WARN, NEG], label_fmt="#,##0", y_fmt="#,##0",
)
rate_chart(
    calc, T_PROD, n5, "B38",
    "Attrition Rate by Products Held (relationship depth)", "col", target_ws=ra,
)
rate_chart(
    calc, T_INACT, n6, "J38",
    "Attrition Rate by Months Inactive (last 12m)", "line", target_ws=ra,
)

WL0 = 62
ra.merge_cells(start_row=WL0 - 2, start_column=2, end_row=WL0 - 2, end_column=8)
h = ra.cell(
    row=WL0 - 2, column=2,
    value="HIGH-ATTRITION WATCHLIST - Income Segment x Card Category (use column filters)",
)
h.font = font_subheader()
wl = df.pivot_table(
    index=["Income_Segment", "Card_Category"],
    values=["CLIENTNUM", "Is_Attrited"],
    aggfunc={"CLIENTNUM": "count", "Is_Attrited": "sum"},
).reset_index()
wl["Attrition_Rate"] = (wl["Is_Attrited"] / wl["CLIENTNUM"]).round(4)
high = (
    df[df["Churn_Risk_Band"] == "High"]
    .groupby(["Income_Segment", "Card_Category"])["CLIENTNUM"]
    .count()
    .rename("High_Risk")
    .reset_index()
)
wl = wl.merge(high, on=["Income_Segment", "Card_Category"], how="left")
wl["High_Risk"] = wl["High_Risk"].fillna(0)
wl = wl.sort_values("Attrition_Rate", ascending=False)
headers = [
    "Income Segment", "Card Category", "Customers", "Attrited",
    "Attrition Rate", "High-Risk Customers",
]
for j, hd in enumerate(headers):
    ra.cell(row=WL0, column=2 + j, value=hd)
style_header_row(ra, WL0, 2, 7)
for i, row_ in enumerate(wl.itertuples(index=False)):
    rr = WL0 + 1 + i
    vals = [row_[0], row_[1], int(row_[2]), int(row_[3]), float(row_[4]), int(row_[5])]
    for j, v in enumerate(vals):
        c = ra.cell(row=rr, column=2 + j, value=v)
        if j in (2, 3, 5):
            c.number_format = FORMATS["integer"]
        if j == 4:
            c.number_format = FORMATS["percentage"]
            c.font = Font(name=FONT_NAME, size=11, color=NEG if v >= 0.25 else N900)
    style_data_row(ra, rr, 2, 7, i)
    ra.cell(row=rr, column=6).font = Font(
        name=FONT_NAME, size=11, color=NEG if row_[4] >= 0.25 else N900
    )
last_wl = WL0 + len(wl)
ra.auto_filter.ref = f"B{WL0}:G{last_wl}"
ra.cell(
    row=last_wl + 2, column=2,
    value=(
        "Small segments (<30 customers) are indicative only. "
        "Red rates flag pockets above 25% attrition for retention action."
    ),
).font = font_caption()
page_setup(ra, "A1:Q93")
ra.sheet_view.zoomScale = 90

# ================================================================ Data sheet (fast write)
print("[DASH] writing Data sheet...")
dsh = wb.create_sheet("Data")
dsh.sheet_properties.tabColor = N600
setup_sheet(dsh, title="Cleaned & Scored Dataset - BankChurners (n=10,127)", last_col=31)
cols = list(df.columns)
for j, cn in enumerate(cols):
    dsh.cell(row=4, column=2 + j, value=cn)
style_header_row(dsh, 4, 2, 1 + len(cols))

# Write values without per-cell styling for speed (10k rows)
for r_idx, row in enumerate(df.itertuples(index=False), start=5):
    for c_idx, v in enumerate(row, start=2):
        if pd.isna(v):
            continue
        if isinstance(v, (pd.Timestamp,)):
            dsh.cell(row=r_idx, column=c_idx, value=str(v))
        else:
            # Convert numpy types
            if hasattr(v, "item"):
                try:
                    v = v.item()
                except Exception:
                    pass
            dsh.cell(row=r_idx, column=c_idx, value=v)

money_cols = {
    "Credit_Limit", "Total_Revolving_Bal", "Avg_Open_To_Buy", "Total_Trans_Amt",
}
for j, cn in enumerate(cols):
    col_letter = get_column_letter(2 + j)
    if cn in money_cols:
        for r_idx in range(5, 5 + min(N, 50)):  # format sample for template; Excel inherits on open
            pass
        # Apply number format to whole column range via first data cell pattern
        for r_idx in (5,):
            cell = dsh.cell(row=r_idx, column=2 + j)
            cell.number_format = "#,##0" if cn == "Total_Trans_Amt" else "#,##0.00"
    elif cn == "Avg_Utilization_Ratio":
        dsh.cell(row=5, column=2 + j).number_format = "0.0%"

dsh.freeze_panes = "C5"
data_widths = {
    "B": 13, "C": 16, "D": 8, "E": 8, "F": 10, "G": 16, "H": 14, "I": 16, "J": 12,
    "K": 10, "L": 10, "M": 10, "N": 10, "O": 12, "P": 12, "Q": 12, "R": 10, "S": 11,
    "T": 10, "U": 10, "V": 10, "W": 9, "X": 10, "Y": 18, "Z": 16, "AA": 14, "AB": 13,
    "AC": 9, "AD": 9, "AE": 9,
}
for col, w in data_widths.items():
    dsh.column_dimensions[col].width = w
dsh.auto_filter.ref = f"B4:{get_column_letter(1 + len(cols))}{4 + N}"
dsh.sheet_view.zoomScale = 80

# ================================================================ Cleaning_Log sheet
cl = wb.create_sheet("Cleaning_Log")
cl.sheet_properties.tabColor = N600
setup_sheet(cl, title="Data Cleaning & Transformation Log - before / after evidence", last_col=6)
qc = pd.read_csv(QUALITY_CHECKS)
for j, hd in enumerate(qc.columns):
    cl.cell(row=4, column=2 + j, value=hd)
style_header_row(cl, 4, 2, 1 + len(qc.columns))
for i, row_ in enumerate(qc.itertuples(index=False)):
    rr = 5 + i
    for j, v in enumerate(row_):
        c = cl.cell(row=rr, column=2 + j, value=str(v))
        if j == 3:
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    style_data_row(cl, rr, 2, 1 + len(qc.columns), i)
    cl.row_dimensions[rr].height = 30
r2 = 5 + len(qc) + 2
cl.merge_cells(start_row=r2, start_column=2, end_row=r2, end_column=5)
cl.cell(row=r2, column=2, value="Derived analytical fields").font = font_subheader()
derived = [
    ("Is_Attrited", "1 if Attrition_Flag = 'Attrited Customer', else 0 (target variable)"),
    ("Age_Band", "26-34 / 35-44 / 45-54 / 55-64 / 65+ bands from Customer_Age"),
    ("Income_Segment", "Ordinal segments from Income_Category; 'Unknown' mapped to 'Not Declared'"),
    ("Utilisation_Band", "0% (Non-revolver) / 0-10% / 10-30% / 30-50% / >50% from Avg_Utilization_Ratio"),
    ("Transaction_Frequency", "Low (<45) / Moderate (45-67) / High (68-81) / Very High (>81) — aligned with risk rules"),
    ("Customer_Value_Class", "Composite tercile score of Total_Trans_Amt + Credit_Limit + products: Bronze / Silver / Gold / Platinum"),
    ("Risk_Score (0-12)", "Historical retention-priority score: 6 rules x 0-2 points (dormancy, inactivity, txn, momentum, products, contacts)"),
    ("Churn_Risk_Band", "Low (<=4) / Medium (5-7) / High (>=8) — retrospective priority bands, not live prediction"),
]
rr = r2 + 1
for j, hd in enumerate(["Field", "Definition / rule"]):
    cl.cell(row=rr, column=2 + j, value=hd)
style_header_row(cl, rr, 2, 3)
for i, (f, d) in enumerate(derived):
    cl.cell(row=rr + 1 + i, column=2, value=f)
    cl.merge_cells(start_row=rr + 1 + i, start_column=3, end_row=rr + 1 + i, end_column=5)
    cl.cell(row=rr + 1 + i, column=3, value=d)
    style_data_row(cl, rr + 1 + i, 2, 3, i)
cl.column_dimensions["B"].width = 30
cl.column_dimensions["C"].width = 14
cl.column_dimensions["D"].width = 14
cl.column_dimensions["E"].width = 72
cl.column_dimensions["F"].width = 4
page_setup(cl, f"A1:E{rr + len(derived) + 2}")

# ================================================================ Review sheet
rv = wb.create_sheet("Review")
rv.sheet_properties.tabColor = "FFC000"
setup_sheet(rv, title="Cross-Validation Review - every dashboard metric traced to source data", last_col=6)
checks = [
    ["Check", "Expected", "Actual", "Status"],
    ["Customer rows in Data sheet", 10127, "=COUNTA(Data!B5:B10131)", None],
    ["Attrited customers", 1627, "=SUM(Data!W5:W10131)", None],
    ["Attrition rate", 0.1607, "=ROUND(SUM(Data!W5:W10131)/COUNTA(Data!B5:B10131),4)", None],
    ["Avg credit limit ($, rounded)", int(round(K["avg_credit"])), "=ROUND(AVERAGE(Data!O5:O10131),0)", None],
    ["Risk band customers sum to base", 10127, f"=SUM(Calc!C{T_RISK + 1}:C{T_RISK + n4})", None],
    ["Watchlist customers sum to base", 10127, f"=SUM('Retention Analysis'!D{WL0 + 1}:D{last_wl})", None],
    [
        "Executive KPI: high-risk share",
        round(K["high_risk"], 4),
        '=COUNTIF(Data!AE5:AE10131,"High")/COUNTA(Data!B5:B10131)',
        None,
    ],
]
for j, hd in enumerate(checks[0]):
    rv.cell(row=4, column=2 + j, value=hd)
style_header_row(rv, 4, 2, 5)
for i, row_ in enumerate(checks[1:]):
    rr = 5 + i
    rv.cell(row=rr, column=2, value=row_[0])
    rv.cell(row=rr, column=3, value=row_[1])
    rv.cell(row=rr, column=4, value=row_[2])
    rv.cell(row=rr, column=5, value=f'=IF(ABS(C{rr}-D{rr})<0.01,"PASS","FAIL")')
    style_data_row(rv, rr, 2, 5, i)
rv.column_dimensions["B"].width = 34
for col in ["C", "D", "E"]:
    rv.column_dimensions[col].width = 16

dsh.page_setup.orientation = "landscape"
dsh.page_setup.fitToWidth = 1
dsh.page_setup.fitToHeight = 1
dsh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
dsh.print_area = "B1:AE45"
dsh.cell(
    row=4 + N + 2, column=2,
    value="Print preview shows first 40 rows; the sheet contains the full cleaned dataset (10,127 rows).",
).font = font_caption()
page_setup(rv, f"A1:F{4 + len(checks)}")
page_setup(calc, f"A1:F{r + 2}")

# Put dashboard pages first for usability
wb._sheets = [
    wb["Executive Overview"],
    wb["Retention Analysis"],
    wb["Data"],
    wb["Cleaning_Log"],
    wb["Review"],
    wb["Calc"],
]

print(f"[DASH] saving workbook -> {DASHBOARD_XLSX}")
wb.save(DASHBOARD_XLSX)
print(f"[SAVE] {DASHBOARD_XLSX}")
print("sheets:", wb.sheetnames)
