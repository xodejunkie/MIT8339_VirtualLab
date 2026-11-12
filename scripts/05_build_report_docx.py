#!/usr/bin/env python3
"""
Build the MIT 8339 analytical report DOCX with:
- fixed relative-path codebase references
- dataset description table
- data preparation & cleaning decisions table
- regenerated figures and dashboard screenshots
"""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt, RGBColor, Cm, Twips

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import ROOT, FIGS, DATA, OUTPUTS, SCRIPTS

OUT = ROOT / "MIT8339_BI_VirtualLab_Report_T.AKINDEKO1787@MIVA.EDU.NG.docx"
REPO_URL = "https://github.com/xodejunkie/MIT8339_VirtualLab"


def set_run_font(run, size=11, bold=False, italic=False, color=None, name="Calibri"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color


def add_para(doc, text, size=11, bold=False, italic=False, space_after=8, space_before=0, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def add_heading_custom(doc, text, level=1):
    # Use built-in heading styles for outline, then tweak font
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        set_run_font(run, size=14 if level == 1 else 12, bold=True, color=RGBColor(0x1B, 0x2A, 0x4A))
    h.paragraph_format.space_before = Pt(14 if level == 1 else 10)
    h.paragraph_format.space_after = Pt(6)
    return h


def shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_text(cell, text, bold=False, size=9, color=None, center=False):
    cell.text = ""
    p = cell.paragraphs[0]
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(str(text))
    set_run_font(run, size=size, bold=bold, color=color)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)


def add_table(doc, headers, rows, col_widths=None, header_fill="1B2A4A"):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.autofit = True
    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        set_cell_text(cell, h, bold=True, size=9, color=RGBColor(255, 255, 255), center=True)
        shade_cell(cell, header_fill)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.rows[i + 1].cells[j]
            set_cell_text(cell, val, size=8.5)
            if i % 2 == 1:
                shade_cell(cell, "F0F2F5")
    if col_widths:
        for row in table.rows:
            for j, w in enumerate(col_widths):
                row.cells[j].width = Inches(w)
    doc.add_paragraph()
    return table


def add_caption(doc, text):
    add_para(doc, text, size=9, italic=True, space_after=10, align=WD_ALIGN_PARAGRAPH.CENTER)


def add_image(doc, path, width=6.0):
    path = Path(path)
    if not path.exists():
        add_para(doc, f"[Image missing: {path.name}]", italic=True, size=9)
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(str(path), width=Inches(width))
    p.paragraph_format.space_after = Pt(2)


def build():
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)

    # ---------- Cover ----------
    add_para(doc, "MIT 8339 VIRTUAL LAB", size=12, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(
        doc, "Customer Retention and Performance Intelligence",
        size=20, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4,
    )
    add_para(
        doc,
        "A Business Intelligence and Decision Support System for Retail Card Banking",
        size=12, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14,
    )
    for line in [
        "Course: MIT 8339 – Business Intelligence and Decision Support Systems",
        "Programme: Master of Information Technology (MIT)",
        "Student Name: AKINDEKO TOLULOPE",
        "Student Email: t.akindeko1787@miva.edu.ng",
        "Session: 2026/2027",
        "MIVA Open University – School of Post Graduate Studies",
        f"Code repository: {REPO_URL}",
    ]:
        add_para(doc, line, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)

    doc.add_page_break()

    # ---------- Executive Summary ----------
    add_heading_custom(doc, "Executive Summary", 1)
    add_para(
        doc,
        "This report presents a complete Business Intelligence and Decision Support (BI/DSS) "
        "solution for customer attrition in the retail card portfolio of a Nigerian commercial bank. "
        "Built on the Credit Card Customers (BankChurners) dataset of 10,127 customers, the lab "
        "executes the full BI workflow: data collection, cleaning and transformation, exploratory "
        "analysis and KPI development, a two-page dashboard, and evidence-based decision-support "
        "recommendations. The reproducible codebase is published at GitHub "
        f"({REPO_URL}).",
    )
    add_para(
        doc,
        "Portfolio attrition is 16.1% (1,627 customers). Three data-verified patterns dominate: "
        "a transaction-activity cliff (lowest quartile attrites at 37.1% versus 1.1% in the highest), "
        "card dormancy (0% utilisation attrites at 36.2%), and shallow relationships "
        "(25.6–27.8% attrition at one or two products versus 10.5% at six). A six-rule risk score "
        "yields Low, Medium and High bands with realised attrition of 1.6%, 12.3% and 69.3%; the "
        "High band holds 12.7% of the base yet captures 54.7% of all churn. A logistic regression "
        "(holdout AUC 0.912) validates the same drivers.",
    )
    add_para(
        doc,
        "Five prioritised recommendations follow: dormant-card reactivation, a transaction-activity "
        "lift programme, product cross-sell to single-product customers, a service-contact escalation "
        "trigger, and monthly war-room use of the risk score. Two limitations and NDPR-aligned "
        "governance duties close the report.",
    )

    # ---------- 1 ----------
    add_heading_custom(doc, "1. Business Understanding and KPI Definition", 1)
    add_heading_custom(doc, "1.1 The business problem and the decision to support", 2)
    add_para(
        doc,
        "The bank is losing retail customers across its credit-card portfolio, and management "
        "suspects the losses are linked to activity level, transaction behaviour, product usage, "
        "credit-limit utilisation, complaint patterns and income category. The decision this BI "
        "system must support is concrete: which customer groups receive which retention intervention "
        "next quarter, and how the bank verifies the interventions worked. The deliverable is "
        "therefore a working pipeline plus a dashboard that ranks risk, not a report that merely "
        "describes churn.",
    )
    add_para(
        doc,
        "The analysis follows the classical BI/DSS process chain (Sharda et al., 2020): source data "
        "(cleaned BankChurners extract), transformation (Python/Pandas), analysis (KPIs, segmentation, "
        "rule scoring, logistic validation), reporting (two-page Excel dashboard) and decision support "
        "(prioritised recommendations, re-measured monthly).",
    )

    add_heading_custom(doc, "1.2 Key performance indicators", 2)
    add_para(
        doc,
        "Eight KPIs were defined before analysis so that every visual links to a management question. "
        "Table 1 lists each KPI, its formula and its baseline value from the cleaned dataset.",
    )
    add_table(
        doc,
        ["KPI", "Definition / formula", "Baseline", "Management question"],
        [
            ["Attrition rate", "Attrited customers / total customers", "16.1%", "How large is the leakage?"],
            ["Engaged customer ratio", "Share with High/Very High transaction frequency (68+ txns)", "49.4%", "How deep is the active base?"],
            ["Card dormancy rate", "Share with 0% credit utilisation", "24.4%", "How many cards are idle?"],
            ["Average credit limit", "Mean of Credit_Limit", "$8,632", "What is the portfolio risk envelope?"],
            ["Average utilisation", "Mean of Avg_Utilization_Ratio", "27.5%", "How intensively is credit used?"],
            ["Avg transaction amount", "Mean of Total_Trans_Amt per customer", "$4,404", "What revenue does engagement generate?"],
            ["Avg products per customer", "Mean of Total_Relationship_Count", "3.81", "How sticky are relationships?"],
            ["High-risk share", "Customers in High risk band", "12.7%", "How large is the priority retention list?"],
        ],
        col_widths=[1.5, 2.4, 0.9, 1.8],
    )
    add_caption(doc, "Table 1: KPI definitions and baseline values (n = 10,127)")

    # ---------- 2 ----------
    add_heading_custom(doc, "2. Data Collection, Cleaning and Transformation", 1)
    add_heading_custom(doc, "2.1 Dataset description", 2)
    add_para(
        doc,
        "Table 2 documents the dataset used for this virtual lab: identity, coverage, fitness for "
        "the banking scenario, and acquisition method. The same description is exported as "
        "data/dataset_description.csv in the codebase.",
    )
    add_table(
        doc,
        ["Attribute", "Detail"],
        [
            ["Dataset name", "Credit Card Customers (BankChurners)"],
            ["Primary source", "Goyal, S. (2021). Kaggle – sakshigoyal7/credit-card-customers"],
            ["Acquisition method", "Public mirror of the identical Kaggle CSV (GitHub raw)"],
            ["Records (rows)", "10,127"],
            ["Raw / cleaned columns", "23 raw → 27 cleaned (2 NB columns dropped; 6 fields added)"],
            ["Target variable", "Attrition_Flag (1,627 attrited; 8,500 existing)"],
            ["Domain coverage", "Demographics, products, utilisation, 12-month behaviour, contacts"],
            ["Business scenario fit", "Attrition analysis, segmentation, KPIs, risk scoring, dashboard"],
            ["Alternative considered", "Maven Bank Churn – rejected (missing utilisation/txn fields)"],
            ["Unit of analysis", "Individual credit-card customer (CLIENTNUM)"],
            ["Time window", "Cross-sectional snapshot with 12-month behavioural aggregates"],
            ["Pipeline outputs", "Cleaned/scored CSV; quality log; Excel dashboard; segment tables"],
        ],
        col_widths=[1.8, 4.8],
    )
    add_caption(doc, "Table 2: Dataset description")

    add_heading_custom(doc, "2.2 Data preparation and cleaning decisions", 2)
    add_para(
        doc,
        "Profiling found zero missing cells and zero duplicate CLIENTNUM keys, so no imputation or "
        "de-duplication was needed; the consequential issues were structural. Two embedded Naive Bayes "
        "classifier columns were dropped because they leak a pre-fitted prediction into the feature set. "
        "Three categorical fields carry masked ‘Unknown’ labels — education (1,519), income (1,112), "
        "marital status (749) — retained as explicit categories rather than imputed. An IQR scan flagged "
        "credit-limit and transaction-amount outliers, retained as genuine premium-card and heavy-transactor "
        "behaviour. Table 3 is the preparation and cleaning decisions register (also "
        "data/data_preparation_cleaning_decisions.csv).",
    )
    add_table(
        doc,
        ["Stage", "Issue / check", "Decision", "Justification"],
        [
            ["Acquisition", "Dataset selection (BankChurners vs Maven)", "Use BankChurners via public mirror", "Has utilisation & transaction fields the scenario requires"],
            ["Profiling", "Missing NaN cells = 0", "No imputation", "No true missingness to treat"],
            ["Profiling", "Duplicate CLIENTNUM = 0", "No de-duplication", "Keys already unique"],
            ["Profiling", "Unknown labels (Educ/Income/Marital)", "Retain; map Income → Not Declared", "Avoid fabricating socioeconomic attributes"],
            ["Profiling", "2 Naive_Bayes_Classifier_* columns", "Drop both", "Prevent leakage of pre-fitted predictions"],
            ["Profiling", "Credit_Limit IQR outliers (984)", "Retain", "Genuine premium-card customers"],
            ["Profiling", "Total_Trans_Amt IQR outliers (896)", "Retain", "Genuine heavy transactors"],
            ["Profiling", "Avg_Utilization_Ratio outliers", "Do not cap", "Dormancy is a primary risk signal"],
            ["Profiling", "Customer_Age 26–73", "No correction", "Plausible adult cardholder range"],
            ["Standardise", "Whitespace in categoricals", "Strip headers & labels", "Prevent false segment splits"],
            ["Encode", "Attrition_Flag string", "Create Is_Attrited (0/1)", "Rates, tests, scoring, dashboard formulas"],
            ["Engineer", "Age / income / utilisation / txn count", "Create bands & segments", "Filter-ready for dashboard slicers"],
            ["Engineer", "Spend + limit + products", "Customer_Value_Class", "Link engagement depth to attrition"],
            ["Engineer", "Six behavioural risk rules", "Risk_Score + Churn_Risk_Band", "Explainable High-risk target list (1,285)"],
            ["Export", "Evidence artefacts", "CSV/JSON/Excel Cleaning_Log", "Before/after QC for assessment deliverables"],
        ],
        col_widths=[1.1, 2.0, 1.8, 1.7],
    )
    add_caption(doc, "Table 3: Data preparation and cleaning decisions")

    add_para(doc, "Table 4 summarises the before-and-after quality checks that ship in the dashboard Cleaning_Log sheet.")
    add_table(
        doc,
        ["Check", "Before", "After", "Treatment / justification"],
        [
            ["Rows", "10,127", "10,127", "No rows removed; duplicates absent"],
            ["Columns", "23", "27", "2 leaked NB columns dropped; 6 analytical fields added"],
            ["Missing cells (NaN)", "0", "0", "None at either stage"],
            ["Duplicate CLIENTNUM", "0", "0", "None at either stage"],
            ["Education ‘Unknown’", "1,519", "1,519", "Retained as explicit category"],
            ["Income ‘Unknown’", "1,112", "1,112", "Mapped to Not Declared segment"],
            ["Credit_Limit IQR outliers", "984", "984", "Retained; genuine premium-card customers"],
            ["Total_Trans_Amt IQR outliers", "896", "896", "Retained; genuine heavy transactors"],
            ["Customer_Age range", "26–73", "26–73", "Plausible; no correction needed"],
        ],
        col_widths=[1.8, 0.9, 0.9, 3.0],
    )
    add_caption(doc, "Table 4: Before-and-after data quality checks")

    add_image(doc, FIGS / "screenshot_cleaning_log.png", width=6.2)
    add_caption(doc, "Screenshot 1: Cleaning_Log sheet – preparation evidence in the dashboard workbook")

    add_heading_custom(doc, "2.3 Derived analytical fields", 2)
    add_para(
        doc,
        "Six analytical fields make the raw measures decision-ready: Age_Band and Income_Segment for "
        "demographic filtering; Utilisation_Band separating dormant to heavy card users; "
        "Transaction_Frequency quartiles (Low under 45, Moderate 45–67, High 67–81, Very High above 81); "
        "Customer_Value_Class combining spend, credit-limit and relationship terciles; and Is_Attrited "
        "encoding the target. Risk_Score and Churn_Risk_Band from Section 3.5 complete the set.",
    )

    # ---------- 3 ----------
    add_heading_custom(doc, "3. Exploratory Analysis, KPI Results and Risk Classification", 1)
    add_heading_custom(doc, "3.1 Attrited versus existing customers", 2)
    add_para(
        doc,
        "Attrited customers differ from existing customers on engagement rather than demographics. "
        "They average 44.9 transactions per year against 68.7, spend $3,095 against $4,655, and show "
        "quarter-on-quarter momentum of 0.554 against 0.742. They hold fewer products (3.28 versus 3.92), "
        "have been inactive longer (2.69 versus 2.27 months) and generate more service contacts "
        "(2.97 versus 2.36). Welch t-tests confirm every engagement difference at p < 0.001; age, tenure "
        "and card category are not significant (chi-square p = 0.525): the problem is behavioural, not "
        "product-tier-based.",
    )
    add_table(
        doc,
        ["Variable", "Attrited", "Existing", "Difference", "Welch t", "p-value"],
        [
            ["Total_Trans_Ct (per year)", "44.9", "68.7", "−23.7", "−54.1", "<0.001"],
            ["Total_Trans_Amt ($ per year)", "3,095", "4,655", "−1,560", "−22.7", "<0.001"],
            ["Total_Ct_Chng_Q4_Q1", "0.554", "0.742", "−0.188", "−30.6", "<0.001"],
            ["Total_Relationship_Count", "3.28", "3.92", "−0.64", "−14.9", "<0.001"],
            ["Months_Inactive_12_mon", "2.69", "2.27", "+0.42", "+16.9", "<0.001"],
            ["Contacts_Count_12_mon", "2.97", "2.36", "+0.62", "+20.9", "<0.001"],
            ["Avg_Utilization_Ratio", "0.162", "0.296", "−0.134", "−18.6", "<0.001"],
            ["Total_Revolving_Bal ($)", "673", "1,257", "−584", "−24.1", "<0.001"],
            ["Customer_Age (years)", "46.7", "46.3", "+0.4", "+1.9", "0.058"],
            ["Months_on_book (tenure)", "36.2", "35.9", "+0.3", "+1.4", "0.160"],
        ],
        col_widths=[1.9, 0.9, 0.9, 1.0, 0.8, 0.8],
    )
    add_caption(doc, "Table 5: Attrited versus existing customers – key variable comparison")

    add_heading_custom(doc, "3.2 Pattern 1 – the transaction-activity cliff", 2)
    add_para(
        doc,
        "Attrition risk concentrates in the lowest transaction-frequency quartile: 37.1% of Low-frequency "
        "customers (under 45 transactions per year) have already attrited, against 20.5% (Moderate), "
        "4.5% (High) and 1.1% (Very High) — Figure 1. The pattern holds within every income segment, so "
        "activity, not affluence, is the dividing line.",
    )
    add_image(doc, FIGS / "fig1_transaction_cliff.png", width=5.8)
    add_caption(doc, "Figure 1: Attrition rate by transaction-frequency quartile – the engagement cliff")

    add_heading_custom(doc, "3.3 Pattern 2 – dormant cards, not over-leverage", 2)
    add_para(
        doc,
        "Contrary to the intuition that heavily utilised customers leave, the opposite holds. The 0% "
        "utilisation (non-revolver) band attrites at 36.2%, while the 30–50% and above-50% bands attrite "
        "at only 9.1% and 9.0% (Figure 2). An unused card is an unloved product: dormancy is a leading "
        "indicator of disengagement, and 24.4% of the portfolio currently sits in this band.",
    )
    add_image(doc, FIGS / "fig2_utilisation_bands.png", width=5.8)
    add_caption(doc, "Figure 2: Attrition rate by utilisation band – dormant cards are the risk pocket")

    add_heading_custom(doc, "3.4 Pattern 3 – relationship depth protects revenue", 2)
    add_para(
        doc,
        "Attrition falls as product depth rises: 25.6% at one product and 27.8% at two, versus roughly "
        "11–12% at four to six products (Figure 3). Value class sharpens the picture — Bronze customers "
        "attrite at 39.8%, more than triple the Gold rate of 8.0%.",
    )
    add_image(doc, FIGS / "fig3_products_held.png", width=5.8)
    add_caption(doc, "Figure 3: Attrition rate by number of products held (relationship depth)")

    add_heading_custom(doc, "3.5 Retention-risk scoring and statistical validation", 2)
    add_para(
        doc,
        "A transparent rule-based score gives every point a business meaning. Six behavioural rules each "
        "contribute 0–2 points: dormancy or very low utilisation, two or more inactive months, transaction "
        "volume below the 45/67 quartile lines, weak quarter-on-quarter momentum (under 0.75), shallow "
        "product depth (two or fewer products), and high contact intensity (four or more contacts). The "
        "12-point score maps to Low (0–4), Medium (5–7) and High (8–12) bands. Realised attrition rises "
        "monotonically — 1.6%, 12.3%, 69.3% (Figure 4) — and the High band captures 54.7% of all attrited "
        "customers from just 12.7% of the base (1,285 customers).",
    )
    add_image(doc, FIGS / "fig4_risk_bands.png", width=5.8)
    add_caption(doc, "Figure 4: Risk-band population size versus realised attrition rate")
    add_para(
        doc,
        "To validate the drivers, a logistic regression on ten standardised behavioural features was "
        "trained on 70% of the data and tested on a stratified 30% holdout, reaching AUC 0.912 and "
        "89.6% accuracy. Standardised odds ratios confirm the rule design (Figure 5): transaction count "
        "(OR 0.08 per SD) and momentum (OR 0.52) cut attrition odds most strongly, product count "
        "(OR 0.47) and revolving balance (OR 0.46) protect, while contact intensity (OR 1.68) and "
        "inactivity (OR 1.66) raise it.",
    )
    add_image(doc, FIGS / "fig5_odds_ratios.png", width=5.8)
    add_caption(doc, "Figure 5: Logistic regression odds ratios per standard deviation (holdout AUC 0.912)")

    # ---------- 4 ----------
    add_heading_custom(doc, "4. Dashboard Development and Visualisation", 1)
    add_heading_custom(doc, "4.1 Architecture and interactive controls", 2)
    add_para(
        doc,
        "The dashboard is delivered as an Excel workbook built programmatically with openpyxl "
        "(outputs/MIT8339_Customer_Retention_Dashboard.xlsx), so the build is fully reproducible from "
        "scripts/03_excel_dashboard.py. It contains the two required pages — Executive Overview and "
        "Retention Analysis — supported by a Data sheet with the full cleaned dataset, a Cleaning_Log "
        "sheet, a Review sheet with seven live cross-checks, and a Calc sheet holding every chart’s "
        "source table. Interactivity meets the three-control minimum: two dropdown segment explorers "
        "(income and card category) and AutoFilters on the watchlist and dataset.",
    )
    add_heading_custom(doc, "4.2 Page 1 – Executive overview", 2)
    add_para(
        doc,
        "The executive page leads with eight KPI cards — leakage (16.1%), engaged base (49.4%), "
        "dormancy (24.4%) and the priority list (12.7%) — followed by four charts: portfolio split, "
        "attrition by income segment, age band and customer-value class.",
    )
    add_image(doc, FIGS / "screenshot_executive_overview.png", width=6.2)
    add_caption(doc, "Screenshot 2: Executive Overview dashboard page")

    add_heading_custom(doc, "4.3 Page 2 – Retention analysis", 2)
    add_para(
        doc,
        "The retention page is the working surface for the retention team: the two segment explorers, "
        "four risk charts (risk-band attrition and population, products held, months-inactive trend) "
        "and a sorted income-by-card watchlist with filter controls. Pockets above 25% attrition are "
        "highlighted in red.",
    )
    add_image(doc, FIGS / "screenshot_retention_analysis.png", width=6.2)
    add_caption(doc, "Screenshot 3: Retention Analysis dashboard page with interactive explorers")

    add_heading_custom(doc, "4.4 Why these chart choices", 2)
    add_para(
        doc,
        "Column bars carry category comparisons because bar length is the most accurate pre-attentive "
        "encoding for ranked rates; a line chart is used only for months-inactive, a true ordered scale "
        "where slope carries meaning; and the risk-band charts pair rate with population so a high rate "
        "on a small base cannot mislead. Pie charts and 3-D effects were deliberately avoided.",
    )

    # ---------- 5 ----------
    add_heading_custom(doc, "5. Decision-Support Recommendations", 1)
    add_para(
        doc,
        "Five recommendations follow directly from the quantified findings, ordered by impact against "
        "effort. Each names an owner, the evidence and the KPI it should move.",
    )
    add_heading_custom(doc, "5.1 R1 – Reactivate dormant cards within 30 days", 2)
    add_para(
        doc,
        "Owner: Retail Product Team. Evidence: the 0% utilisation band attrites at 36.2% and holds "
        "2,470 customers. Action: staged reactivation (statement nudge, cash-back on next three "
        "transactions, then a call). Target: dormancy rate below 20% in one quarter.",
    )
    add_heading_custom(doc, "5.2 R2 – Launch an activity-lift programme for the Low transaction band", 2)
    add_para(
        doc,
        "Owner: Segment Marketing. Evidence: customers under 45 transactions per year attrite at 37.1%. "
        "Action: category-spend offers and autopay enrolment for the Low band. Target: engaged customer "
        "ratio above 55%.",
    )
    add_heading_custom(doc, "5.3 R3 – Deepen single-product relationships", 2)
    add_para(
        doc,
        "Owner: Branch and Relationship Managers. Evidence: one-product customers attrite at 25.6%; "
        "product count OR 0.47. Action: next-best-product offer prioritising Bronze and Silver. "
        "Target: average products per customer above 4.2.",
    )
    add_heading_custom(doc, "5.4 R4 – Intervene where service contacts spike", 2)
    add_para(
        doc,
        "Owner: Customer Experience Team. Evidence: attrition climbs with contact count (12.5% at two, "
        "33.5% at five, 100% at six); contacts OR 1.68. Action: treat the fourth contact in twelve months "
        "as an escalation trigger. Target: lower attrition among four-plus-contact customers.",
    )
    add_heading_custom(doc, "5.5 R5 – Operationalise the risk score in a monthly retention war-room", 2)
    add_para(
        doc,
        "Owner: Chief Data Office with Retail Banking. Evidence: High band holds 1,285 customers and "
        "captures 54.7% of churn at 69.3% realised attrition. Action: refresh the score monthly from the "
        "same pipeline and hand the High list to retention agents. Target: portfolio attrition under 12% "
        "within two quarters.",
    )

    # ---------- 6 ----------
    add_heading_custom(doc, "6. Limitations, Governance and Ethical Considerations", 1)
    add_heading_custom(doc, "6.1 Analytical limitations", 2)
    add_para(
        doc,
        "Two limitations bound the conclusions. First, the dataset is a single snapshot: attrition is "
        "observed at one point in time, so the analysis identifies correlates of churn rather than "
        "verified causes. Second, customer value is proxied by spend, credit limit and product count "
        "because the dataset has no revenue fields, and the Not Declared income segment (11% of the base) "
        "limits income-based targeting.",
    )
    add_heading_custom(doc, "6.2 Governance and ethics", 2)
    add_para(
        doc,
        "Privacy: under the Nigeria Data Protection Regulation (NITDA, 2019) the bank must process "
        "customer-level financial behaviour on a lawful basis, restrict access to the retention team, "
        "and prefer aggregated published views. Fairness: the score uses behavioural and product "
        "variables only; income imputation is avoided; the score must never deny credit or price "
        "products — its sole sanctioned use is retention-offer prioritisation. Misuse control: model "
        "logic is documented in the repository; the score must be revalidated against realised attrition "
        "quarterly.",
    )

    # ---------- References ----------
    add_heading_custom(doc, "References", 1)
    refs = [
        "Goyal, S. (2021). Credit card customers [Data set]. Kaggle. https://www.kaggle.com/datasets/sakshigoyal7/credit-card-customers",
        "Hunter, J. D. (2007). Matplotlib: A 2D graphics environment. Computing in Science & Engineering, 9(3), 90–95.",
        "McKinney, W. (2010). Data structures for statistical computing in Python. In Proceedings of the 9th Python in Science Conference (pp. 56–61).",
        "National Information Technology Development Agency. (2019). Nigeria Data Protection Regulation 2019. Federal Republic of Nigeria Official Gazette.",
        "Openpyxl developers. (2024). Openpyxl: A Python library to read/write Excel 2010 xlsx/xlsm files [Computer software]. https://openpyxl.readthedocs.io/",
        "Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., … Duchesnay, É. (2011). Scikit-learn: Machine learning in Python. Journal of Machine Learning Research, 12, 2825–2830.",
        "Sharda, R., Delen, D., & Turban, E. (2020). Analytics, data science, & artificial intelligence: Systems for decision support (11th ed.). Pearson.",
    ]
    for r in refs:
        p = doc.add_paragraph(r)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.first_line_indent = Inches(-0.25)
        p.paragraph_format.space_after = Pt(6)
        for run in p.runs:
            set_run_font(run, size=10)

    # ---------- Appendix A ----------
    doc.add_page_break()
    add_heading_custom(doc, "Appendix A: Code Base (Technical Evidence)", 1)
    add_para(
        doc,
        "The full reproducible pipeline is published at "
        f"{REPO_URL}. Relative paths are used throughout "
        "(scripts/paths.py). Run order:",
    )
    add_para(doc, "1. scripts/01_clean_transform.py — Task 2 cleaning & derived fields", size=10)
    add_para(doc, "2. scripts/02_eda_kpis.py — Task 3 KPIs, EDA, risk score, logistic validation", size=10)
    add_para(doc, "3. scripts/03_excel_dashboard.py — Task 4 two-page Excel dashboard", size=10)
    add_para(doc, "4. scripts/04_report_figures.py — report figures (optional)", size=10)
    add_para(
        doc,
        "Key artefacts: data/cleaned_dataset.csv, data/cleaned_dataset_scored.csv, "
        "data/dataset_description.csv, data/data_preparation_cleaning_decisions.csv, "
        "data/quality_checks.csv, outputs/MIT8339_Customer_Retention_Dashboard.xlsx.",
    )
    add_image(doc, FIGS / "screenshot_console_eda.png", width=6.2)
    add_caption(doc, "Screenshot 4: Console evidence from 02_eda_kpis.py (KPI / EDA / risk-score run)")

    # Include first ~80 lines of clean script as technical evidence sample
    add_heading_custom(doc, "A.1 Excerpt – 01_clean_transform.py (path setup + profiling)", 2)
    clean_src = (SCRIPTS / "01_clean_transform.py").read_text()
    excerpt = "\n".join(clean_src.splitlines()[:55])
    p = doc.add_paragraph()
    run = p.add_run(excerpt)
    set_run_font(run, size=7.5, name="Consolas")
    p.paragraph_format.space_after = Pt(8)

    add_heading_custom(doc, "A.2 Excerpt – paths.py (relative project roots)", 2)
    paths_src = (SCRIPTS / "paths.py").read_text()
    p = doc.add_paragraph()
    run = p.add_run(paths_src)
    set_run_font(run, size=7.5, name="Consolas")

    # ---------- Appendix B ----------
    add_heading_custom(doc, "Appendix B: Cleaned Dataset Data Dictionary (selected fields)", 1)
    add_table(
        doc,
        ["Field", "Type", "Description"],
        [
            ["CLIENTNUM", "ID", "Unique customer identifier"],
            ["Attrition_Flag", "Category", "Existing Customer / Attrited Customer"],
            ["Is_Attrited", "Binary", "1 if attrited, else 0"],
            ["Customer_Age / Age_Band", "Num / Band", "Age in years; 26–34 … 65+"],
            ["Income_Category / Income_Segment", "Category", "Income band; Unknown → Not Declared"],
            ["Card_Category", "Category", "Blue / Silver / Gold / Platinum"],
            ["Credit_Limit", "Currency", "Customer credit limit"],
            ["Avg_Utilization_Ratio / Utilisation_Band", "Ratio / Band", "Credit utilisation; dormancy at 0%"],
            ["Total_Trans_Ct / Transaction_Frequency", "Count / Band", "Annual txn count; quartile bands"],
            ["Total_Trans_Amt", "Currency", "Annual transaction amount"],
            ["Total_Relationship_Count", "Integer", "Number of products held (1–6)"],
            ["Months_Inactive_12_mon", "Integer", "Inactive months in last 12"],
            ["Contacts_Count_12_mon", "Integer", "Service contacts in last 12"],
            ["Customer_Value_Class", "Category", "Bronze / Silver / Gold / Platinum"],
            ["Risk_Score / Churn_Risk_Band", "Score / Band", "0–12 rule score; Low / Medium / High"],
        ],
        col_widths=[2.4, 1.1, 3.1],
    )
    add_caption(doc, "Table 6: Selected fields in the cleaned and scored dataset")

    doc.save(OUT)
    print(f"[SAVE] {OUT}")
    # rough word count of body paragraphs
    words = 0
    for p in doc.paragraphs:
        words += len(p.text.split())
    print(f"[INFO] approx paragraph words: {words}")


if __name__ == "__main__":
    build()
