MIT 8339 - Business Intelligence and Decision Support Systems (Virtual Lab)
CODE BASE - Customer Retention and Performance Intelligence
=============================================================

Student: AKINDEKO TOLULOPE (t.akindeko1787@miva.edu.ng)
Repository: https://github.com/xodejunkie/MIT8339_VirtualLab

Folder layout
-------------
  MIT8339_BI_VirtualLab_Report_T.AKINDEKO1787@MIVA.EDU.NG.docx
                               Analytical report (updated)
  data/BankChurners_raw.csv          Raw dataset
  data/cleaned_dataset.csv           Cleaned export (Task 2)
  data/cleaned_dataset_scored.csv    Scored dataset with risk bands
  data/dataset_description.csv       Dataset description table
  data/data_preparation_cleaning_decisions.csv
                                     Preparation & cleaning decisions register
  data/quality_checks.csv            Before/after quality log
  data/cleaning_log.json             Cleaning decisions JSON
  data/kpis.json                     KPI pack + model metrics
  data/segments/*.csv                Segment attrition tables
  data/figs/*.png                    Report figures + screenshots
  outputs/MIT8339_Customer_Retention_Dashboard.xlsx
  scripts/01_clean_transform.py
  scripts/02_eda_kpis.py
  scripts/03_excel_dashboard.py
  scripts/04_report_figures.py
  scripts/05_build_report_docx.py    Rebuild the Word report
  scripts/paths.py                   Shared relative paths
  scripts/xlsx_style.py              Self-contained Excel styling

Setup
-----
  python3 -m pip install -r requirements.txt
  python3 -m pip install python-docx   # only needed to rebuild the report

RUN ORDER (Python 3.10+)
------------------------
  cd scripts
  python3 01_clean_transform.py
  python3 02_eda_kpis.py
  python3 03_excel_dashboard.py
  python3 04_report_figures.py      # optional report figures
  python3 05_build_report_docx.py   # optional rebuild of the DOCX

Input data : data/BankChurners_raw.csv (Kaggle Credit Card Customers / sakshigoyal7)
All paths are relative to this project folder.
