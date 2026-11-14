MIT 8339 Business Intelligence and Decision Support Systems: Virtual Lab
Customer Retention and Performance Intelligence

Student:    Akindeko Tolulope (t.akindeko1787@miva.edu.ng)
Repository: https://github.com/xodejunkie/MIT8339_VirtualLab
Dataset:    Credit Card Customers (BankChurners), Kaggle: sakshigoyal7/credit-card-customers


What this is
------------
Code and outputs for the MIT 8339 virtual lab. The scenario is a Nigerian bank
losing retail credit-card customers. The data is the public BankChurners file
(10,127 customers, US-dollar values), used as a stand-in because no Nigerian
bank data was available, so the numbers below describe that dataset and not a
Nigerian portfolio.

The scripts clean the data, work out KPIs and segment attrition rates, build a
six-rule retention-priority score, check the drivers with a logistic
regression, and produce a two-page Excel dashboard.


Main results
------------
- 16.1% of customers left (1,627 of 10,127).
- Customers with fewer than 45 transactions a year left at 36.5%; the most
  active group (more than 81) left at 1.1%.
- 2,470 customers (24.4%) had 0% credit utilisation, and 36.2% of them left.
- One- and two-product customers left at 25.6% and 27.8%, against about
  11-12% for customers with four to six products.
- Attrition rises with service contacts: 12.5% at two, 33.5% at five, and
  every customer with six contacts left.
- Priority bands from the score: Low 3,286 customers (1.6% left), Medium 5,556
  (12.3%), High 1,285 (69.3%). The High band is 12.7% of customers but 54.7% of
  leavers. These are historical rates on the same data the rules were built on,
  not forecasts.
- Logistic regression on a 30% holdout: AUC 0.912, accuracy 89.6%, and for the
  customers who left, precision 0.75, recall 0.53, F1 0.62.


How to run
----------
Python 3.10 or later.

  python3 -m pip install -r requirements.txt
  cd scripts
  python3 01_clean_transform.py
  python3 02_eda_kpis.py
  python3 03_excel_dashboard.py
  python3 04_report_figures.py      # optional, redraws the report charts

Every script finds its files relative to the project root (scripts/paths.py),
so the folder can be moved anywhere.


What each script does
---------------------
01_clean_transform.py
  Loads data/BankChurners_raw.csv (23 columns), drops the two Naive Bayes
  columns (they hold another model's predictions and would leak the answer),
  trims whitespace, keeps "Unknown" education, income and marital labels as
  their own category, and runs an IQR outlier scan (outliers are kept). Adds
  six fields: Is_Attrited, Age_Band, Income_Segment, Utilisation_Band,
  Transaction_Frequency (under 45 / 45-67 / 68-81 / above 81) and
  Customer_Value_Class. Writes data/cleaned_dataset.csv (27 columns),
  data/quality_checks.csv and data/cleaning_log.json.

02_eda_kpis.py
  Works out the KPIs, Welch t-tests and chi-square tests, and the segment
  tables in data/segments/. Scores every customer with six rules worth 0-2
  points each (utilisation, inactive months, transaction count, change in
  activity, products held, contacts) and cuts the 0-12 total into Low (0-4),
  Medium (5-7) and High (8-12). Fits the logistic regression (70/30 split) and
  writes precision, recall, F1 and the confusion matrix to data/kpis.json.
  Writes data/cleaned_dataset_scored.csv (30 columns).

03_excel_dashboard.py
  Builds outputs/MIT8339_Customer_Retention_Dashboard.xlsx with openpyxl.
  Executive Overview: eight KPI cards and four charts.
  Retention Analysis: income and card-category dropdown explorers, four charts
  and an income-by-card watchlist with AutoFilter (pockets above 25% attrition
  in red). The Data, Cleaning_Log, Review and Calc sheets hold the scored data,
  the cleaning log, seven live cross-checks and the chart source tables.

04_report_figures.py
  Draws the matplotlib charts in data/figs/ used in the written report.

paths.py and xlsx_style.py are shared helpers (file locations, and the
workbook's colours, fonts and chart setup).


Files
-----
data/BankChurners_raw.csv                     raw Kaggle file
data/cleaned_dataset.csv                      cleaned data, 27 columns
data/cleaned_dataset_scored.csv               cleaned data plus the score, 30 columns
data/dataset_description.csv                  dataset description table
data/data_preparation_cleaning_decisions.csv  cleaning decisions
data/quality_checks.csv                       before/after checks
data/cleaning_log.json                        profile, Unknown counts, IQR fences
data/kpis.json                                KPIs, tests, model results
data/segments/*.csv                           attrition by segment
data/figs/*.png                               report charts and screenshots
outputs/MIT8339_Customer_Retention_Dashboard.xlsx   the dashboard


Recommendations from the analysis
---------------------------------
R1  Reactivate the 2,470 zero-utilisation cards (statement message, cash-back
    on the next three transactions, then a call).
R2  Offer spend rewards and autopay to the 2,482 customers with fewer than 45
    transactions a year.
R3  Make a next-best-product offer to one- and two-product customers, starting
    with Bronze and Silver.
R4  Escalate a customer's fourth service contact in twelve months.
R5  Refresh the score monthly and give the High-band list (1,285 customers) to
    the retention team.
