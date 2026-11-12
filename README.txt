================================================================================
MIT 8339: Business Intelligence & Decision Support Systems (Virtual Lab)
Customer Retention and Performance Intelligence Codebase
================================================================================

Student:    Akindeko Tolulope (t.akindeko1787@miva.edu.ng)
Repository: https://github.com/xodejunkie/MIT8339_VirtualLab
Dataset:    Credit Card Customers (BankChurners) - Kaggle sakshigoyal7
Target:     Customer Attrition (Attrition_Flag: Attrited vs Existing)


1. PROJECT OVERVIEW & BUSINESS SCENARIO
---------------------------------------
This repository contains the end-to-end Business Intelligence and Decision Support
pipeline developed for the MIT 8339 Virtual Lab.

The project addresses customer attrition in commercial retail credit card banking.
Analyzing a portfolio of 10,127 cardholders, the objective is to move from reactive
churn management to proactive, data-driven retention intelligence.

The pipeline delivers:
  * Systematic data hygiene, profiling, and leakage removal.
  * Behavioral feature engineering and multi-dimensional customer segmentation.
  * Portfolio KPI architecture and statistical hypothesis testing (Welch t-tests, Chi-square).
  * An explainable, rule-based 12-point Churn Risk Score and 3 operational risk bands.
  * Predictive validation via stratified Logistic Regression (AUC = 0.912).
  * An automated, interactive two-page Excel BI Dashboard built with openpyxl.
  * Publication-ready analytical visualizations for executive reporting.


2. KEY ANALYTICAL FINDINGS & METRICS
------------------------------------
  * Portfolio Attrition Rate: 16.07% (1,627 attrited customers out of 10,127).
  * The Engagement Cliff: Customers with Low transaction volume (<= 45 txns/year)
    exhibit an attrition rate of 37.1%, compared to just 4.8% for Very High transactors.
  * Card Dormancy: 24.39% of the portfolio (2,470 customers) have 0% credit card
    utilisation. Realised attrition in this dormant pocket reaches 36.2%.
  * Relationship Depth: Customers holding only 1 or 2 products attrite at 25.6%
    and 27.8%, whereas holders of 4 to 6 products attrite at only 11% - 12%.
  * Service Contact Friction: Attrition climbs steeply with customer service contact
    frequency (12.5% at 2 contacts, 33.5% at 5 contacts, and 100% at 6 contacts).
  * Risk Score Efficiency: The High-risk band (scores 8-12) captures 54.7% of all
    portfolio churn from only 12.69% of the customer base (1,285 customers) with
    a realised attrition rate of 69.3%.


3. PIPELINE ARCHITECTURE & SCRIPT BREAKDOWN
-------------------------------------------
The codebase executes in a clean 4-stage sequential pipeline. All scripts resolve
paths dynamically relative to the project root via scripts/paths.py:

  Stage 1: scripts/01_clean_transform.py (Task 2 - Data Preparation & Cleaning)
  -----------------------------------------------------------------------------
  * Loads data/BankChurners_raw.csv (10,127 rows x 23 columns).
  * Drops the 2 pre-fitted Naive Bayes classifier columns to eliminate predictive leakage.
  * Strips whitespace from column headers and categorical string values.
  * Formally audits categorical "Unknown" labels in Education_Level (1,519),
    Marital_Status (749), and Income_Category (1,112). Preserves them as explicit
    "Not Declared" categories to avoid fabricating socioeconomic attributes.
  * Runs IQR outlier scans across 9 numerical metrics (fences Q1 - 1.5*IQR, Q3 + 1.5*IQR).
    Retains high-value transactors and premium credit limits as genuine business signals.
  * Engineers 6 analytical fields:
      - Is_Attrited: Binary target (1 = Attrited Customer, 0 = Existing Customer).
      - Age_Band: 26-34, 35-44, 45-54, 55-64, 65+.
      - Income_Segment: Low (<$40K) to High (>$120K) plus Not Declared.
      - Utilisation_Band: 0% (Non-revolver), 0-10%, 10-30%, 30-50%, >50%.
      - Transaction_Frequency: Quartile activity tiers (Low, Moderate, High, Very High).
      - Customer_Value_Class: Composite tercile score across spend, limit, and products
        (Bronze, Silver, Gold, Platinum).
  * Exports: data/cleaned_dataset.csv, data/quality_checks.csv, data/cleaning_log.json.

  Stage 2: scripts/02_eda_kpis.py (Task 3 - Exploratory Analysis, KPIs & Modeling)
  --------------------------------------------------------------------------------
  * Computes overall portfolio KPIs (spend, utilisation, dormancy, engaged ratio).
  * Conducts Welch's t-tests on numerical drivers, confirming statistically significant
    declines among attriters in transaction count (t = -65.4), spend, and revolving balance.
  * Conducts Chi-square tests of independence across demographic and product categories.
  * Generates segment attrition tables across income, card tier, age, activity, and value class.
  * Implements the 12-point Rule-Based Retention Risk Score (0-2 pts across 6 rules):
      - R1 (Utilisation): 0% = 2 pts; <= 10% = 1 pt.
      - R2 (Inactivity): >= 4 months = 2 pts; 2-3 months = 1 pt.
      - R3 (Transaction Count): < 45 = 2 pts; 45-67 = 1 pt.
      - R4 (Momentum): Q4/Q1 transaction count ratio < 0.75 = 2 pts; 0.75-1.0 = 1 pt.
      - R5 (Relationship Depth): <= 2 products = 2 pts; 3 products = 1 pt.
      - R6 (Service Contacts): >= 4 contacts = 2 pts; 2-3 contacts = 1 pt.
  * Maps total score into operational risk bands:
      - Low Risk:    Score 0 - 4  (Attrition:  1.6%, n = 5,190)
      - Medium Risk: Score 5 - 7  (Attrition: 12.3%, n = 3,652)
      - High Risk:   Score 8 - 12 (Attrition: 69.3%, n = 1,285)
  * Validates behavioral drivers via stratified Logistic Regression (70% train / 30% test):
      - Test ROC-AUC = 0.912, Accuracy = 89.6%.
      - Extracts standardized odds ratios confirming transaction frequency (OR 0.08)
        and product count (OR 0.47) as primary protective drivers.
  * Exports: data/cleaned_dataset_scored.csv, data/kpis.json, data/segments/*.csv.

  Stage 3: scripts/03_excel_dashboard.py (Task 4 - Two-Page BI Excel Dashboard)
  -----------------------------------------------------------------------------
  * Programmatically compiles outputs/MIT8339_Customer_Retention_Dashboard.xlsx
    using openpyxl without external workbook templates.
  * Page 1 ("Executive Overview"):
      - 8 styled executive KPI cards with format masks.
      - 4 bar charts: Portfolio Status Split, Attrition by Income Segment,
        Attrition by Age Band, and Attrition by Customer Value Class.
  * Page 2 ("Retention Analysis"):
      - 2 interactive Dropdown Segment Explorers (Income Segment and Card Category)
        powered by live COUNTIF / AVERAGEIF / COUNTIFS Excel formulas.
      - 4 risk & behavioral charts: Risk Band Attrition, Risk Band Customer Count,
        Attrition by Products Held, and Months Inactive Trend Line.
      - Income x Card Attrition Watchlist with AutoFilter and conditional red flags
        highlighting pockets with > 25% attrition.
  * Supporting Worksheets:
      - Data: Complete scored dataset (10,127 rows x 30 cols) with frozen headers & AutoFilter.
      - Cleaning_Log: Quality audit trail and derived field rule definitions.
      - Review: 7 live formula cross-checks validating dashboard numbers against source data.
      - Calc: Backing aggregation tables feeding all charts.

  Stage 4: scripts/04_report_figures.py (Reporting Visualizations)
  ---------------------------------------------------------------
  * Generates high-resolution matplotlib charts for documentation and briefings:
      - fig1_transaction_cliff.png: Engagement cliff visualization.
      - fig2_utilisation_bands.png: Dormancy risk pocket breakdown.
      - fig3_products_held.png: Attrition decay by relationship depth.
      - fig4_risk_bands.png: Risk band population vs realised attrition.
      - fig5_odds_ratios.png: Logistic regression standardized odds ratios.
      - fig6_months_inactive.png: Inactivity duration trend.


4. REPOSITORY STRUCTURE & ARTIFACT CATALOG
------------------------------------------
MIT8339_BI_VirtualLab_Codebase/
├── README.txt                                # Project documentation & pipeline guide
├── requirements.txt                          # Python dependencies
├── .gitignore                                # Excluded artifacts & local files
├── data/
│   ├── BankChurners_raw.csv                  # Raw Kaggle extract (10,127 x 23)
│   ├── cleaned_dataset.csv                   # Cleaned dataset (Task 2 deliverable)
│   ├── cleaned_dataset_scored.csv            # Scored dataset with Risk_Score & Risk_Band
│   ├── dataset_description.csv               # Dataset attribute & metadata table
│   ├── data_preparation_cleaning_decisions.csv # Cleaning decisions register
│   ├── quality_checks.csv                    # Before / after profiling quality log
│   ├── cleaning_log.json                     # Machine-readable cleaning log & fences
│   ├── kpis.json                             # Summary KPIs, statistical tests, model metrics
│   ├── segments/                             # Segment aggregation tables (*.csv)
│   │   ├── by_age_band.csv
│   │   ├── by_card.csv
│   │   ├── by_contacts.csv
│   │   ├── by_income.csv
│   │   ├── by_months_inactive.csv
│   │   ├── by_products.csv
│   │   ├── by_risk_band.csv
│   │   ├── by_trans_frequency.csv
│   │   ├── by_utilisation.csv
│   │   ├── by_value_class.csv
│   │   ├── watch_income_card_count.csv
│   │   └── watch_income_card_rate.csv
│   └── figs/                                 # Report figures & dashboard screenshots
│       ├── fig1_transaction_cliff.png
│       ├── fig2_utilisation_bands.png
│       ├── fig3_products_held.png
│       ├── fig4_risk_bands.png
│       ├── fig5_odds_ratios.png
│       ├── fig6_months_inactive.png
│       ├── screenshot_cleaning_log.png
│       ├── screenshot_console_eda.png
│       ├── screenshot_executive_overview.png
│       └── screenshot_retention_analysis.png
├── outputs/
│   └── MIT8339_Customer_Retention_Dashboard.xlsx # Two-page interactive BI dashboard
└── scripts/
    ├── paths.py                              # Shared relative path resolver
    ├── xlsx_style.py                         # Corporate styling module for openpyxl
    ├── 01_clean_transform.py                 # Task 2: Data cleaning & transformation
    ├── 02_eda_kpis.py                        # Task 3: EDA, KPIs, risk scoring & validation
    ├── 03_excel_dashboard.py                 # Task 4: Automated Excel workbook compiler
    └── 04_report_figures.py                  # Analytical visualization generator


5. SETUP & REPRODUCTION
-----------------------
Prerequisites:
  * Python 3.10+ (tested on Python 3.10, 3.11, 3.12, 3.14)
  * Standard virtual environment recommended

Installation:
  python3 -m pip install -r requirements.txt

Run Order:
  cd scripts
  python3 01_clean_transform.py
  python3 02_eda_kpis.py
  python3 03_excel_dashboard.py
  python3 04_report_figures.py      # Optional: regenerates report figures

All scripts resolve file locations dynamically relative to the project root;
they can be executed from within `scripts/` or from the repository root.


6. STRATEGIC RECOMMENDATIONS (DECISION SUPPORT)
-----------------------------------------------
The pipeline translates findings into 5 concrete operational interventions:
  * R1 (Dormancy Reactivation): Target the 2,470 cardholders with 0% utilisation
    via automated statement nudges and a 3-transaction cashback incentive.
  * R2 (Activity Lift): Target the 2,611 cardholders with <= 45 transactions/year
    via category spend campaigns and recurring utility bill autopay enrolment.
  * R3 (Cross-Sell Relationship Deepening): Prioritize 1- and 2-product holders in
    Bronze/Silver tiers with pre-approved savings or consumer credit lines.
  * R4 (Friction Escalation Trigger): Route customers reaching 4 service contacts
    in 12 months directly to a senior customer experience resolution desk.
  * R5 (Monthly Retention War-Room): Establish monthly operational reviews using the
    scored priority list (1,285 High-Risk accounts) to deploy proactive outreach.
================================================================================
