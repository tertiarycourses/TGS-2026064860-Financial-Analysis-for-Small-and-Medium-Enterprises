"""
SINGLE SOURCE OF TRUTH — course metadata for
WSQ Financial Analysis for Small and Medium Enterprises (TGS-2026064860).

Every artifact (PPT, LP, LG, LG.md, activities index) is generated from this
file + data_domain1..3.py so they stay 100% aligned.
"""

# ------------------------------------------------------------------ metadata
TITLE        = "Financial Analysis for Small and Medium Enterprises"
SHORT_TITLE  = "Financial Analysis for Small and Medium Enterprises"
COURSE_CODE  = "TGS-2026064860"
VERSION      = "v12"
VERSION_DATE = "11 August 2026"
ORG          = "Tertiary Infotech Academy Pte Ltd"
UEN          = "UEN: 201200696W"
TRAINER      = "Dr Alfred Ang"
DAYS         = 2

# ------------------------------------------------------------------ skills framework
TSC_TITLE = "Financial Analysis"
TSC_CODE  = "ACC-MAC-5004-1.1"
TSC_KNOWLEDGE = [
    ("K1", "Statement of financial position"),
    ("K2", "Balance sheet"),
    ("K3", "Income and cash flow statements"),
    ("K4", "Statement of changes in equity"),
    ("K5", "Financial statement analysis techniques"),
]
TSC_ABILITIES = [
    ("A1", "Identify trends by comparing ratios across multiple time periods and statement types"),
    ("A2", "Prepare and interpret performance and position of an organisation using financial statements"),
]

# ------------------------------------------------------------------ outcomes
LEARNING_OUTCOMES = [
    "LO1: Understand the financial statements such as balance sheet, income, and cash flow statements.",
    "LO2: Evaluate organization's financial performance from the trend of financial ratios.",
    "LO3: Analyze financial statements and prepare the organization's position.",
]

# ------------------------------------------------------------------ topics
TOPICS = [
    dict(num=1, code="01",
         title="Understanding Financial Statements",
         subtitle="Overview of Finance and Chart of Accounts · Balance Sheet Statement · Profit and Loss (P&L) Statement · Cash Flow Statement",
         weighting="K1 · K2 · K3",
         concepts=[
            ("Finance & accounting", "Why finance matters to every business decision — the language of the boardroom."),
            ("Chart of accounts", "The general-ledger account structure that every transaction is recorded against."),
            ("Balance sheet", "Assets = Liabilities + Equity — the statement of financial position at a point in time."),
            ("Income statement", "Revenue − Expenses = Profit — financial performance over an accounting period."),
            ("Cash flow statement", "Operating, investing and financing cash movements — where the cash actually went."),
            ("Double-entry system", "Debits equal credits, so the accounting equation always balances."),
         ]),
    dict(num=2, code="02",
         title="Analysing Financial Ratios",
         subtitle="Ratios for Corporate Profitability · Ratios for Corporate Performance · Equity Changes Statement",
         weighting="K4 · A1",
         concepts=[
            ("Liquidity ratios", "Current, quick, cash and operating-cash-flow ratios — can the firm pay its short-term bills?"),
            ("Leverage ratios", "Debt, debt-to-equity, interest coverage — how much risk is carried in the capital structure?"),
            ("Efficiency ratios", "Inventory, receivables, payables and asset turnover — how hard are the assets working?"),
            ("Profitability ratios", "Gross, operating, net margins, ROA, ROCE, ROE, EPS — is the business earning enough?"),
            ("Equity changes statement", "Reconciles opening to closing equity — share capital, dividends, retained earnings."),
            ("Trend analysis", "Compare ratios across periods and statement types to spot direction and risk (A1)."),
         ]),
    dict(num=3, code="03",
         title="Planning & Budgeting using Financial Statements",
         subtitle="Analyse Financial Statements · Financial Planning · Capital Budgeting",
         weighting="K5 · A2",
         concepts=[
            ("Budgeting", "Baseline, incremental, zero-based and hybrid budgets — planning income vs spending."),
            ("Capital budgeting", "Payback, discounted payback, NPV, PI and IRR — evaluating long-term investments."),
            ("Time value of money", "PV = FV / (1+i)^n — a dollar today is worth more than a dollar tomorrow."),
            ("Forecast & variance", "Budget vs actual vs forecast — favourable and adverse variances, thresholds."),
            ("Financial health", "Read the balance sheet, income and cash flow statements together, then ratio-check."),
            ("Analysis methods", "Ratio, horizontal (trend) and vertical analysis, plus industry benchmarking (A2)."),
         ]),
]

# ------------------------------------------------------------------ day themes (8 training hours/day)
DAY_THEMES = {
    1: "Financial Statements & Financial Ratios",
    2: "Planning, Budgeting & Final Assessment",
}

# ------------------------------------------------------------------ assessment
ASSESSMENT = dict(
    written="Written Assessment (WA) — Short-Answer Questions (SAQ), 1 hour, open book.",
    practical="Practical Performance (PP) — financial-analysis tasks on real statements, 1 hour, open book.",
    note="A minimum of 75% attendance is required to be eligible for assessment and funding.",
)

RECOMMENDED_COURSES = [
    "WSQ - Quickbooks Accounting System for Small and Medium Enterprises",
    "WSQ - Unlocking the Power of Accounting and Tax Systems for SMEs",
    "WSQ - Statistical Data Analysis with Excel for Beginners",
    "WSQ - Budgeting for Small and Medium Enterprises",
    "WSQ - Software Automation with Excel VBA Programming",
]
