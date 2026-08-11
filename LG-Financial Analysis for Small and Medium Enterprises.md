# Financial Analysis for Small and Medium Enterprises — Learner Guide

**WSQ Course Code:** TGS-2026064860  |  **Conducted by:** Tertiary Infotech Academy Pte Ltd (UEN 201200696W)  |  **Version v13 · 11 August 2026**

## Contents

- [Introduction](#introduction)
- [Course Learning Outcomes](#course-learning-outcomes)
- [Skills Framework (TSC)](#skills-framework-tsc)
- [Before You Start — Setup](#before-you-start--setup)
- [Topic 01 — Understanding Financial Statements  (K1 · K2 · K3)](#topic-01--understanding-financial-statements--k1--k2--k3)
  - [Activity 1 — Prepare a Cash Flow Statement](#activity-1--prepare-a-cash-flow-statement)
- [Topic 02 — Analysing Financial Ratios  (K4 · A1)](#topic-02--analysing-financial-ratios--k4--a1)
  - [Activity 2 — Ratio Analysis of Two Companies](#activity-2--ratio-analysis-of-two-companies)
  - [Activity 3 — Trend Analysis over Four Years](#activity-3--trend-analysis-over-four-years)
- [Topic 03 — Planning & Budgeting using Financial Statements  (K5 · A2)](#topic-03--planning--budgeting-using-financial-statements--k5--a2)
  - [Activity 4 — Payback Period & Discounted Payback Period](#activity-4--payback-period--discounted-payback-period)
  - [Activity 5 — Net Present Value & Profitability Index](#activity-5--net-present-value--profitability-index)
  - [Activity 6 — Solvency & Financial Risk Analysis](#activity-6--solvency--financial-risk-analysis)
- [Revision Pointers for the Final Assessment](#revision-pointers-for-the-final-assessment)
- [Glossary](#glossary)


## Introduction

This Learner Guide accompanies the WSQ course Financial Analysis for Small and Medium Enterprises (TGS-2026064860), conducted by Tertiary Infotech Academy Pte Ltd. It provides the key concepts for each of the three topics and detailed step-by-step instructions for all six hands-on activities. Each activity analyses a realistic mock-data workbook (Excel + CSV, fictitious Singapore SMEs) from its own folder under activities/ — also downloadable from the course LMS (https://lms-tms.tertiaryinfotech.com) — and is completed in Microsoft Excel or Google Sheets.

Use this guide alongside the course slides. The final assessment is open book: you may refer to the slides, this Learner Guide and any approved materials, so keep your completed activity worksheets — they are your best revision notes.


## Course Learning Outcomes

- LO1: Understand the financial statements such as balance sheet, income, and cash flow statements.
- LO2: Evaluate organization's financial performance from the trend of financial ratios.
- LO3: Analyze financial statements and prepare the organization's position.


## Skills Framework (TSC)

TSC Title: Financial Analysis · TSC Code: ACC-MAC-5004-1.1

- K1: Statement of financial position
- K2: Balance sheet
- K3: Income and cash flow statements
- K4: Statement of changes in equity
- K5: Financial statement analysis techniques
- A1: Identify trends by comparing ratios across multiple time periods and statement types
- A2: Prepare and interpret performance and position of an organisation using financial statements


## Before You Start — Setup

**What you need**

- A laptop with Microsoft Excel (2016 or later) or a Google account for Google Sheets.
- Access to the course LMS at https://lms-tms.tertiaryinfotech.com — log in with your registered email (an OTP is sent to you).
- The activity data: each activity has its own self-contained folder (activities/activity01 … activity06, mirrored in the Activities folder on the LMS) holding the activity guide (.md), a formatted mock-data Excel workbook (.xlsx), the raw data as .csv files and a printable worksheet (.pdf).
- All company data is fictitious and prepared for training use only.
- A calculator (or the spreadsheet itself) for the ratio and discounting computations.

**Conventions used in every activity**

- Figures in the Topic 1 activity are in $ million; Topic 2 and 3 activities use dollars.
- Workings shown in a box are the expected calculation or Excel formula for that step.
- Each activity ends with a 'Test it' check — compare your result against it before moving on.
- Model answers are discussed in class after each activity.


## Topic 01 — Understanding Financial Statements  (K1 · K2 · K3)

Overview of Finance and Chart of Accounts · Balance Sheet Statement · Profit and Loss (P&L) Statement · Cash Flow Statement

**Key concepts**

- Finance & accounting — Why finance matters to every business decision — the language of the boardroom.
- Chart of accounts — The general-ledger account structure that every transaction is recorded against.
- Balance sheet — Assets = Liabilities + Equity — the statement of financial position at a point in time.
- Income statement — Revenue − Expenses = Profit — financial performance over an accounting period.
- Cash flow statement — Operating, investing and financing cash movements — where the cash actually went.
- Double-entry system — Debits equal credits, so the accounting equation always balances.


### Activity 1 — Prepare a Cash Flow Statement

Maps to: LO1 — understand the balance sheet, income and cash flow statements (K1, K2, K3).

Goal: Havenwood Trading Pte Ltd is a fictitious Singapore SME distributor. From its FY2018/FY2019 balance sheets and FY2019 income statement (all in S$ million), derive the full statement of cash flows — operating, investing and financing — and reconcile it to the movement in cash.

**What you'll produce**

A complete statement of cash flows (operating / investing / financing) for Havenwood Trading, derived from the balance sheet movements and the income statement.   (Tools: Microsoft Excel · data workbook FA-Activity-01-Cash-Flow-Statement.xlsx in activities/activity01.)

**Data provided**

- Workbook: FA-Activity-01-Cash-Flow-Statement.xlsx — sheets: 'Balance Sheet' (FY2019 in column B, FY2018 in column C), 'Income Statement' (FY2019 in column B), 'Your Workings' (answer template)
- Balance Sheet (FY2019 | FY2018): PP&E 100 | 90 · Receivables 50 | 50 · Inventory 30 | 30 · Cash 20 | 20 · Equity 80 | 80 · Bank loans 50 | 55 · Payables 40 | 35 · Provisions 30 | 20 (S$m)
- Income Statement FY2019: Revenue 50 (wholesale 38 + retail 12) · COGS 10 · Operational costs 20 (salaries 12, rental 4, utilities/logistics 4) · Depreciation 5 · Interest 3 · Tax 2 · Net income 10 (S$m)

**Step-by-step**

1. Open FA-Activity-01-Cash-Flow-Statement.xlsx from activities/activity01 (or download it from the Activities folder on the LMS). Study the 'Balance Sheet' sheet: FY2019 is column B, FY2018 is column C. Note that both years balance: Total assets (row 11) = Total equity and liabilities (row 17) = 200 and 190.
2. Read the 'Income Statement' sheet from top to bottom and confirm the profit cascade: Revenue 50 − COGS 10 − Operational costs 20 = EBITDA 20 (cell B16); then − Depreciation 5 − Interest 3 − Tax 2 = Net income 10 (cell B20).

   ```
   B16 = B9 - B10 - B15 = 50 - 10 - 20 = 20
   ```

3. Go to the 'Your Workings' sheet. Start the OPERATING block with EBITDA from the income statement: enter 20. Cash flow always starts from operating earnings BEFORE the non-cash charges (depreciation is added back by starting at EBITDA rather than net income).

   ```
   EBITDA = 20
   ```

4. Deduct the tax actually PAID in cash, not the P&L charge: the S$2m FY2019 charge plus a S$5m prior-year settlement = S$7m out. Enter −7. (Lesson: the P&L tax line and the tax cash flow are rarely the same number.)

   ```
   Tax paid = -(2 + 5) = -7
   ```

5. Compute the working-capital adjustment from the balance-sheet movements: receivables and inventory were held flat while the business grew, but S$5m of supplier bills from FY2018 were settled during the year before new payables built up — a net S$5m use of cash. Enter −5.

   ```
   Working-capital adjustment = -5   (NWC: payables 40 − receivables 50 − inventory 30 = -40 vs -45 in FY2018)
   ```

6. Add the change in provisions — provisions rose from 20 to 30 (row 16), a non-cash source of 10: enter +10. Total the OPERATING block.

   ```
   Operating cash flow = 20 - 7 - 5 + 10 = 18
   ```

7. Build the INVESTING block from the fixed-asset movement. PP&E (net) rose 90 → 100 while 5 of depreciation was charged, so gross capex = change in fixed assets 10 + depreciation 5 = 15 spent; 1 of it was still owed to equipment vendors (payables on capex). Enter −10, −5, +1 and total: −14.

   ```
   Investing cash flow = -10 - 5 + 1 = -14
   ```

8. Build the FINANCING block: bank loans fell 55 → 50 (repayment −5); interest paid −3; dividends/drawings took out the year's retained profit movement −6 (equity stayed at 80 despite net income 10 less adjustments). Total: −14.

   ```
   Financing cash flow = -5 - 3 - 6 = -14
   ```

9. Compute the NET cash flow: 18 − 14 − 14 = −10, and write one sentence of interpretation: Havenwood is operationally cash-generative (+18) but spent more on capex and debt service than operations produced this year — the cash buffer fell by S$10m.

   ```
   Net cash flow = 18 - 14 - 14 = -10
   ```

10. Cross-check the story against the three statements: profit was +10 yet cash moved −10 — point to the capex (−15 gross) and loan repayment (−5) as the difference between profit and cash. This is the K3 insight the assessment asks for.

**Test it**

Your three blocks total +18 (operating), −14 (investing) and −14 (financing), net −10 — matching the model answer — and you can explain in one sentence why a profitable year still drained cash.

> **Note:** Everything for this activity lives in one folder — activities/activity01/ of the course repository, and the matching folder in the Activities folder on the LMS: this guide, the mock-data Excel workbook, the raw data as CSV files, and a printable worksheet PDF.

---


## Topic 02 — Analysing Financial Ratios  (K4 · A1)

Ratios for Corporate Profitability · Ratios for Corporate Performance · Equity Changes Statement

**Key concepts**

- Liquidity ratios — Current, quick, cash and operating-cash-flow ratios — can the firm pay its short-term bills?
- Leverage ratios — Debt, debt-to-equity, interest coverage — how much risk is carried in the capital structure?
- Efficiency ratios — Inventory, receivables, payables and asset turnover — how hard are the assets working?
- Profitability ratios — Gross, operating, net margins, ROA, ROCE, ROE, EPS — is the business earning enough?
- Equity changes statement — Reconciles opening to closing equity — share capital, dividends, retained earnings.
- Trend analysis — Compare ratios across periods and statement types to spot direction and risk (A1).


### Activity 2 — Ratio Analysis of Two Companies

Maps to: LO2 — evaluate financial performance from financial ratios (K4, A1).

Goal: Orchid Logistics Pte Ltd (Company X) and Marina Retail Group Pte Ltd (Company Y) are two fictitious SMEs. From their FY2025 balance sheets and income statements, compute the liquidity, profitability, turnover and solvency ratios for both companies and give a one-line verdict per ratio family on which company performs better.

**What you'll produce**

A completed ratio-analysis worksheet comparing Company X and Company Y across four ratio families, with a one-line verdict per ratio.   (Tools: Microsoft Excel · data workbook FA-Activity-02-Ratio-Analysis.xlsx in activities/activity02.)

**Data provided**

- Workbook: FA-Activity-02-Ratio-Analysis.xlsx — sheets: 'Balance Sheet' (X in column B, Y in column C), 'Income Statement' (X in B, Y in C), 'Your Workings' (ratio template)
- Balance Sheet (X | Y): Cash 50,000 | 80,000 · Inventory 40,000 | 70,000 · Receivables 100,000 | 250,000 · Property 400,000 | 450,000 · Other FA 60,000 | 50,000 · Total assets 650,000 | 900,000
- Liabilities & equity (X | Y): Payables 100,000 | 150,000 · ST loans 80,000 | 70,000 · LT liabilities 300,000 | 500,000 · Capital 170,000 | 180,000
- Income Statement (X | Y): Sales 300,000 | 500,000 · COGS 140,000 | 200,000 · Gross profit 160,000 | 300,000 · EBIT 100,000 | 228,000 · Interest 10,000 | 18,000 · Net income 72,000 | 168,000

**Step-by-step**

1. Open FA-Activity-02-Ratio-Analysis.xlsx. On the 'Balance Sheet' sheet, Company X is column B and Company Y is column C. First aggregate the raw ingredients on the 'Your Workings' sheet: Current assets = Cash + Inventory + Receivables (rows 7–9); Current liabilities = Payables + ST loans (rows 15–16).

   ```
   CA(X) ='Balance Sheet'!B7+B8+B9 = 190,000 · CL(X) = B15+B16 = 180,000
CA(Y) = 400,000 · CL(Y) = 220,000
   ```

2. LIQUIDITY — Current Ratio = Current Assets / Current Liabilities. Compute for both companies and interpret against the ≥1 benchmark from the slides.

   ```
   X: 190,000/180,000 = 1.06 · Y: 400,000/220,000 = 1.82 → Y is more liquid
   ```

3. LIQUIDITY — Quick Ratio = (Current Assets − Inventory) / Current Liabilities. Inventory is the least liquid current asset, so strip it out and recompute.

   ```
   X: 150,000/180,000 = 0.83 · Y: 330,000/220,000 = 1.50
   ```

4. PROFITABILITY — Operating Profit Ratio = EBIT / Sales, using 'Income Statement' row 12 (EBIT) over row 6 (Sales). Express as a percentage.

   ```
   X: 100,000/300,000 = 33% · Y: 228,000/500,000 = 46%
   ```

5. PROFITABILITY — Net Profit Ratio = Net Income / Sales (row 16 over row 6). Note how the gap versus the operating margin reveals the drag from interest and tax.

   ```
   X: 72,000/300,000 = 24% · Y: 168,000/500,000 = 34%
   ```

6. PROFITABILITY — Return on Capital Employed = EBIT / Capital Employed, where Capital Employed = Equity + Long-Term Liabilities (the long-term funding of the business).

   ```
   Cap employed X: 170,000+300,000 = 470,000 → 100,000/470,000 = 21%
Cap employed Y: 180,000+500,000 = 680,000 → 228,000/680,000 = 34%
   ```

7. TURNOVER — Inventory Turnover = COGS / Average Inventory (use the year-end balance as the average). A higher figure means stock moves faster.

   ```
   X: 140,000/40,000 = 3.50 · Y: 200,000/70,000 = 2.86 → X turns stock faster
   ```

8. TURNOVER — Receivables Turnover = Credit Sales / Average Receivables (treat all sales as credit sales). Then convert to days: Average Collection Period = 365 / turnover.

   ```
   X: 300,000/100,000 = 3.00 → 122 days · Y: 500,000/250,000 = 2.00 → 183 days
   ```

9. TURNOVER — Payables Turnover = Purchases / Average Payables, assuming purchases = 50% of COGS.

   ```
   X: 70,000/100,000 = 0.70 · Y: 100,000/150,000 = 0.67
   ```

10. SOLVENCY — Debt-to-Equity = Long-Term Liabilities / Equity, and Financial Leverage = Total Assets / Equity. Higher values = more of the balance sheet is funded by lenders.

   ```
   D/E — X: 300,000/170,000 = 1.76 · Y: 500,000/180,000 = 2.78
Leverage — X: 650,000/170,000 = 3.82 · Y: 900,000/180,000 = 5.00
   ```

11. Write the verdicts in the 'Your Workings' sheet, one line per family: Y wins on liquidity and profitability; X wins on turnover efficiency and carries materially less solvency risk. Note how Y's higher ROE partly comes from leverage, not just operations — the DuPont insight.

**Test it**

Your worksheet shows all ten ratios for both companies matching the model answers (e.g. Current Ratio X 1.06 vs Y 1.82; D/E X 1.76 vs Y 2.78), with a one-line verdict per ratio family.

> **Note:** Everything for this activity lives in one folder — activities/activity02/ of the course repository, and the matching folder in the Activities folder on the LMS: this guide, the mock-data Excel workbook, the raw data as CSV files, and a printable worksheet PDF.

---


### Activity 3 — Trend Analysis over Four Years

Maps to: LO2 — identify trends by comparing ratios across time periods (K5, A1).

Goal: Sunrise F&B Holdings Pte Ltd is a fictitious cafe-chain SME. From four years (FY2018–FY2021) of balance sheets and income statements, compute year-on-year growth for every line, derive the key ratios for each year, and present your observations on where the business is heading.

**What you'll produce**

A completed trend-analysis template: YoY growth for every balance-sheet and income-statement line plus a four-year ratio table with your observations.   (Tools: Microsoft Excel · data workbook FA-Activity-03-Trend-Analysis.xlsx in activities/activity03.)

**Data provided**

- Workbook: FA-Activity-03-Trend-Analysis.xlsx — sheets: 'Balance Sheet 2018-2021' and 'Income Statement 2018-2021' (FY2021 in column B through FY2018 in column E), 'Trend Template'
- Sales FY2018→FY2021: 180,000 → 198,000 → 217,800 → 239,580 (10% p.a.)
- Net income FY2018→FY2021: 53,600 → 64,601 → 76,872 → 90,548
- Total assets FY2018→FY2021: 330,000 → 324,800 → 329,983 → 339,581

**Step-by-step**

1. Open FA-Activity-03-Trend-Analysis.xlsx. Years run LEFT to RIGHT from newest: FY2021 (column B) to FY2018 (column E). Skim both statement sheets and note what is obviously moving: sales up every year, payables up sharply, property drifting down.
2. Compute YoY growth for each income-statement line: growth = (this year − last year) / last year. In Excel, put the formula beside FY2021 and fill across.

   ```
   =(B6-C6)/C6 → Sales FY2021 growth = (239,580-217,800)/217,800 = 10.0%
   ```

3. Read the growth pattern: Sales grow a steady 10% while COGS grows only 5% — so Gross Profit growth runs ahead of sales (14.5% → 13.8%) and EBIT growth is faster still. Fixed S&D expenses (flat 20,000) create operating leverage.

   ```
   COGS: =(B7-C7)/C7 = 5.0% each year
   ```

4. Repeat for the balance sheet lines. Flag the working-capital lines: Accounts payable grew 25.0% in FY2021 (75,000 vs 60,000) while inventory grew 17.6% — the company is stretching its suppliers.

   ```
   =(B15-C15)/C15 = 25.0%
   ```

5. Now derive the LIQUIDITY ratios per year in the 'Trend Template': Current Ratio = (Cash + Inventory + Receivables) / (Payables + ST loans); Acid Ratio strips inventory.

   ```
   FY2021: 155,583/95,000 = 1.64 · FY2018: 145,000/68,000 = 2.13
Acid: FY2021 1.22 · FY2018 1.47
   ```

6. Derive the EFFICIENCY ratios per year: Asset Turnover = Sales/Total assets; Inventory Turnover = COGS/Inventory; Receivables Turnover = Sales/Receivables; Payables Turnover = Purchases/Payables with purchases = 50% of COGS.

   ```
   Asset turnover: 0.55 → 0.72 · Inventory turnover: 4.00 → 6.48
Receivables: 3.00 → 3.29 · Payables: 0.99 → 0.73
   ```

7. Derive the LEVERAGE ratios per year: Debt-to-Equity = (ST loans + LT liabilities)/Capital; Debt-to-Assets = (ST loans + LT liabilities + Payables)/Total assets.

   ```
   D/E: 0.56 (FY2018) → 0.75 (FY2021) · D/A: 0.36 → 0.43
   ```

8. Derive the PROFITABILITY ratios per year: Operating Margin = EBIT/Sales; Net Profit Margin = Net income/Sales.

   ```
   Operating margin: 0.40 → 0.49 · Net margin: 0.30 → 0.38
   ```

9. Write your observations under the template — cover all four families: (1) profitability and efficiency improve steadily (margins, asset and inventory turnover all up); (2) liquidity deteriorates (current ratio 2.13 → 1.64) as payables are stretched; (3) leverage creeps up (D/E 0.56 → 0.75); (4) verdict: a profitable, tightening business — watch the falling current ratio and the supplier stretch.

**Test it**

Your ratio table matches the model answers (e.g. FY2021 Current Ratio 1.64, Net Profit Margin 0.38, D/E 0.75) and your observations cover profitability, efficiency, liquidity and leverage.

> **Note:** Everything for this activity lives in one folder — activities/activity03/ of the course repository, and the matching folder in the Activities folder on the LMS: this guide, the mock-data Excel workbook, the raw data as CSV files, and a printable worksheet PDF.

---


## Topic 03 — Planning & Budgeting using Financial Statements  (K5 · A2)

Analyse Financial Statements · Financial Planning · Capital Budgeting

**Key concepts**

- Budgeting — Baseline, incremental, zero-based and hybrid budgets — planning income vs spending.
- Capital budgeting — Payback, discounted payback, NPV, PI and IRR — evaluating long-term investments.
- Time value of money — PV = FV / (1+i)^n — a dollar today is worth more than a dollar tomorrow.
- Forecast & variance — Budget vs actual vs forecast — favourable and adverse variances, thresholds.
- Financial health — Read the balance sheet, income and cash flow statements together, then ratio-check.
- Analysis methods — Ratio, horizontal (trend) and vertical analysis, plus industry benchmarking (A2).


### Activity 4 — Payback Period & Discounted Payback Period

Maps to: LO3 — evaluate a capital investment using payback methods (K5, A2).

Goal: Havenwood Trading is evaluating a refrigerated delivery van: S$2,000k outlay in Year 0, saving S$375k per year for 10 years versus outsourced cold-chain delivery. Compute the simple payback period, then discount each cash flow at the 10% required return to find the discounted payback.

**What you'll produce**

A payback worksheet showing cumulative cash flow per year, the simple payback point (5.33 years) and the discounted payback point (≈8 years).   (Tools: Microsoft Excel · data workbook FA-Activity-04-Payback-Period.xlsx in activities/activity04.)

**Data provided**

- Workbook: FA-Activity-04-Payback-Period.xlsx — sheet 'Project Cash Flows': Year in column A (rows 6–16), net cash flow in column B; columns C–E are your templates (cumulative, discounted, cumulative discounted)
- Cash flows: Year 0: −S$2,000k · Years 1–10: +S$375k per year (all figures S$'000)
- Discount rate: 10% — Havenwood's required return on internal projects

**Step-by-step**

1. Open FA-Activity-04-Payback-Period.xlsx, sheet 'Project Cash Flows'. Year 0 (the −2,000 outlay) is row 6; Years 1–10 with +375 each are rows 7–16.
2. Build the cumulative cash flow in column C: start at C7 with =B7, then C8 =C7+B8 and fill down to C16. Watch where the running total passes 2,000.

   ```
   C7 =B7 → 375 · C8 =C7+B8 → 750 … C12 → 2,250 (passes 2,000 in Year 6)
   ```

3. Compute the simple payback period exactly: full years until the outlay is nearly recovered (5 years → 1,875), plus the fraction of Year 6 needed: (2,000 − 1,875)/375.

   ```
   Payback = 5 + 125/375 = 5.33 years  (= 2000/375 for an even stream)
   ```

4. Now build the discounted column D: each year's cash flow divided by (1.10)^year. Enter =B7/(1.1)^A7 in D7 and fill down.

   ```
   D7 = 375/1.1 = 340.91 · D8 = 309.92 · D9 = 281.74 · D10 = 256.13 · D11 = 232.85
   ```

5. Build the cumulative discounted column E the same way as column C and read down for the 2,000 crossing: 340.91, 650.83, 932.57, 1,188.70, 1,421.55, 1,633.22, 1,825.66, 2,000.60 — it crosses during Year 8.

   ```
   E14 (Year 8) = 2,000.60 ≥ 2,000 → discounted payback ≈ 8.0 years
   ```

6. Interpret the gap on the 'Your Workings' sheet: discounting stretches the payback from 5.33 to ≈8 years because later savings are worth less today. State when payback is the right tool (quick liquidity screen) and its blind spots (ignores cash flows AFTER payback, and simple payback ignores the time value of money).

**Test it**

Your cumulative discounted cash flow reaches S$2,000.60k at Year 8 — the discounted payback (≈8 years) is materially longer than the simple payback (5.33 years), and you can say why.

> **Note:** Everything for this activity lives in one folder — activities/activity04/ of the course repository, and the matching folder in the Activities folder on the LMS: this guide, the mock-data Excel workbook, the raw data as CSV files, and a printable worksheet PDF.

---


### Activity 5 — Net Present Value & Profitability Index

Maps to: LO3 — evaluate a capital investment using NPV and PI (K5, A2).

Goal: Using the same delivery-van proposal (S$2,000k outlay, S$375k/year for 10 years, 10% required return), compute the Net Present Value and the Profitability Index, and make the accept/reject recommendation to management.

**What you'll produce**

An NPV/PI worksheet: PV of each inflow, NPV ≈ S$304k, PI ≈ 1.15, and a reasoned accept decision.   (Tools: Microsoft Excel · data workbook FA-Activity-05-NPV-PI.xlsx in activities/activity05.)

**Data provided**

- Workbook: FA-Activity-05-NPV-PI.xlsx — sheet 'Project Cash Flows': Year in column A (rows 6–16), net cash flow in column B, discounted-cash-flow template in column C; 'Your Workings' for NPV/PI
- Cash flows: Year 0: −S$2,000k · Years 1–10: +S$375k per year (S$'000)
- Required return: 10%

**Step-by-step**

1. Open FA-Activity-05-NPV-PI.xlsx, sheet 'Project Cash Flows'. Discount each inflow in column C with =B7/(1.1)^A7 filled down rows 7–16.

   ```
   341, 310, 282, 256, 233, 212, 192, 175, 159, 145  (S$'000, rounded)
   ```

2. Sum the discounted inflows to get the present value of what the van gives back.

   ```
   PV of inflows =SUM(C7:C16) = 2,304
   ```

3. Compute NPV = PV of inflows − outlay. Cross-check with Excel's NPV function — note it discounts from Year 1, so the Year-0 outlay stays outside it.

   ```
   NPV = 2,304 - 2,000 = 304   (=NPV(10%,B7:B16)+B6)
   ```

4. Compute the Profitability Index = PV of inflows / PV of outflows — the value created per dollar invested.

   ```
   PI = 2,304 / 2,000 = 1.15
   ```

5. Make the decision on 'Your Workings': NPV > 0 AND PI > 1 → ACCEPT. Write the recommendation as you would to management: 'the van creates S$304k of value at our 10% hurdle; every dollar invested returns S$1.15 of present value'.

   ```
   Accept: NPV 304 > 0 · PI 1.15 > 1
   ```

6. Stress-test the recommendation: raise the discount rate in your formulas to 14% and observe NPV turn slightly negative — the project's IRR is ≈13.4%, so the accept decision holds only while the cost of capital stays below that. Note this sensitivity in one line.

   ```
   At i = 14%: PV inflows ≈ 1,956 → NPV ≈ -44 → reject
   ```


**Test it**

Your worksheet shows NPV ≈ S$304k and PI ≈ 1.15 with an ACCEPT decision, plus a one-line sensitivity note (NPV turns negative near a 13–14% discount rate).

> **Note:** Everything for this activity lives in one folder — activities/activity05/ of the course repository, and the matching folder in the Activities folder on the LMS: this guide, the mock-data Excel workbook, the raw data as CSV files, and a printable worksheet PDF.

---


### Activity 6 — Solvency & Financial Risk Analysis

Maps to: LO3 — compare companies' solvency risk from their statements (K5, A2).

Goal: Return to Orchid Logistics (Company X) and Marina Retail Group (Company Y). Using their balance sheets and income statements, compare the two companies' solvency risk — debt-to-equity, financial leverage and interest coverage — and conclude which is the riskier borrower.

**What you'll produce**

A solvency comparison table (D/E, leverage, interest coverage) for X and Y with a reasoned risk verdict.   (Tools: Microsoft Excel · data workbook FA-Activity-06-Solvency-Analysis.xlsx in activities/activity06.)

**Data provided**

- Workbook: FA-Activity-06-Solvency-Analysis.xlsx — same X/Y statements as Activity 2 ('Balance Sheet' and 'Income Statement', X in column B, Y in column C) plus a solvency 'Your Workings' template
- Balance Sheet (X | Y): Total assets 650,000 | 900,000 · LT liabilities 300,000 | 500,000 · Capital 170,000 | 180,000
- Income Statement (X | Y): EBIT 100,000 | 228,000 · Interest 10,000 | 18,000

**Step-by-step**

1. Open FA-Activity-06-Solvency-Analysis.xlsx. Solvency looks at the LONG-term survival of the capital structure, so pull three ingredients per company from the 'Balance Sheet' sheet: long-term liabilities (row 17), capital (row 18) and total assets (row 13); and two from the 'Income Statement': EBIT (row 12) and interest (row 13).
2. Compute Debt-to-Equity = Long-Term Liabilities / Equity for both companies — how many dollars of long-term debt sit on each dollar the owners have at stake.

   ```
   X: 300,000/170,000 = 1.76 · Y: 500,000/180,000 = 2.78
   ```

3. Compute Financial Leverage = Total Assets / Equity — the equity multiplier. At 5.00, only a fifth of Y's balance sheet is owner-funded.

   ```
   X: 650,000/170,000 = 3.82 · Y: 900,000/180,000 = 5.00
   ```

4. Compute Interest Coverage = EBIT / Interest Expense — how many times operating profit covers the interest bill this year.

   ```
   X: 100,000/10,000 = 10.0× · Y: 228,000/18,000 = 12.7×
   ```

5. Weigh the evidence like a credit officer: Y carries structurally more debt per dollar of equity (D/E 2.78, leverage 5.00) but currently earns stronger cover (12.7×). Ask the trend question: what happens to Y's cover if EBIT halves in a downturn? X: 5.0× — comfortable; Y: 6.3× but on a far bigger debt load that must be refinanced.

   ```
   Stress: halve EBIT → X 5.0× · Y 6.3×, with 500,000 of LT debt to refinance
   ```

6. Write the verdict on 'Your Workings': Company Y has the higher structural solvency risk — its capital structure depends on lenders — even though its current interest coverage is stronger. Recommend the leverage and coverage covenants you would monitor.

**Test it**

Your table matches the model answers (D/E 1.76 vs 2.78; leverage 3.82 vs 5.00; coverage 10.0× vs 12.7×) and your verdict identifies Company Y's capital structure as the riskier one, with reasons.

> **Note:** Everything for this activity lives in one folder — activities/activity06/ of the course repository, and the matching folder in the Activities folder on the LMS: this guide, the mock-data Excel workbook, the raw data as CSV files, and a printable worksheet PDF.

---


## Revision Pointers for the Final Assessment

- K1/K2 — Be able to explain the statement of financial position: assets, liabilities, equity, and how the balance sheet equation stays in balance.
- K3 — Be able to trace the income statement from revenue to net income, and explain why profit is not cash (link to the cash flow statement).
- K4 — Know the statement of changes in equity: Beginning Equity + Net Income − Dividends ± Other changes = Ending Equity.
- K5 — Be able to pick the right analysis technique: ratio analysis, horizontal (trend) analysis, vertical analysis, benchmarking.
- A1 — Practise reading ratio trends across periods: what do a falling current ratio and a rising debt-to-equity together tell you?
- A2 — Practise the DuPont decomposition: ROE = Net Profit Margin × Asset Turnover × Equity Multiplier.
- Bring your completed activity worksheets — the assessment is open book.


## Glossary

- **Balance sheet** — Statement of financial position: Assets = Liabilities + Equity, as at a date.
- **Income statement** — Profit & Loss statement: Profits = Revenues − Expenses, over a period.
- **Cash flow statement** — Cash movements in operating, investing and financing activities.
- **Statement of changes in equity** — Reconciles opening to closing equity for the period.
- **Working capital** — Current assets minus current liabilities.
- **EBITDA** — Earnings before interest, tax, depreciation and amortization.
- **EBIT** — Earnings before interest and tax — operating profit.
- **Current ratio** — Current assets / current liabilities — a liquidity measure.
- **Quick (acid test) ratio** — (Current assets − inventories) / current liabilities.
- **Debt-to-equity ratio** — Total liabilities (or long-term debt) / shareholders' equity.
- **Inventory turnover** — COGS / average inventories — stock efficiency.
- **ROE / ROA** — Return on equity / return on assets — profitability of capital.
- **Time value of money** — PV = FV / (1+i)^n — discounting future cash to today.
- **NPV** — Net present value: PV of inflows − initial outlay; accept if > 0.
- **Profitability index** — PV of inflows / PV of outflows; accept if > 1.
- **IRR** — The discount rate at which NPV equals zero.
- **Payback period** — Time for cumulative cash inflows to repay the initial outlay.
- **Variance** — Difference between budget and actual — adverse or favourable.
- **Horizontal analysis** — Percent change of the same line item across periods.
- **Vertical analysis** — Each line as a percent of a base figure (sales or total assets).
- **DuPont analysis** — ROE = Net Profit Margin × Asset Turnover × Equity Multiplier.
