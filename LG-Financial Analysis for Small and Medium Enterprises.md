# Financial Analysis for Small and Medium Enterprises — Learner Guide

**WSQ Course Code:** TGS-2026064860  |  **Conducted by:** Tertiary Infotech Academy Pte Ltd (UEN 201200696W)  |  **Version v11 · 10 August 2026**

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

This Learner Guide accompanies the WSQ course Financial Analysis for Small and Medium Enterprises (TGS-2026064860), conducted by Tertiary Infotech Academy Pte Ltd. It provides the key concepts for each of the three topics and detailed step-by-step instructions for all six hands-on activities. Each activity uses a worksheet downloadable from the course LMS (https://lms-tms.tertiaryinfotech.com) and is completed in Microsoft Excel or Google Sheets.

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
- The activity worksheets, downloaded from your course page on the LMS.
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

Goal: You are given the Balance Sheet (2018 and 2019) and the Income Statement (2019) of a company, all figures in $ million. Working in Excel or on paper, derive the statement of cash flows — operating, investing and financing — and reconcile it to the change in cash.

**What you'll produce**

A complete statement of cash flows (operating / investing / financing) derived from the balance sheet and income statement.   (Tools: Microsoft Excel (or Google Sheets), activity worksheet from the LMS.)

**Data provided**

- Balance Sheet 2019: Fixed Assets 100 · Receivables 50 · Inventory 30 · Cash 20 · Total Assets 200 | Equity 80 · Loans 50 · Payables 40 · Provisions 30 · Total 200
- Balance Sheet 2018: Fixed Assets 90 · Receivables 50 · Inventory 30 · Cash 20 · Total Assets 190 | Equity 80 · Loans 55 · Payables 35 · Provisions 20 · Total 190
- Income Statement 2019: Revenue 50 · COGS 10 · Operational costs 20 · Depreciation 5 · Interest 3 · Tax 2 · Net Income 10

**Step-by-step**

1. Download the activity worksheet 'Cash Flow Statement' from the LMS (lms-tms.tertiaryinfotech.com) and open it in Excel.
2. Compute EBITDA from the income statement: Revenue − COGS − Operational costs = 50 − 10 − 20 = 20.

   ```
   EBITDA = 20
   ```

3. Compute Net Working Capital movement: NWC = Payables − Receivables − Inventory = 40 − 50 − 30 = −40 (2019) vs 35 − 50 − 30 = −45 (2018); the change is −5.

   ```
   ΔNWC = −5
   ```

4. Build the Operating Cash Flow block: EBITDA 20, Tax −6, Change in NWC −5, Change in Provisions +10 → Operating Cash Flow = 18.

   ```
   OCF = 20 − 6 − 5 + 10 = 18
   ```

5. Build the Investing Cash Flow block: Change in Fixed Assets −10, Depreciation 2019 −5, Change in trade payables on capex +1 → Investing Cash Flow = −14.

   ```
   ICF = −10 − 5 + 1 = −14
   ```

6. Build the Financing Cash Flow block: Change in financial debt −5 (loans 55 → 50), Interest cost −3, Change in earnings −6 → Financing Cash Flow = −14.

   ```
   FCF = −5 − 3 − 6 = −14
   ```

7. Compute the Net Cash Flow and reconcile: 18 − 14 − 14 = −10.

   ```
   Net Cash Flow = −10
   ```


**Test it**

Your three blocks total 18 (operating), −14 (investing) and −14 (financing), giving a net cash flow of −10 that reconciles with the balance-sheet movement.

> **Note:** The full worksheet for this activity is available on the LMS, and a printable copy is in activities/activity-01 of the course repository.

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

Goal: You are a financial analyst evaluating two companies, X and Y, from their balance sheets and income statements. Compute the liquidity, profitability, turnover and solvency ratios for both companies and comment on which company performs better on each parameter.

**What you'll produce**

A completed ratio-analysis worksheet comparing Company X and Company Y across four ratio families, with a one-line verdict per ratio.   (Tools: Microsoft Excel (or Google Sheets), activity worksheet from the LMS.)

**Data provided**

- Balance Sheet (X | Y): Cash 50,000 | 80,000 · Inventory 40,000 | 70,000 · AR 100,000 | 250,000 · Property 400,000 | 450,000 · Other FA 60,000 | 50,000 · Total 650,000 | 900,000
- Liabilities (X | Y): AP 100,000 | 150,000 · ST Loans 80,000 | 70,000 · LT Liabilities 300,000 | 500,000 · Capital 170,000 | 180,000
- Income Statement (X | Y): Sales 300,000 | 500,000 · COGS 140,000 | 200,000 · S&D 50,000 | 60,000 · EBIT 100,000 | 228,000 · Interest 10,000 | 18,000 · Tax 18,000 | 42,000 · Net Income 72,000 | 168,000

**Step-by-step**

1. Download the activity worksheet 'Ratio Analysis' from the LMS and open it in Excel.
2. Liquidity — Current Ratio = Current Assets / Current Liabilities. X: 190,000/180,000 = 1.06; Y: 400,000/220,000 = 1.82. Y is more liquid.

   ```
   X 1.06 · Y 1.82
   ```

3. Liquidity — Quick Ratio = (Current Assets − Inventory) / Current Liabilities. X: 150,000/180,000 = 0.83; Y: 330,000/220,000 = 1.50.

   ```
   X 0.83 · Y 1.50
   ```

4. Profitability — Operating Profit Ratio = EBIT / Sales. X: 100,000/300,000 = 33%; Y: 228,000/500,000 = 46%.

   ```
   X 33% · Y 46%
   ```

5. Profitability — Net Profit Ratio = Net Income / Sales. X: 72,000/300,000 = 24%; Y: 168,000/500,000 = 34%.

   ```
   X 24% · Y 34%
   ```

6. Profitability — Return on Capital Employed = EBIT / Capital Employed. X: 100,000/470,000 = 21%; Y: 228,000/680,000 = 34%.

   ```
   X 21% · Y 34%
   ```

7. Turnover — Inventory Turnover = COGS / Average Inventory. X: 140,000/40,000 = 3.50; Y: 200,000/70,000 = 2.86. X turns stock faster.

   ```
   X 3.50 · Y 2.86
   ```

8. Turnover — Receivables Turnover = Credit Sales / Average Receivables. X: 300,000/100,000 = 3.00; Y: 500,000/250,000 = 2.00.

   ```
   X 3.00 · Y 2.00
   ```

9. Turnover — Payables Turnover (assume purchases = 50% of COGS). X: 70,000/100,000 = 0.70; Y: 100,000/150,000 = 0.67.

   ```
   X 0.70 · Y 0.67
   ```

10. Solvency — Debt-to-Equity = Long-Term Debt / Equity. X: 300,000/170,000 = 1.76; Y: 500,000/180,000 = 2.78. Y is more leveraged.

   ```
   X 1.76 · Y 2.78
   ```

11. Solvency — Financial Leverage = Total Assets / Equity. X: 650,000/170,000 = 3.82; Y: 900,000/180,000 = 5.00.

   ```
   X 3.82 · Y 5.00
   ```

12. Write a one-line verdict per ratio family: Y wins on liquidity and profitability; X wins on turnover efficiency and carries less solvency risk.

**Test it**

Your worksheet shows all ten ratios for both companies and your verdicts match the model answers (e.g. Current Ratio X 1.06 vs Y 1.82; D/E X 1.76 vs Y 2.78).

> **Note:** The full worksheet for this activity is available on the LMS, and a printable copy is in activities/activity-02 of the course repository.

---


### Activity 3 — Trend Analysis over Four Years

Maps to: LO2 — identify trends by comparing ratios across time periods (K5, A1).

Goal: You are given four years (2018–2021) of a company's balance sheet and income statement. Compute year-on-year growth for every line item, derive the key ratios for each year, and present your observations on the company's direction.

**What you'll produce**

A completed trend-analysis template: YoY growth for every balance-sheet and income-statement line plus a four-year ratio table with your observations.   (Tools: Microsoft Excel (or Google Sheets), activity template from the LMS.)

**Data provided**

- Sales 2018→2021: 180,000 → 198,000 → 217,800 → 239,580 (10% p.a.)
- Net Income 2018→2021: 53,600 → 64,601 → 76,872 → 90,548
- Total Assets 2018→2021: 330,000 → 324,800 → 329,983 → 339,581

**Step-by-step**

1. Download the activity template 'Trend Analysis' from the LMS and open it in Excel.
2. Fill the YoY Growth columns of the balance sheet: growth = (this year − last year) / last year. E.g. Inventory 2021 = (40,000 − 34,000)/34,000 = 17.65%.

   ```
   =(C5-E5)/E5
   ```

3. Fill the YoY Growth columns of the income statement. Sales grow a steady 10% each year while COGS grows only 5% — so Gross Profit growth accelerates (14.47% → 13.78%).
4. Compute the liquidity ratios per year: Current Ratio falls from 2.13 (2018) to 1.64 (2021); Acid Ratio falls from 1.47 to 1.22.

   ```
   Current = CA/CL
   ```

5. Compute the efficiency ratios per year: Asset Turnover rises 0.55 → 0.72; Inventory Turnover rises 4.00 → 6.48; Receivables Turnover rises 3.00 → 3.29.
6. Compute the leverage ratios per year: Debt-to-Equity rises 0.56 → 0.75; Debt-to-Assets rises 0.36 → 0.43.
7. Compute the profitability ratios per year: Operating Margin rises 0.40 → 0.49; Net Profit Margin rises 0.30 → 0.38.
8. Write your observations: profitability and efficiency improve steadily; liquidity declines and leverage creeps up — flag the falling current ratio as the trend to watch.

**Test it**

Your ratio table matches the model answers (e.g. 2021 Current Ratio 1.64, Net Profit Margin 0.38) and your observations cover profitability, efficiency, liquidity and leverage trends.

> **Note:** The full worksheet for this activity is available on the LMS, and a printable copy is in activities/activity-03 of the course repository.

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

Goal: A project requires an initial outlay of $2,000 and returns $375 per year for 10 years. Compute the simple payback period, then repeat the exercise discounting each cash flow at 10% to find the discounted payback period.

**What you'll produce**

A payback worksheet showing cumulative cash flow per year, the simple payback point (~5.3 years) and the discounted payback point (~8 years).   (Tools: Microsoft Excel (or Google Sheets), activity worksheet from the LMS.)

**Data provided**

- Cash flows: Year 0: −$2,000; Years 1–10: +$375 per year
- Discount rate: 10% for the discounted payback

**Step-by-step**

1. Download the activity worksheet 'Cash Flow Payback' from the LMS and open it in Excel.
2. Build the cumulative cash flow column: 375, 750, 1,125, 1,500, 1,875, 2,250 … the cumulative figure passes $2,000 during Year 6.

   ```
   =SUM($B$2:B7)
   ```

3. Compute the simple payback period: 2,000 / 375 = 5.33 years (about 5 years 4 months).

   ```
   Payback = 2000/375 = 5.33 yr
   ```

4. Add a discounted cash flow column: DCF = 375 / (1.10)^n → 340.91, 309.92, 281.74, 256.13, 232.85, 211.68, 192.43, 174.94 …

   ```
   =375/(1.1)^A2
   ```

5. Build the cumulative discounted column: 340.91, 650.83, 932.57, 1,188.70, 1,421.55, 1,633.22, 1,825.66, 2,000.60 — it crosses $2,000 in Year 8.
6. State the discounted payback period: ≈ 8 years, and explain why discounting lengthens the payback.

   ```
   Discounted payback ≈ 8.0 yr
   ```


**Test it**

Your cumulative discounted cash flow reaches $2,000.60 at Year 8 — the discounted payback (~8 years) is materially longer than the simple payback (5.33 years).

> **Note:** The full worksheet for this activity is available on the LMS, and a printable copy is in activities/activity-04 of the course repository.

---


### Activity 5 — Net Present Value & Profitability Index

Maps to: LO3 — evaluate a capital investment using NPV and PI (K5, A2).

Goal: Using the same project ($2,000 outlay, $375/year for 10 years, 10% required return), compute the Net Present Value and the Profitability Index, and decide whether to accept the project.

**What you'll produce**

An NPV/PI worksheet: PV of each inflow, NPV ≈ $304, PI ≈ 1.15, and an accept/reject decision.   (Tools: Microsoft Excel (or Google Sheets), activity worksheet from the LMS.)

**Data provided**

- Cash flows: Year 0: −$2,000; Years 1–10: +$375 per year
- Required return: 10%

**Step-by-step**

1. Download the activity worksheet 'Discounted Cash Flow' from the LMS and open it in Excel.
2. Discount each inflow at 10%: 341, 310, 282, 256, 233, 212, 192, 175, 159, 145.

   ```
   =375/(1.1)^A2
   ```

3. Sum the discounted inflows: PV of cash inflows = $2,304.

   ```
   PV inflows = 2,304
   ```

4. Compute NPV = PV of inflows − initial outlay = 2,304 − 2,000 = $304 (or use =NPV(10%, range) − 2000).

   ```
   NPV = 304
   ```

5. Compute the Profitability Index = PV of inflows / PV of outflows = 2,304 / 2,000 = 1.15.

   ```
   PI = 1.15
   ```

6. Decide: NPV > 0 and PI > 1 → accept the project; note how the decision rule would flip if the discount rate rose.

**Test it**

Your worksheet shows NPV ≈ $304 and PI ≈ 1.15, and your decision is ACCEPT because NPV > 0 and PI > 1.

> **Note:** The full worksheet for this activity is available on the LMS, and a printable copy is in activities/activity-05 of the course repository.

---


### Activity 6 — Solvency & Financial Risk Analysis

Maps to: LO3 — compare companies' solvency risk from their statements (K5, A2).

Goal: Using the balance sheets and income statements of companies X and Y, compare the two companies in terms of solvency risk: debt-to-equity, financial leverage and interest coverage. Conclude which company is the riskier borrower.

**What you'll produce**

A solvency comparison table (D/E, leverage, interest coverage) for X and Y with a risk verdict.   (Tools: Microsoft Excel (or Google Sheets), activity worksheet from the LMS.)

**Data provided**

- Balance Sheet (X | Y): Total Assets 650,000 | 900,000 · LT Debt 300,000 | 500,000 · Equity 170,000 | 180,000
- Income Statement (X | Y): EBIT 100,000 | 228,000 · Interest 10,000 | 18,000

**Step-by-step**

1. Download the activity worksheet 'Solvency Analysis' from the LMS and open it in Excel.
2. Compute Debt-to-Equity = Long-Term Debt / Equity. X: 300,000/170,000 = 1.76; Y: 500,000/180,000 = 2.78.

   ```
   X 1.76 · Y 2.78
   ```

3. Compute Financial Leverage = Total Assets / Equity. X: 650,000/170,000 = 3.82; Y: 900,000/180,000 = 5.00.

   ```
   X 3.82 · Y 5.00
   ```

4. Compute Interest Coverage = EBIT / Interest Expense. X: 100,000/10,000 = 10.0; Y: 228,000/18,000 = 12.7.

   ```
   X 10.0 · Y 12.7
   ```

5. Weigh the evidence: Y carries more debt per dollar of equity and higher leverage, but also earns more cover for its interest bill.
6. Write the verdict: Y has the higher structural solvency risk (D/E 2.78, leverage 5.00) even though its interest coverage is currently stronger.

**Test it**

Your table matches the model answers (D/E 1.76 vs 2.78; leverage 3.82 vs 5.00; coverage 10.0 vs 12.7) and your verdict identifies Company Y as the riskier capital structure.

> **Note:** The full worksheet for this activity is available on the LMS, and a printable copy is in activities/activity-06 of the course repository.

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
