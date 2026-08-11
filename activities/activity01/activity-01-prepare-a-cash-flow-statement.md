# Activity 1 — Prepare a Cash Flow Statement

**Topic 1** · LO1 — understand the balance sheet, income and cash flow statements (K1, K2, K3)

**Goal:** Havenwood Trading Pte Ltd is a fictitious Singapore SME distributor. From its FY2018/FY2019 balance sheets and FY2019 income statement (all in S$ million), derive the full statement of cash flows — operating, investing and financing — and reconcile it to the movement in cash.

**You'll produce:** A complete statement of cash flows (operating / investing / financing) for Havenwood Trading, derived from the balance sheet movements and the income statement.

**Tools:** Microsoft Excel · data workbook FA-Activity-01-Cash-Flow-Statement.xlsx in activities/activity01

## Data provided

- **Workbook:** FA-Activity-01-Cash-Flow-Statement.xlsx — sheets: 'Balance Sheet' (FY2019 in column B, FY2018 in column C), 'Income Statement' (FY2019 in column B), 'Your Workings' (answer template)
- **Balance Sheet (FY2019 | FY2018):** PP&E 100 | 90 · Receivables 50 | 50 · Inventory 30 | 30 · Cash 20 | 20 · Equity 80 | 80 · Bank loans 50 | 55 · Payables 40 | 35 · Provisions 30 | 20 (S$m)
- **Income Statement FY2019:** Revenue 50 (wholesale 38 + retail 12) · COGS 10 · Operational costs 20 (salaries 12, rental 4, utilities/logistics 4) · Depreciation 5 · Interest 3 · Tax 2 · Net income 10 (S$m)

## Step-by-step

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

## Files in this folder

- `FA-Activity-01-Cash-Flow-Statement.xlsx` — mock-data workbook (Excel)
- `activity-01-worksheet-question.pdf` — printable worksheet (PDF)
- `balance_sheet.csv` — raw data (CSV)
- `income_statement.csv` — raw data (CSV)

> The model-answer PDF (`*-worksheet-answer.pdf`) in this folder is trainer material.


## Test it

Your three blocks total +18 (operating), −14 (investing) and −14 (financing), net −10 — matching the model answer — and you can explain in one sentence why a profitable year still drained cash.
