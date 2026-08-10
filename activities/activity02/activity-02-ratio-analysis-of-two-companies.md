# Activity 2 — Ratio Analysis of Two Companies

**Topic 2** · LO2 — evaluate financial performance from financial ratios (K4, A1)

**Goal:** Orchid Logistics Pte Ltd (Company X) and Marina Retail Group Pte Ltd (Company Y) are two fictitious SMEs. From their FY2025 balance sheets and income statements, compute the liquidity, profitability, turnover and solvency ratios for both companies and give a one-line verdict per ratio family on which company performs better.

**You'll produce:** A completed ratio-analysis worksheet comparing Company X and Company Y across four ratio families, with a one-line verdict per ratio.

**Tools:** Microsoft Excel · data workbook FA-Activity-02-Ratio-Analysis.xlsx in activities/activity02

## Data provided

- **Workbook:** FA-Activity-02-Ratio-Analysis.xlsx — sheets: 'Balance Sheet' (X in column B, Y in column C), 'Income Statement' (X in B, Y in C), 'Your Workings' (ratio template)
- **Balance Sheet (X | Y):** Cash 50,000 | 80,000 · Inventory 40,000 | 70,000 · Receivables 100,000 | 250,000 · Property 400,000 | 450,000 · Other FA 60,000 | 50,000 · Total assets 650,000 | 900,000
- **Liabilities & equity (X | Y):** Payables 100,000 | 150,000 · ST loans 80,000 | 70,000 · LT liabilities 300,000 | 500,000 · Capital 170,000 | 180,000
- **Income Statement (X | Y):** Sales 300,000 | 500,000 · COGS 140,000 | 200,000 · Gross profit 160,000 | 300,000 · EBIT 100,000 | 228,000 · Interest 10,000 | 18,000 · Net income 72,000 | 168,000

## Step-by-step

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

## Data files in this folder

- `FA-Activity-02-Ratio-Analysis.xlsx`
- `balance_sheet_x_y.csv`
- `income_statement_x_y.csv`

## Test it

Your worksheet shows all ten ratios for both companies matching the model answers (e.g. Current Ratio X 1.06 vs Y 1.82; D/E X 1.76 vs Y 2.78), with a one-line verdict per ratio family.
