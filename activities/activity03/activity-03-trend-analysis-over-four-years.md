# Activity 3 — Trend Analysis over Four Years

**Topic 2** · LO2 — identify trends by comparing ratios across time periods (K5, A1)

**Goal:** Sunrise F&B Holdings Pte Ltd is a fictitious cafe-chain SME. From four years (FY2018–FY2021) of balance sheets and income statements, compute year-on-year growth for every line, derive the key ratios for each year, and present your observations on where the business is heading.

**You'll produce:** A completed trend-analysis template: YoY growth for every balance-sheet and income-statement line plus a four-year ratio table with your observations.

**Tools:** Microsoft Excel · data workbook FA-Activity-03-Trend-Analysis.xlsx in activities/activity03

## Data provided

- **Workbook:** FA-Activity-03-Trend-Analysis.xlsx — sheets: 'Balance Sheet 2018-2021' and 'Income Statement 2018-2021' (FY2021 in column B through FY2018 in column E), 'Trend Template'
- **Sales FY2018→FY2021:** 180,000 → 198,000 → 217,800 → 239,580 (10% p.a.)
- **Net income FY2018→FY2021:** 53,600 → 64,601 → 76,872 → 90,548
- **Total assets FY2018→FY2021:** 330,000 → 324,800 → 329,983 → 339,581

## Step-by-step

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

## Data files in this folder

- `FA-Activity-03-Trend-Analysis.xlsx`
- `balance_sheet_2018_2021.csv`
- `income_statement_2018_2021.csv`

## Test it

Your ratio table matches the model answers (e.g. FY2021 Current Ratio 1.64, Net Profit Margin 0.38, D/E 0.75) and your observations cover profitability, efficiency, liquidity and leverage.
