# Activity 6 — Solvency & Financial Risk Analysis

**Topic 3** · LO3 — compare companies' solvency risk from their statements (K5, A2)

**Goal:** Return to Orchid Logistics (Company X) and Marina Retail Group (Company Y). Using their balance sheets and income statements, compare the two companies' solvency risk — debt-to-equity, financial leverage and interest coverage — and conclude which is the riskier borrower.

**You'll produce:** A solvency comparison table (D/E, leverage, interest coverage) for X and Y with a reasoned risk verdict.

**Tools:** Microsoft Excel · data workbook FA-Activity-06-Solvency-Analysis.xlsx in activities/activity06

## Data provided

- **Workbook:** FA-Activity-06-Solvency-Analysis.xlsx — same X/Y statements as Activity 2 ('Balance Sheet' and 'Income Statement', X in column B, Y in column C) plus a solvency 'Your Workings' template
- **Balance Sheet (X | Y):** Total assets 650,000 | 900,000 · LT liabilities 300,000 | 500,000 · Capital 170,000 | 180,000
- **Income Statement (X | Y):** EBIT 100,000 | 228,000 · Interest 10,000 | 18,000

## Step-by-step

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

## Data files in this folder

- `FA-Activity-06-Solvency-Analysis.xlsx`
- `balance_sheet_x_y.csv`
- `income_statement_x_y.csv`

## Test it

Your table matches the model answers (D/E 1.76 vs 2.78; leverage 3.82 vs 5.00; coverage 10.0× vs 12.7×) and your verdict identifies Company Y's capital structure as the riskier one, with reasons.
