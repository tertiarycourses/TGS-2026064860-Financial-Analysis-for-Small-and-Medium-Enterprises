# Activity 4 — Payback Period & Discounted Payback Period

**Topic 3** · LO3 — evaluate a capital investment using payback methods (K5, A2)

**Goal:** Havenwood Trading is evaluating a refrigerated delivery van: S$2,000k outlay in Year 0, saving S$375k per year for 10 years versus outsourced cold-chain delivery. Compute the simple payback period, then discount each cash flow at the 10% required return to find the discounted payback.

**You'll produce:** A payback worksheet showing cumulative cash flow per year, the simple payback point (5.33 years) and the discounted payback point (≈8 years).

**Tools:** Microsoft Excel · data workbook FA-Activity-04-Payback-Period.xlsx in activities/activity04

## Data provided

- **Workbook:** FA-Activity-04-Payback-Period.xlsx — sheet 'Project Cash Flows': Year in column A (rows 6–16), net cash flow in column B; columns C–E are your templates (cumulative, discounted, cumulative discounted)
- **Cash flows:** Year 0: −S$2,000k · Years 1–10: +S$375k per year (all figures S$'000)
- **Discount rate:** 10% — Havenwood's required return on internal projects

## Step-by-step

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

## Files in this folder

- `FA-Activity-04-Payback-Period.xlsx` — mock-data workbook (Excel)
- `activity-04-worksheet-question.pdf` — printable worksheet (PDF)
- `project_cash_flows.csv` — raw data (CSV)

> The model-answer PDF (`*-worksheet-answer.pdf`) in this folder is trainer material.


## Test it

Your cumulative discounted cash flow reaches S$2,000.60k at Year 8 — the discounted payback (≈8 years) is materially longer than the simple payback (5.33 years), and you can say why.
