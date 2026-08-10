# Activity 4 — Payback Period & Discounted Payback Period

**Topic 3** · LO3 — evaluate a capital investment using payback methods (K5, A2)

**Goal:** A project requires an initial outlay of $2,000 and returns $375 per year for 10 years. Compute the simple payback period, then repeat the exercise discounting each cash flow at 10% to find the discounted payback period.

**You'll produce:** A payback worksheet showing cumulative cash flow per year, the simple payback point (~5.3 years) and the discounted payback point (~8 years).

**Tools:** Microsoft Excel (or Google Sheets), activity worksheet from the LMS

## Data provided

- **Cash flows:** Year 0: −$2,000; Years 1–10: +$375 per year
- **Discount rate:** 10% for the discounted payback

## Step-by-step

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


## Test it

Your cumulative discounted cash flow reaches $2,000.60 at Year 8 — the discounted payback (~8 years) is materially longer than the simple payback (5.33 years).
