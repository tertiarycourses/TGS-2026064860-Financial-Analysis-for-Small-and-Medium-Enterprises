# Activity 5 — Net Present Value & Profitability Index

**Topic 3** · LO3 — evaluate a capital investment using NPV and PI (K5, A2)

**Goal:** Using the same delivery-van proposal (S$2,000k outlay, S$375k/year for 10 years, 10% required return), compute the Net Present Value and the Profitability Index, and make the accept/reject recommendation to management.

**You'll produce:** An NPV/PI worksheet: PV of each inflow, NPV ≈ S$304k, PI ≈ 1.15, and a reasoned accept decision.

**Tools:** Microsoft Excel · data workbook FA-Activity-05-NPV-PI.xlsx in activities/activity05

## Data provided

- **Workbook:** FA-Activity-05-NPV-PI.xlsx — sheet 'Project Cash Flows': Year in column A (rows 6–16), net cash flow in column B, discounted-cash-flow template in column C; 'Your Workings' for NPV/PI
- **Cash flows:** Year 0: −S$2,000k · Years 1–10: +S$375k per year (S$'000)
- **Required return:** 10%

## Step-by-step

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


## Files in this folder

- `FA-Activity-05-NPV-PI.xlsx` — mock-data workbook (Excel)
- `activity-05-worksheet-question.pdf` — printable worksheet (PDF)
- `project_cash_flows.csv` — raw data (CSV)

> The model-answer PDF (`*-worksheet-answer.pdf`) in this folder is trainer material.


## Test it

Your worksheet shows NPV ≈ S$304k and PI ≈ 1.15 with an ACCEPT decision, plus a one-line sensitivity note (NPV turns negative near a 13–14% discount rate).
