# Activity 5 — Net Present Value & Profitability Index

**Topic 3** · LO3 — evaluate a capital investment using NPV and PI (K5, A2)

**Goal:** Using the same project ($2,000 outlay, $375/year for 10 years, 10% required return), compute the Net Present Value and the Profitability Index, and decide whether to accept the project.

**You'll produce:** An NPV/PI worksheet: PV of each inflow, NPV ≈ $304, PI ≈ 1.15, and an accept/reject decision.

**Tools:** Microsoft Excel (or Google Sheets), activity worksheet from the LMS

## Data provided

- **Cash flows:** Year 0: −$2,000; Years 1–10: +$375 per year
- **Required return:** 10%

## Step-by-step

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

## Test it

Your worksheet shows NPV ≈ $304 and PI ≈ 1.15, and your decision is ACCEPT because NPV > 0 and PI > 1.
