"""Topic 3 — Planning & Budgeting using Financial Statements: hands-on activities."""

DOMAIN3 = [
    dict(
        num=4,
        topic=3,
        title="Payback Period & Discounted Payback Period",
        objective="LO3 — evaluate a capital investment using payback methods (K5, A2)",
        desc="A project requires an initial outlay of $2,000 and returns $375 per year for 10 years. "
             "Compute the simple payback period, then repeat the exercise discounting each cash flow at "
             "10% to find the discounted payback period.",
        build="A payback worksheet showing cumulative cash flow per year, the simple payback point "
              "(~5.3 years) and the discounted payback point (~8 years).",
        services="Microsoft Excel (or Google Sheets), activity worksheet from the LMS",
        data=[
            ("Cash flows", "Year 0: −$2,000; Years 1–10: +$375 per year"),
            ("Discount rate", "10% for the discounted payback"),
        ],
        steps=[
            ("Download the activity worksheet 'Cash Flow Payback' from the LMS and open it in Excel.", ""),
            ("Build the cumulative cash flow column: 375, 750, 1,125, 1,500, 1,875, 2,250 … the cumulative figure passes $2,000 during Year 6.", "=SUM($B$2:B7)"),
            ("Compute the simple payback period: 2,000 / 375 = 5.33 years (about 5 years 4 months).", "Payback = 2000/375 = 5.33 yr"),
            ("Add a discounted cash flow column: DCF = 375 / (1.10)^n → 340.91, 309.92, 281.74, 256.13, 232.85, 211.68, 192.43, 174.94 …", "=375/(1.1)^A2"),
            ("Build the cumulative discounted column: 340.91, 650.83, 932.57, 1,188.70, 1,421.55, 1,633.22, 1,825.66, 2,000.60 — it crosses $2,000 in Year 8.", ""),
            ("State the discounted payback period: ≈ 8 years, and explain why discounting lengthens the payback.", "Discounted payback ≈ 8.0 yr"),
        ],
        test="Your cumulative discounted cash flow reaches $2,000.60 at Year 8 — the discounted payback "
             "(~8 years) is materially longer than the simple payback (5.33 years).",
    ),
    dict(
        num=5,
        topic=3,
        title="Net Present Value & Profitability Index",
        objective="LO3 — evaluate a capital investment using NPV and PI (K5, A2)",
        desc="Using the same project ($2,000 outlay, $375/year for 10 years, 10% required return), "
             "compute the Net Present Value and the Profitability Index, and decide whether to accept "
             "the project.",
        build="An NPV/PI worksheet: PV of each inflow, NPV ≈ $304, PI ≈ 1.15, and an accept/reject decision.",
        services="Microsoft Excel (or Google Sheets), activity worksheet from the LMS",
        data=[
            ("Cash flows", "Year 0: −$2,000; Years 1–10: +$375 per year"),
            ("Required return", "10%"),
        ],
        steps=[
            ("Download the activity worksheet 'Discounted Cash Flow' from the LMS and open it in Excel.", ""),
            ("Discount each inflow at 10%: 341, 310, 282, 256, 233, 212, 192, 175, 159, 145.", "=375/(1.1)^A2"),
            ("Sum the discounted inflows: PV of cash inflows = $2,304.", "PV inflows = 2,304"),
            ("Compute NPV = PV of inflows − initial outlay = 2,304 − 2,000 = $304 (or use =NPV(10%, range) − 2000).", "NPV = 304"),
            ("Compute the Profitability Index = PV of inflows / PV of outflows = 2,304 / 2,000 = 1.15.", "PI = 1.15"),
            ("Decide: NPV > 0 and PI > 1 → accept the project; note how the decision rule would flip if the discount rate rose.", ""),
        ],
        test="Your worksheet shows NPV ≈ $304 and PI ≈ 1.15, and your decision is ACCEPT because NPV > 0 "
             "and PI > 1.",
    ),
    dict(
        num=6,
        topic=3,
        title="Solvency & Financial Risk Analysis",
        objective="LO3 — compare companies' solvency risk from their statements (K5, A2)",
        desc="Using the balance sheets and income statements of companies X and Y, compare the two "
             "companies in terms of solvency risk: debt-to-equity, financial leverage and interest "
             "coverage. Conclude which company is the riskier borrower.",
        build="A solvency comparison table (D/E, leverage, interest coverage) for X and Y with a "
              "risk verdict.",
        services="Microsoft Excel (or Google Sheets), activity worksheet from the LMS",
        data=[
            ("Balance Sheet (X | Y)", "Total Assets 650,000 | 900,000 · LT Debt 300,000 | 500,000 · Equity 170,000 | 180,000"),
            ("Income Statement (X | Y)", "EBIT 100,000 | 228,000 · Interest 10,000 | 18,000"),
        ],
        steps=[
            ("Download the activity worksheet 'Solvency Analysis' from the LMS and open it in Excel.", ""),
            ("Compute Debt-to-Equity = Long-Term Debt / Equity. X: 300,000/170,000 = 1.76; Y: 500,000/180,000 = 2.78.", "X 1.76 · Y 2.78"),
            ("Compute Financial Leverage = Total Assets / Equity. X: 650,000/170,000 = 3.82; Y: 900,000/180,000 = 5.00.", "X 3.82 · Y 5.00"),
            ("Compute Interest Coverage = EBIT / Interest Expense. X: 100,000/10,000 = 10.0; Y: 228,000/18,000 = 12.7.", "X 10.0 · Y 12.7"),
            ("Weigh the evidence: Y carries more debt per dollar of equity and higher leverage, but also earns more cover for its interest bill.", ""),
            ("Write the verdict: Y has the higher structural solvency risk (D/E 2.78, leverage 5.00) even though its interest coverage is currently stronger.", ""),
        ],
        test="Your table matches the model answers (D/E 1.76 vs 2.78; leverage 3.82 vs 5.00; coverage 10.0 "
             "vs 12.7) and your verdict identifies Company Y as the riskier capital structure.",
    ),
]
