"""Topic 3 — Planning & Budgeting using Financial Statements: hands-on activities (mock-data workbooks)."""

DOMAIN3 = [
    dict(
        num=4,
        topic=3,
        title="Payback Period & Discounted Payback Period",
        objective="LO3 — evaluate a capital investment using payback methods (K5, A2)",
        desc="Havenwood Trading is evaluating a refrigerated delivery van: S$2,000k outlay in Year 0, saving "
             "S$375k per year for 10 years versus outsourced cold-chain delivery. Compute the simple payback "
             "period, then discount each cash flow at the 10% required return to find the discounted payback.",
        build="A payback worksheet showing cumulative cash flow per year, the simple payback point (5.33 years) "
              "and the discounted payback point (≈8 years).",
        services="Microsoft Excel · data workbook FA-Activity-04-Payback-Period.xlsx in activities/activity04",
        data=[
            ("Workbook", "FA-Activity-04-Payback-Period.xlsx — sheet 'Project Cash Flows': Year in column A "
             "(rows 6–16), net cash flow in column B; columns C–E are your templates (cumulative, discounted, "
             "cumulative discounted)"),
            ("Cash flows", "Year 0: −S$2,000k · Years 1–10: +S$375k per year (all figures S$'000)"),
            ("Discount rate", "10% — Havenwood's required return on internal projects"),
        ],
        steps=[
            ("Open FA-Activity-04-Payback-Period.xlsx, sheet 'Project Cash Flows'. Year 0 (the −2,000 outlay) is "
             "row 6; Years 1–10 with +375 each are rows 7–16.", ""),
            ("Build the cumulative cash flow in column C: start at C7 with =B7, then C8 =C7+B8 and fill down to "
             "C16. Watch where the running total passes 2,000.", "C7 =B7 → 375 · C8 =C7+B8 → 750 … C12 → 2,250 (passes 2,000 in Year 6)"),
            ("Compute the simple payback period exactly: full years until the outlay is nearly recovered (5 years "
             "→ 1,875), plus the fraction of Year 6 needed: (2,000 − 1,875)/375.", "Payback = 5 + 125/375 = 5.33 years  (= 2000/375 for an even stream)"),
            ("Now build the discounted column D: each year's cash flow divided by (1.10)^year. Enter "
             "=B7/(1.1)^A7 in D7 and fill down.", "D7 = 375/1.1 = 340.91 · D8 = 309.92 · D9 = 281.74 · D10 = 256.13 · D11 = 232.85"),
            ("Build the cumulative discounted column E the same way as column C and read down for the 2,000 "
             "crossing: 340.91, 650.83, 932.57, 1,188.70, 1,421.55, 1,633.22, 1,825.66, 2,000.60 — it crosses "
             "during Year 8.", "E14 (Year 8) = 2,000.60 ≥ 2,000 → discounted payback ≈ 8.0 years"),
            ("Interpret the gap on the 'Your Workings' sheet: discounting stretches the payback from 5.33 to ≈8 "
             "years because later savings are worth less today. State when payback is the right tool (quick "
             "liquidity screen) and its blind spots (ignores cash flows AFTER payback, and simple payback "
             "ignores the time value of money).", ""),
        ],
        test="Your cumulative discounted cash flow reaches S$2,000.60k at Year 8 — the discounted payback "
             "(≈8 years) is materially longer than the simple payback (5.33 years), and you can say why.",
    ),
    dict(
        num=5,
        topic=3,
        title="Net Present Value & Profitability Index",
        objective="LO3 — evaluate a capital investment using NPV and PI (K5, A2)",
        desc="Using the same delivery-van proposal (S$2,000k outlay, S$375k/year for 10 years, 10% required "
             "return), compute the Net Present Value and the Profitability Index, and make the accept/reject "
             "recommendation to management.",
        build="An NPV/PI worksheet: PV of each inflow, NPV ≈ S$304k, PI ≈ 1.15, and a reasoned accept decision.",
        services="Microsoft Excel · data workbook FA-Activity-05-NPV-PI.xlsx in activities/activity05",
        data=[
            ("Workbook", "FA-Activity-05-NPV-PI.xlsx — sheet 'Project Cash Flows': Year in column A (rows 6–16), "
             "net cash flow in column B, discounted-cash-flow template in column C; 'Your Workings' for NPV/PI"),
            ("Cash flows", "Year 0: −S$2,000k · Years 1–10: +S$375k per year (S$'000)"),
            ("Required return", "10%"),
        ],
        steps=[
            ("Open FA-Activity-05-NPV-PI.xlsx, sheet 'Project Cash Flows'. Discount each inflow in column C with "
             "=B7/(1.1)^A7 filled down rows 7–16.", "341, 310, 282, 256, 233, 212, 192, 175, 159, 145  (S$'000, rounded)"),
            ("Sum the discounted inflows to get the present value of what the van gives back.", "PV of inflows =SUM(C7:C16) = 2,304"),
            ("Compute NPV = PV of inflows − outlay. Cross-check with Excel's NPV function — note it discounts "
             "from Year 1, so the Year-0 outlay stays outside it.", "NPV = 2,304 - 2,000 = 304   (=NPV(10%,B7:B16)+B6)"),
            ("Compute the Profitability Index = PV of inflows / PV of outflows — the value created per dollar "
             "invested.", "PI = 2,304 / 2,000 = 1.15"),
            ("Make the decision on 'Your Workings': NPV > 0 AND PI > 1 → ACCEPT. Write the recommendation as "
             "you would to management: 'the van creates S$304k of value at our 10% hurdle; every dollar "
             "invested returns S$1.15 of present value'.", "Accept: NPV 304 > 0 · PI 1.15 > 1"),
            ("Stress-test the recommendation: raise the discount rate in your formulas to 14% and observe NPV "
             "turn slightly negative — the project's IRR is ≈13.4%, so the accept decision holds only while "
             "the cost of capital stays below that. Note this sensitivity in one line.", "At i = 14%: PV inflows ≈ 1,956 → NPV ≈ -44 → reject"),
        ],
        test="Your worksheet shows NPV ≈ S$304k and PI ≈ 1.15 with an ACCEPT decision, plus a one-line "
             "sensitivity note (NPV turns negative near a 13–14% discount rate).",
    ),
    dict(
        num=6,
        topic=3,
        title="Solvency & Financial Risk Analysis",
        objective="LO3 — compare companies' solvency risk from their statements (K5, A2)",
        desc="Return to Orchid Logistics (Company X) and Marina Retail Group (Company Y). Using their balance "
             "sheets and income statements, compare the two companies' solvency risk — debt-to-equity, "
             "financial leverage and interest coverage — and conclude which is the riskier borrower.",
        build="A solvency comparison table (D/E, leverage, interest coverage) for X and Y with a reasoned "
              "risk verdict.",
        services="Microsoft Excel · data workbook FA-Activity-06-Solvency-Analysis.xlsx in activities/activity06",
        data=[
            ("Workbook", "FA-Activity-06-Solvency-Analysis.xlsx — same X/Y statements as Activity 2 "
             "('Balance Sheet' and 'Income Statement', X in column B, Y in column C) plus a solvency "
             "'Your Workings' template"),
            ("Balance Sheet (X | Y)", "Total assets 650,000 | 900,000 · LT liabilities 300,000 | 500,000 · "
             "Capital 170,000 | 180,000"),
            ("Income Statement (X | Y)", "EBIT 100,000 | 228,000 · Interest 10,000 | 18,000"),
        ],
        steps=[
            ("Open FA-Activity-06-Solvency-Analysis.xlsx. Solvency looks at the LONG-term survival of the "
             "capital structure, so pull three ingredients per company from the 'Balance Sheet' sheet: "
             "long-term liabilities (row 17), capital (row 18) and total assets (row 13); and two from the "
             "'Income Statement': EBIT (row 12) and interest (row 13).", ""),
            ("Compute Debt-to-Equity = Long-Term Liabilities / Equity for both companies — how many dollars of "
             "long-term debt sit on each dollar the owners have at stake.", "X: 300,000/170,000 = 1.76 · Y: 500,000/180,000 = 2.78"),
            ("Compute Financial Leverage = Total Assets / Equity — the equity multiplier. At 5.00, only a fifth "
             "of Y's balance sheet is owner-funded.", "X: 650,000/170,000 = 3.82 · Y: 900,000/180,000 = 5.00"),
            ("Compute Interest Coverage = EBIT / Interest Expense — how many times operating profit covers the "
             "interest bill this year.", "X: 100,000/10,000 = 10.0× · Y: 228,000/18,000 = 12.7×"),
            ("Weigh the evidence like a credit officer: Y carries structurally more debt per dollar of equity "
             "(D/E 2.78, leverage 5.00) but currently earns stronger cover (12.7×). Ask the trend question: "
             "what happens to Y's cover if EBIT halves in a downturn? X: 5.0× — comfortable; Y: 6.3× but on a "
             "far bigger debt load that must be refinanced.", "Stress: halve EBIT → X 5.0× · Y 6.3×, with 500,000 of LT debt to refinance"),
            ("Write the verdict on 'Your Workings': Company Y has the higher structural solvency risk — its "
             "capital structure depends on lenders — even though its current interest coverage is stronger. "
             "Recommend the leverage and coverage covenants you would monitor.", ""),
        ],
        test="Your table matches the model answers (D/E 1.76 vs 2.78; leverage 3.82 vs 5.00; coverage 10.0× vs "
             "12.7×) and your verdict identifies Company Y's capital structure as the riskier one, with reasons.",
    ),
]
