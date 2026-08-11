"""Topic 2 — Analysing Financial Ratios: hands-on activities (mock-data workbooks)."""

DOMAIN2 = [
    dict(
        num=2,
        topic=2,
        title="Ratio Analysis of Two Companies",
        company="Orchid Logistics (X) vs Marina Retail Group (Y)",
        objective="LO2 — evaluate financial performance from financial ratios (K4, A1)",
        desc="Orchid Logistics Pte Ltd (Company X) and Marina Retail Group Pte Ltd (Company Y) are two "
             "fictitious SMEs. From their FY2025 balance sheets and income statements, compute the liquidity, "
             "profitability, turnover and solvency ratios for both companies and give a one-line verdict per "
             "ratio family on which company performs better.",
        build="A completed ratio-analysis worksheet comparing Company X and Company Y across four ratio "
              "families, with a one-line verdict per ratio.",
        services="Microsoft Excel · data workbook FA-Activity-02-Ratio-Analysis.xlsx in activities/activity02",
        data=[
            ("Workbook", "FA-Activity-02-Ratio-Analysis.xlsx — sheets: 'Balance Sheet' (X in column B, Y in "
             "column C), 'Income Statement' (X in B, Y in C), 'Your Workings' (ratio template)"),
            ("Balance Sheet (X | Y)", "Cash 50,000 | 80,000 · Inventory 40,000 | 70,000 · Receivables 100,000 | 250,000 · "
             "Property 400,000 | 450,000 · Other FA 60,000 | 50,000 · Total assets 650,000 | 900,000"),
            ("Liabilities & equity (X | Y)", "Payables 100,000 | 150,000 · ST loans 80,000 | 70,000 · LT liabilities "
             "300,000 | 500,000 · Capital 170,000 | 180,000"),
            ("Income Statement (X | Y)", "Sales 300,000 | 500,000 · COGS 140,000 | 200,000 · Gross profit 160,000 | 300,000 · "
             "EBIT 100,000 | 228,000 · Interest 10,000 | 18,000 · Net income 72,000 | 168,000"),
        ],
        steps=[
            ("Open FA-Activity-02-Ratio-Analysis.xlsx. On the 'Balance Sheet' sheet, Company X is column B and "
             "Company Y is column C. First aggregate the raw ingredients on the 'Your Workings' sheet: "
             "Current assets = Cash + Inventory + Receivables (rows 7–9); Current liabilities = Payables + "
             "ST loans (rows 15–16).", "CA(X) ='Balance Sheet'!B7+B8+B9 = 190,000 · CL(X) = B15+B16 = 180,000\nCA(Y) = 400,000 · CL(Y) = 220,000"),
            ("LIQUIDITY — Current Ratio = Current Assets / Current Liabilities. Compute for both companies and "
             "interpret against the ≥1 benchmark from the slides.", "X: 190,000/180,000 = 1.06 · Y: 400,000/220,000 = 1.82 → Y is more liquid"),
            ("LIQUIDITY — Quick Ratio = (Current Assets − Inventory) / Current Liabilities. Inventory is the "
             "least liquid current asset, so strip it out and recompute.", "X: 150,000/180,000 = 0.83 · Y: 330,000/220,000 = 1.50"),
            ("PROFITABILITY — Operating Profit Ratio = EBIT / Sales, using 'Income Statement' row 12 (EBIT) over "
             "row 6 (Sales). Express as a percentage.", "X: 100,000/300,000 = 33% · Y: 228,000/500,000 = 46%"),
            ("PROFITABILITY — Net Profit Ratio = Net Income / Sales (row 16 over row 6). Note how the gap versus "
             "the operating margin reveals the drag from interest and tax.", "X: 72,000/300,000 = 24% · Y: 168,000/500,000 = 34%"),
            ("PROFITABILITY — Return on Capital Employed = EBIT / Capital Employed, where Capital Employed = "
             "Equity + Long-Term Liabilities (the long-term funding of the business).", "Cap employed X: 170,000+300,000 = 470,000 → 100,000/470,000 = 21%\nCap employed Y: 180,000+500,000 = 680,000 → 228,000/680,000 = 34%"),
            ("TURNOVER — Inventory Turnover = COGS / Average Inventory (use the year-end balance as the average). "
             "A higher figure means stock moves faster.", "X: 140,000/40,000 = 3.50 · Y: 200,000/70,000 = 2.86 → X turns stock faster"),
            ("TURNOVER — Receivables Turnover = Credit Sales / Average Receivables (treat all sales as credit "
             "sales). Then convert to days: Average Collection Period = 365 / turnover.", "X: 300,000/100,000 = 3.00 → 122 days · Y: 500,000/250,000 = 2.00 → 183 days"),
            ("TURNOVER — Payables Turnover = Purchases / Average Payables, assuming purchases = 50% of COGS.", "X: 70,000/100,000 = 0.70 · Y: 100,000/150,000 = 0.67"),
            ("SOLVENCY — Debt-to-Equity = Long-Term Liabilities / Equity, and Financial Leverage = Total Assets "
             "/ Equity. Higher values = more of the balance sheet is funded by lenders.", "D/E — X: 300,000/170,000 = 1.76 · Y: 500,000/180,000 = 2.78\nLeverage — X: 650,000/170,000 = 3.82 · Y: 900,000/180,000 = 5.00"),
            ("Write the verdicts in the 'Your Workings' sheet, one line per family: Y wins on liquidity and "
             "profitability; X wins on turnover efficiency and carries materially less solvency risk. Note how "
             "Y's higher ROE partly comes from leverage, not just operations — the DuPont insight.", ""),
        ],
        test="Your worksheet shows all ten ratios for both companies matching the model answers (e.g. Current "
             "Ratio X 1.06 vs Y 1.82; D/E X 1.76 vs Y 2.78), with a one-line verdict per ratio family.",
    ),
    dict(
        num=3,
        topic=2,
        title="Trend Analysis over Four Years",
        company="Sunrise F&B Holdings Pte Ltd (FY2018–FY2021)",
        objective="LO2 — identify trends by comparing ratios across time periods (K5, A1)",
        desc="Sunrise F&B Holdings Pte Ltd is a fictitious cafe-chain SME. From four years (FY2018–FY2021) of "
             "balance sheets and income statements, compute year-on-year growth for every line, derive the key "
             "ratios for each year, and present your observations on where the business is heading.",
        build="A completed trend-analysis template: YoY growth for every balance-sheet and income-statement "
              "line plus a four-year ratio table with your observations.",
        services="Microsoft Excel · data workbook FA-Activity-03-Trend-Analysis.xlsx in activities/activity03",
        data=[
            ("Workbook", "FA-Activity-03-Trend-Analysis.xlsx — sheets: 'Balance Sheet 2018-2021' and 'Income "
             "Statement 2018-2021' (FY2021 in column B through FY2018 in column E), 'Trend Template'"),
            ("Sales FY2018→FY2021", "180,000 → 198,000 → 217,800 → 239,580 (10% p.a.)"),
            ("Net income FY2018→FY2021", "53,600 → 64,601 → 76,872 → 90,548"),
            ("Total assets FY2018→FY2021", "330,000 → 324,800 → 329,983 → 339,581"),
        ],
        steps=[
            ("Open FA-Activity-03-Trend-Analysis.xlsx. Years run LEFT to RIGHT from newest: FY2021 (column B) to "
             "FY2018 (column E). Skim both statement sheets and note what is obviously moving: sales up every "
             "year, payables up sharply, property drifting down.", ""),
            ("Compute YoY growth for each income-statement line: growth = (this year − last year) / last year. "
             "In Excel, put the formula beside FY2021 and fill across.", "=(B6-C6)/C6 → Sales FY2021 growth = (239,580-217,800)/217,800 = 10.0%"),
            ("Read the growth pattern: Sales grow a steady 10% while COGS grows only 5% — so Gross Profit growth "
             "runs ahead of sales (14.5% → 13.8%) and EBIT growth is faster still. Fixed S&D expenses (flat "
             "20,000) create operating leverage.", "COGS: =(B7-C7)/C7 = 5.0% each year"),
            ("Repeat for the balance sheet lines. Flag the working-capital lines: Accounts payable grew 25.0% in "
             "FY2021 (75,000 vs 60,000) while inventory grew 17.6% — the company is stretching its suppliers.", "=(B15-C15)/C15 = 25.0%"),
            ("Now derive the LIQUIDITY ratios per year in the 'Trend Template': Current Ratio = (Cash + Inventory "
             "+ Receivables) / (Payables + ST loans); Acid Ratio strips inventory.", "FY2021: 155,583/95,000 = 1.64 · FY2018: 145,000/68,000 = 2.13\nAcid: FY2021 1.22 · FY2018 1.47"),
            ("Derive the EFFICIENCY ratios per year: Asset Turnover = Sales/Total assets; Inventory Turnover = "
             "COGS/Inventory; Receivables Turnover = Sales/Receivables; Payables Turnover = Purchases/Payables "
             "with purchases = 50% of COGS.", "Asset turnover: 0.55 → 0.72 · Inventory turnover: 4.00 → 6.48\nReceivables: 3.00 → 3.29 · Payables: 0.99 → 0.73"),
            ("Derive the LEVERAGE ratios per year: Debt-to-Equity = (ST loans + LT liabilities)/Capital; "
             "Debt-to-Assets = (ST loans + LT liabilities + Payables)/Total assets.", "D/E: 0.56 (FY2018) → 0.75 (FY2021) · D/A: 0.36 → 0.43"),
            ("Derive the PROFITABILITY ratios per year: Operating Margin = EBIT/Sales; Net Profit Margin = Net "
             "income/Sales.", "Operating margin: 0.40 → 0.49 · Net margin: 0.30 → 0.38"),
            ("Write your observations under the template — cover all four families: (1) profitability and "
             "efficiency improve steadily (margins, asset and inventory turnover all up); (2) liquidity "
             "deteriorates (current ratio 2.13 → 1.64) as payables are stretched; (3) leverage creeps up "
             "(D/E 0.56 → 0.75); (4) verdict: a profitable, tightening business — watch the falling current "
             "ratio and the supplier stretch.", ""),
        ],
        test="Your ratio table matches the model answers (e.g. FY2021 Current Ratio 1.64, Net Profit Margin "
             "0.38, D/E 0.75) and your observations cover profitability, efficiency, liquidity and leverage.",
    ),
]
