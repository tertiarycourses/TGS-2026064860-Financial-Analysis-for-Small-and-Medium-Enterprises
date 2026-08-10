"""Topic 2 — Analysing Financial Ratios: hands-on activities."""

DOMAIN2 = [
    dict(
        num=2,
        topic=2,
        title="Ratio Analysis of Two Companies",
        objective="LO2 — evaluate financial performance from financial ratios (K4, A1)",
        desc="You are a financial analyst evaluating two companies, X and Y, from their balance sheets "
             "and income statements. Compute the liquidity, profitability, turnover and solvency ratios "
             "for both companies and comment on which company performs better on each parameter.",
        build="A completed ratio-analysis worksheet comparing Company X and Company Y across four ratio "
              "families, with a one-line verdict per ratio.",
        services="Microsoft Excel (or Google Sheets), activity worksheet from the LMS",
        data=[
            ("Balance Sheet (X | Y)", "Cash 50,000 | 80,000 · Inventory 40,000 | 70,000 · AR 100,000 | 250,000 · "
             "Property 400,000 | 450,000 · Other FA 60,000 | 50,000 · Total 650,000 | 900,000"),
            ("Liabilities (X | Y)", "AP 100,000 | 150,000 · ST Loans 80,000 | 70,000 · LT Liabilities 300,000 | 500,000 · "
             "Capital 170,000 | 180,000"),
            ("Income Statement (X | Y)", "Sales 300,000 | 500,000 · COGS 140,000 | 200,000 · S&D 50,000 | 60,000 · "
             "EBIT 100,000 | 228,000 · Interest 10,000 | 18,000 · Tax 18,000 | 42,000 · Net Income 72,000 | 168,000"),
        ],
        steps=[
            ("Download the activity worksheet 'Ratio Analysis' from the LMS and open it in Excel.", ""),
            ("Liquidity — Current Ratio = Current Assets / Current Liabilities. X: 190,000/180,000 = 1.06; Y: 400,000/220,000 = 1.82. Y is more liquid.", "X 1.06 · Y 1.82"),
            ("Liquidity — Quick Ratio = (Current Assets − Inventory) / Current Liabilities. X: 150,000/180,000 = 0.83; Y: 330,000/220,000 = 1.50.", "X 0.83 · Y 1.50"),
            ("Profitability — Operating Profit Ratio = EBIT / Sales. X: 100,000/300,000 = 33%; Y: 228,000/500,000 = 46%.", "X 33% · Y 46%"),
            ("Profitability — Net Profit Ratio = Net Income / Sales. X: 72,000/300,000 = 24%; Y: 168,000/500,000 = 34%.", "X 24% · Y 34%"),
            ("Profitability — Return on Capital Employed = EBIT / Capital Employed. X: 100,000/470,000 = 21%; Y: 228,000/680,000 = 34%.", "X 21% · Y 34%"),
            ("Turnover — Inventory Turnover = COGS / Average Inventory. X: 140,000/40,000 = 3.50; Y: 200,000/70,000 = 2.86. X turns stock faster.", "X 3.50 · Y 2.86"),
            ("Turnover — Receivables Turnover = Credit Sales / Average Receivables. X: 300,000/100,000 = 3.00; Y: 500,000/250,000 = 2.00.", "X 3.00 · Y 2.00"),
            ("Turnover — Payables Turnover (assume purchases = 50% of COGS). X: 70,000/100,000 = 0.70; Y: 100,000/150,000 = 0.67.", "X 0.70 · Y 0.67"),
            ("Solvency — Debt-to-Equity = Long-Term Debt / Equity. X: 300,000/170,000 = 1.76; Y: 500,000/180,000 = 2.78. Y is more leveraged.", "X 1.76 · Y 2.78"),
            ("Solvency — Financial Leverage = Total Assets / Equity. X: 650,000/170,000 = 3.82; Y: 900,000/180,000 = 5.00.", "X 3.82 · Y 5.00"),
            ("Write a one-line verdict per ratio family: Y wins on liquidity and profitability; X wins on turnover efficiency and carries less solvency risk.", ""),
        ],
        test="Your worksheet shows all ten ratios for both companies and your verdicts match the model "
             "answers (e.g. Current Ratio X 1.06 vs Y 1.82; D/E X 1.76 vs Y 2.78).",
    ),
    dict(
        num=3,
        topic=2,
        title="Trend Analysis over Four Years",
        objective="LO2 — identify trends by comparing ratios across time periods (K5, A1)",
        desc="You are given four years (2018–2021) of a company's balance sheet and income statement. "
             "Compute year-on-year growth for every line item, derive the key ratios for each year, and "
             "present your observations on the company's direction.",
        build="A completed trend-analysis template: YoY growth for every balance-sheet and income-statement "
              "line plus a four-year ratio table with your observations.",
        services="Microsoft Excel (or Google Sheets), activity template from the LMS",
        data=[
            ("Sales 2018→2021", "180,000 → 198,000 → 217,800 → 239,580 (10% p.a.)"),
            ("Net Income 2018→2021", "53,600 → 64,601 → 76,872 → 90,548"),
            ("Total Assets 2018→2021", "330,000 → 324,800 → 329,983 → 339,581"),
        ],
        steps=[
            ("Download the activity template 'Trend Analysis' from the LMS and open it in Excel.", ""),
            ("Fill the YoY Growth columns of the balance sheet: growth = (this year − last year) / last year. E.g. Inventory 2021 = (40,000 − 34,000)/34,000 = 17.65%.", "=(C5-E5)/E5"),
            ("Fill the YoY Growth columns of the income statement. Sales grow a steady 10% each year while COGS grows only 5% — so Gross Profit growth accelerates (14.47% → 13.78%).", ""),
            ("Compute the liquidity ratios per year: Current Ratio falls from 2.13 (2018) to 1.64 (2021); Acid Ratio falls from 1.47 to 1.22.", "Current = CA/CL"),
            ("Compute the efficiency ratios per year: Asset Turnover rises 0.55 → 0.72; Inventory Turnover rises 4.00 → 6.48; Receivables Turnover rises 3.00 → 3.29.", ""),
            ("Compute the leverage ratios per year: Debt-to-Equity rises 0.56 → 0.75; Debt-to-Assets rises 0.36 → 0.43.", ""),
            ("Compute the profitability ratios per year: Operating Margin rises 0.40 → 0.49; Net Profit Margin rises 0.30 → 0.38.", ""),
            ("Write your observations: profitability and efficiency improve steadily; liquidity declines and leverage creeps up — flag the falling current ratio as the trend to watch.", ""),
        ],
        test="Your ratio table matches the model answers (e.g. 2021 Current Ratio 1.64, Net Profit Margin "
             "0.38) and your observations cover profitability, efficiency, liquidity and leverage trends.",
    ),
]
