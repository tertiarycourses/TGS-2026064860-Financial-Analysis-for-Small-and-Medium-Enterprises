"""Topic 1 — Understanding Financial Statements: hands-on activities.

Activity `num` is the GLOBAL contiguous activity number; `topic` is the topic
number. Steps carry (instruction, working) pairs — the detailed step-by-step
lives in the Learner Guide only; the slide deck shows the activity overview.
"""

DOMAIN1 = [
    dict(
        num=1,
        topic=1,
        title="Prepare a Cash Flow Statement",
        objective="LO1 — understand the balance sheet, income and cash flow statements (K1, K2, K3)",
        desc="You are given the Balance Sheet (2018 and 2019) and the Income Statement (2019) of a "
             "company, all figures in $ million. Working in Excel or on paper, derive the statement "
             "of cash flows — operating, investing and financing — and reconcile it to the change in cash.",
        build="A complete statement of cash flows (operating / investing / financing) derived from the "
              "balance sheet and income statement.",
        services="Microsoft Excel (or Google Sheets), activity worksheet from the LMS",
        data=[
            ("Balance Sheet 2019", "Fixed Assets 100 · Receivables 50 · Inventory 30 · Cash 20 · Total Assets 200 | "
             "Equity 80 · Loans 50 · Payables 40 · Provisions 30 · Total 200"),
            ("Balance Sheet 2018", "Fixed Assets 90 · Receivables 50 · Inventory 30 · Cash 20 · Total Assets 190 | "
             "Equity 80 · Loans 55 · Payables 35 · Provisions 20 · Total 190"),
            ("Income Statement 2019", "Revenue 50 · COGS 10 · Operational costs 20 · Depreciation 5 · "
             "Interest 3 · Tax 2 · Net Income 10"),
        ],
        steps=[
            ("Download the activity worksheet 'Cash Flow Statement' from the LMS (lms-tms.tertiaryinfotech.com) and open it in Excel.", ""),
            ("Compute EBITDA from the income statement: Revenue − COGS − Operational costs = 50 − 10 − 20 = 20.", "EBITDA = 20"),
            ("Compute Net Working Capital movement: NWC = Payables − Receivables − Inventory = 40 − 50 − 30 = −40 (2019) vs 35 − 50 − 30 = −45 (2018); the change is −5.", "ΔNWC = −5"),
            ("Build the Operating Cash Flow block: EBITDA 20, Tax −6, Change in NWC −5, Change in Provisions +10 → Operating Cash Flow = 18.", "OCF = 20 − 6 − 5 + 10 = 18"),
            ("Build the Investing Cash Flow block: Change in Fixed Assets −10, Depreciation 2019 −5, Change in trade payables on capex +1 → Investing Cash Flow = −14.", "ICF = −10 − 5 + 1 = −14"),
            ("Build the Financing Cash Flow block: Change in financial debt −5 (loans 55 → 50), Interest cost −3, Change in earnings −6 → Financing Cash Flow = −14.", "FCF = −5 − 3 − 6 = −14"),
            ("Compute the Net Cash Flow and reconcile: 18 − 14 − 14 = −10.", "Net Cash Flow = −10"),
        ],
        test="Your three blocks total 18 (operating), −14 (investing) and −14 (financing), giving a net "
             "cash flow of −10 that reconciles with the balance-sheet movement.",
    ),
]
