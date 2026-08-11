# WSQ — Financial Analysis for Small and Medium Enterprises

![WSQ](https://img.shields.io/badge/WSQ-TGS--2026064860-1F6FEB)
![Duration](https://img.shields.io/badge/Duration-2%20days%20%C2%B7%2016%20hours-10B981)
![Skills Framework](https://img.shields.io/badge/TSC-ACC--MAC--5004--1.1%20Financial%20Analysis-7C3AED)

Courseware for the SkillsFuture WSQ course **Financial Analysis for Small and Medium Enterprises**
(Course Code **TGS-2026064860**), conducted by **Tertiary Infotech Academy Pte Ltd** (UEN 201200696W).

📋 **Register here:** [CASL - Financial Analysis for Small and Medium Enterprises](https://www.tertiarycourses.com.sg/casl-financial-analysis-for-small-and-medium-enterprises.html)

## About the Course

Financial statements tell the story of a business — this course teaches SME owners, managers and
executives to read that story and act on it. Learners work through the three key financial
statements, evaluate performance with financial ratios and trends, and apply planning, budgeting
and capital-budgeting techniques to real numbers in Excel.

### Learning Outcomes

- **LO1** — Understand the financial statements such as balance sheet, income, and cash flow statements.
- **LO2** — Evaluate organization's financial performance from the trend of financial ratios.
- **LO3** — Analyze financial statements and prepare the organization's position.

### Topics

| Topic | Coverage | TSC |
|---|---|---|
| 1. Understanding Financial Statements | Overview of finance & chart of accounts · balance sheet · P&L statement · cash flow statement | K1 · K2 · K3 |
| 2. Analysing Financial Ratios | Liquidity, leverage, efficiency & profitability ratios · equity changes statement · trend analysis | K4 · A1 |
| 3. Planning & Budgeting using Financial Statements | Budgeting · capital budgeting (payback, NPV, PI, IRR) · forecasting, variance & financial health | K5 · A2 |

## Repository Structure

```
├── courseware/
│   ├── Financial Analysis for Small and Medium Enterprises-v13.pptx   # trainer slide deck (122 slides)
│   ├── Financial Analysis for Small and Medium Enterprises-v13.pdf    # learner slide PDF
│   ├── LP-Financial Analysis for Small and Medium Enterprises.docx    # Lesson Plan (+ PDF)
│   ├── LG-Financial Analysis for Small and Medium Enterprises.docx    # Learner Guide (+ PDF)
│   ├── assets/                                                        # slide images & diagrams
│   └── archive/                                                       # superseded deck versions
├── activities/                    # 6 hands-on activities — one self-contained folder each
│   ├── README.md                  # index: activity → folder → data set
│   └── activity01 … activity06    # guide (.md) + workbook (.xlsx) + data (.csv) + worksheet (.pdf)
├── LG-Financial Analysis for Small and Medium Enterprises.md          # Learner Guide (Markdown mirror)
└── .claude/                       # single-source courseware build pipeline
```

Every file belonging to an activity lives in that activity's folder — nothing sits loose at the
`activities/` root. For example `activities/activity02/` holds the guide, the Orchid Logistics vs
Marina Retail workbook, both statement CSVs, and the printable worksheet.

## Hands-On Activities

1. **Prepare a Cash Flow Statement** — derive operating / investing / financing cash flows from a balance sheet and income statement.
2. **Ratio Analysis of Two Companies** — liquidity, profitability, turnover and solvency ratios for Companies X and Y.
3. **Trend Analysis over Four Years** — YoY growth and four-year ratio trends.
4. **Payback Period & Discounted Payback Period** — evaluate a $2,000 project at a 10% discount rate.
5. **Net Present Value & Profitability Index** — NPV and PI accept/reject decision.
6. **Solvency & Financial Risk Analysis** — debt-to-equity, financial leverage and interest coverage comparison.

Every activity ships with realistic mock data for fictitious Singapore SMEs (Havenwood Trading,
Orchid Logistics, Marina Retail Group, Sunrise F&B Holdings) as a formatted Excel workbook plus raw
CSV files — regenerate them with `.claude/skills/courseware-build/build/build_lab_data.py`.

## Regenerating the Courseware

All artifacts (PPT, LP, LG, activity files) build from a single content source
(`.claude/skills/courseware-build/build/course_data.py` + `data_domain1..3.py`):

```bash
COURSE_REPO="$PWD" bash .claude/skills/courseware-build/build/build_courseware.sh
```

## Assessment

Final assessment (Day 2, open book): **Written Assessment (SAQ)** — 1 hour, and
**Practical Performance (PP)** — 1 hour. Assessment papers are confidential and are **not**
included in this repository.

## Support

- 📧 enquiry@tertiaryinfotech.com
- ☎️ +65 6100 0613
- 🌐 [www.tertiarycourses.com.sg](https://www.tertiarycourses.com.sg)

---

Powered by [Tertiary Infotech Academy Pte Ltd](https://www.tertiaryinfotech.com/)
