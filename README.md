# Weekly Compliance Report

A Python script that takes one week of plant data, checks it against set limits, and makes a weekly compliance PDF plus an Excel log of violations that is tracked till each one is closed.

The data here is dummy data for a made-up "Plant A". I automated similar compliance reports for manufacturing customers at work; I can't share those, so I rebuilt the idea with made-up data.

## Sample output

![Page 1 of the weekly compliance report](docs/report_preview.png)

Sample report generated from the dummy data (page 1 of 4). Full PDF: [output/sample_weekly_compliance_report.pdf](output/sample_weekly_compliance_report.pdf). Violation log: [output/violation_log.xlsx](output/violation_log.xlsx)

## Why I made this

Every week someone had to pull data from different places (device status, temperature logs, work in progress, safety cameras, inward inspection, yield, report submissions), compare it with the rules by hand, make tables in Excel and write a summary. It took hours, the rules were only in people's heads, and issues found one week were often forgotten by the next.

## What it does

1. Reads one week of data, plus the previous week for comparison.
2. Checks it against limits kept in a settings file, [config/thresholds.yaml](config/thresholds.yaml), so a limit can be changed without touching the code.
3. Makes a 4-page PDF. Page 1 has 7 KPI cards with red / amber / green status, change vs last week and the top 3 actions for the week with owners. Then one section per check with the count, change vs last week, a daily chart and a short table of the worst items.
4. Keeps a violation log in Excel (ID, date, category, detail, owner, status, closed on) with target dates, overdue flags, an Open / Closed dropdown and a summary sheet that updates when an item is closed.

The weekly work goes from hours to one command, and every issue has an ID, an owner and a closure date, so late closures are visible.

## The 7 checks

| # | Check | Example rule (can be changed) |
|---|-------|-------------------------------|
| 1 | Devices offline | Device offline more than 2 hours in a day |
| 2 | Temperature out of range | Hourly reading outside the allowed band for the furnace |
| 3 | Work in progress ageing | Job still open more than 7 days after start |
| 4 | Safety violations | Missing helmet / safety shoes / restricted zone camera events, by area |
| 5 | Inward material rejections | Supplier lot with more than 3% rejected |
| 6 | Yield below target | Heat yield below the furnace target |
| 7 | Late report submissions | Daily report more than 30 min late, or missing |

## How to run

Needs Python 3.10+ and Google Chrome (for the PDF).

```bash
pip install -r requirements.txt
python generate_sample_data.py   # optional, makes the dummy data again
python build_report.py
python build_report.py --no-pdf  # only HTML and Excel, no Chrome needed
```

To change a rule, edit `config/thresholds.yaml` (for example `max_age_days: 7`) and run `build_report.py` again.

## Notes

- `config/thresholds.yaml` also has the owners and the action wording.
- `generate_sample_data.py` makes two weeks of dummy data with a fixed seed.
- The report layout is in `templates/` (Jinja2 HTML and CSS).
- Made with Python, pandas, openpyxl, Jinja2 and headless Chrome.
- All names, areas, suppliers and figures are made up.
