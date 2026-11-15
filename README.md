# FurniturePro & Co. – Sales Insights (Power BI)

A 10-page Power BI report for the BCG-style **FurniturePro & Co.** case. It redesigns the sales dashboard so a business reader gets from the headline numbers to the action behind them.

## What's in the report
| Page | Question it answers |
|---|---|
| Overview | How are sales, profit and orders doing this year? |
| Targets | Which regions reached their targets, and which fell short? |
| Insights | What should we act on first? |
| RFM | Which customers are big spenders who have gone quiet? |
| Returns | How much profit is lost to returns? |
| What-if | What happens to sales and profit if discounts change? |
| Pareto | Which products and customers drive most of the sales? |
| Team | How does each sales team perform? |
| State detail / Customer detail | Drill-through pages for a single state or customer |

## Design rules
- Calibri throughout
- Palette: navy `#172033`, blue `#2563EB`, slate `#64748B`, green `#16834A`, terracotta `#C2410C`, canvas `#F1F5F9`, border `#E2E8F0`
- Every KPI card has a title, a value, a variance and context
- No clipped text and no unnecessary scrollbars

## Model
Semantic model with 17 tables and 170 measures, including target, RFM, returns, what-if discount and Pareto logic, all in DAX.

## Repository layout
```
powerbi/   FurniturePro.pbip + Report and SemanticModel (PBIR/TMDL, Git-friendly)
data/      FurniturePro_Data.xlsx (source data)
tools/     Python scripts used to generate and validate the report files
```

## How to open
1. Install Power BI Desktop (with the PBIP / developer mode preview features on).
2. Open `powerbi/FurniturePro.pbip`.
3. If prompted, point the data source at `data/FurniturePro_Data.xlsx`.

## Author
Vikash Mishra: BI & Analytics. [GitHub](https://github.com/vm8181)
