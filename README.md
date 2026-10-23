# Data Intelligence Dashboard

An Excel analysis and reporting solution for 250 customer transactions (11 Apr 2024 to 11 Apr 2025). It contains cleaned data, formula-driven analysis, live PivotTables, an interactive dashboard, a regression study, What-If scenarios and a written insights report.

**Deliverable:** `Data_Intelligence_Dashboard.xlsx`
**Best opened in:** desktop Microsoft Excel (2016 or later; Excel 365 recommended)

---

## Table of contents

1. [Project at a glance](#1-project-at-a-glance)
2. [Quick start](#2-quick-start)
3. [Workbook map](#3-workbook-map)
4. [Requirements checklist](#4-requirements-checklist)
5. [Dataset and data dictionary](#5-dataset-and-data-dictionary)
6. [Sheet-by-sheet guide](#6-sheet-by-sheet-guide)
7. [Formula reference](#7-formula-reference)
8. [Pivot tables and pivot charts](#8-pivot-tables-and-pivot-charts)
9. [Interactive filters (slicer equivalent)](#9-interactive-filters-slicer-equivalent)
10. [What-If analysis: Scenario Manager and Goal Seek](#10-what-if-analysis-scenario-manager-and-goal-seek)
11. [Regression analysis](#11-regression-analysis)
12. [Key findings](#12-key-findings)
13. [Data quality notes](#13-data-quality-notes)
14. [Assumptions and limitations](#14-assumptions-and-limitations)
15. [How to maintain and extend the workbook](#15-how-to-maintain-and-extend-the-workbook)
16. [Troubleshooting](#16-troubleshooting)
17. [How the file was built](#17-how-the-file-was-built)
18. [Glossary](#18-glossary)

---

## 1. Project at a glance

| Item | Value |
|---|---|
| Records | 250 transactions, no duplicate Transaction_IDs |
| Period | 11 Apr 2024 to 11 Apr 2025 |
| Customers | 50 (CUST001 to CUST050) |
| Products | 10 across 3 categories (Electronics, Furniture, Appliances) |
| Regions | 5 (Central, East, North, South, West) |
| Payment methods | 4 (Cash, Credit Card, Debit Card, PayPal) |
| Total revenue | $229,192.47 |
| Orders / units | 250 orders, 753 units |
| Average order value | $916.77 |
| Workbook size | 12 sheets, about 4,400 formulas, 12 PivotTables, 15 charts |

**Objective:** analyse customer transactions, find insights, and present them through visualisations and a dashboard.

---

## 2. Quick start

1. Open `Data_Intelligence_Dashboard.xlsx` in desktop Excel. If prompted, click **Enable Editing**.
2. Go to the **Dashboard** tab. Change the yellow dropdowns (Region, Category, Payment, Segment, dates) and watch the charts update.
3. Read the **Report** tab for the written insights.
4. Open **Instructions** for the requirement-by-requirement location map.
5. After changing any raw data, press **Data > Refresh All** so the PivotTables update.

**Colour conventions**

| Style | Meaning |
|---|---|
| Blue text on yellow fill | Input you may change (targets, scenario percentages, filters) |
| Black text | Formula or pivot output |
| Navy and blue headers | Section and table headers |

---

## 3. Workbook map

| # | Sheet | Purpose |
|---|---|---|
| 1 | **Instructions** | Requirement checklist, locations, honest notes |
| 2 | **Dashboard** | KPI cards, interactive filters, charts, targets, Top 10, pivot band and pivot charts |
| 3 | **Report** | Formula-driven written insights and recommendations |
| 4 | **Raw Data** | The 250 transactions (Excel Table `tblSales`) plus helper columns |
| 5 | **Analysis** | Date functions, timestamp, KPIs, product/region/category/payment/segment tables, monthly trend, regression formulas, What-If model |
| 6 | **Customers** | GroupBy by Customer_ID, tiers, Top 10, two-list comparison |
| 7 | **Filter** | Multi-value filtering (FILTER-style) by region and minimum amount |
| 8 | **Pivots** | Seven live PivotTables |
| 9 | **Regression Output** | Regression results in the Analysis ToolPak layout |
| 10 | **Scenario Summary** | Scenario comparison and Goal Seek result |
| 11 | **Visualizations** | Static analysis charts linked to the Analysis tables |
| 12 | **Calc** | Filter engine behind the Dashboard dropdowns (no edits needed) |

---

## 4. Requirements checklist

| Requirement | Status | Where |
|---|---|---|
| TODAY, NOW, DATEDIF, EOMONTH | Done | Analysis section 1 (rows 3 to 12); Raw Data columns P, R, S, T |
| FILTER for multi-value returns | Done | Filter sheet (works in all Excel versions); one-cell FILTER formula for Excel 365 shown on the sheet |
| Conditional formatting with icons and arrows | Done | Dashboard variance column (3-arrow icons), card colours, data bars; Analysis MoM arrows and colour scale |
| Timestamp using NOW and iterative calculation | Done | Analysis section 2 (rows 14 to 21) |
| What-If: Scenario Manager | Done | Three scenarios saved in the Analysis sheet; comparison on Scenario Summary |
| What-If: Goal Seek | Done | Model in Analysis section 10; solved answer on Scenario Summary |
| Linear regression with the Analysis ToolPak | Done (formula version) | Regression Output sheet; same layout as the ToolPak |
| Data analysis and storytelling | Done | Report sheet |
| GroupBy to find high-value customers | Done | Customers sheet columns A to J |
| Most frequent product or repeating names | Done | Analysis rows 47 to 51; Raw Data column X; Customers column J |
| Compare two lists and extract matches | Done | Customers columns Q to T |
| Abbreviations with TEXT functions | Done | Raw Data columns U, V, W |
| Pivot tables | Done | Pivots sheet (7) and Dashboard pivot band (5) |
| Pivot charts: bar, line, pie | Done (linked charts) | Dashboard rows 83 to 98 |
| Slicers and timelines | Substituted | Dashboard dropdown filters; see section 9 for adding real slicers |
| KPI indicators with conditional formatting | Done | Dashboard KPI cards and Performance vs Target table |
| Multi-sheet workbook | Done | Raw Data, Analysis, Visualizations, Dashboard and more |
| Summary insights: revenue, top customers, top products, regional sales | Done | Dashboard, Report, Analysis |

---

## 5. Dataset and data dictionary

Source sheet: "Final Project Dataset" (14 columns). The workbook adds 12 helper columns.

### Original columns

| Col | Field | Description |
|---|---|---|
| A | Transaction_ID | Unique ID (TRX0001 to TRX0250) |
| B | Date | Transaction date |
| C | Customer_ID | Customer key (CUST001 to CUST050) |
| D | Customer_Name | Customer name (one name per ID) |
| E | Product_ID | Product key (P001 to P010) |
| F | Product_Name | Laptop, Smartphone, Headphones, Monitor, Keyboard, Desk, Office Chair, Bookshelf, Coffee Maker, Blender |
| G | Category | Electronics, Furniture, Appliances |
| H | Quantity | Units purchased (1 to 5) |
| I | Unit_Price | Price per unit |
| J | Payment_Method | Cash, Credit Card, Debit Card, PayPal |
| K | Region | Central, East, North, South, West |
| L | Customer_Segment | Basic, Standard, Premium |
| M | Customer_Since | Date the customer joined |
| N | Total_Amount | Quantity x Unit_Price |

### Helper columns added in Raw Data

| Col | Field | Formula (row 2) | Purpose |
|---|---|---|---|
| O | Check | `=IF(ROUND(H2*I2-N2,2)=0,"OK","Error")` | Validates Total_Amount |
| P | Month_End | `=EOMONTH(B2,0)` | Month bucketing |
| Q | Month_Label | `=TEXT(B2,"mmm-yy")` | Readable month |
| R | Tenure_Years | `=DATEDIF(M2,B2,"Y")` | Customer age at purchase |
| S | Tenure_Months | `=DATEDIF(M2,B2,"M")` | Same, in months |
| T | Days_Since_Sale | `=TODAY()-B2` | Recency (changes daily) |
| U | Initials | `=LEFT(D2,1)&MID(D2,FIND(" ",D2)+1,1)` | Name abbreviation |
| V | Product_Code | `=UPPER(LEFT(F2,3))` | Product abbreviation |
| W | Order_Code | `=UPPER(LEFT(K2,1))&"-"&TEXT(B2,"yymm")&"-"&RIGHT(A2,3)` | Composite order code |
| X | Repeat_Flag | `=IF(COUNTIF($D$2:D2,D2)>1,"Repeat","First")` | Flags repeat purchases |
| Y | Filter_Seq | Running count of rows matching the Filter sheet inputs | Powers the Filter sheet |
| Z | Year_Month | `=TEXT(B2,"yyyy-mm")` | Sortable month key for pivots |

---

## 6. Sheet-by-sheet guide

### Instructions
Requirement checklist, where each item lives, and the honest notes about what is substituted or approximate.

### Dashboard
Top to bottom:

| Rows | Section | Notes |
|---|---|---|
| 1 to 2 | Title banner and data range | Subtitle refreshes with NOW() |
| 4 to 7 | **KPI cards** | Revenue, Orders, Units, Avg Order Value, Customers, Repeat Rate; green or red text shows variance against target |
| 9 to 14 | **Interactive filters** | Inputs in B11 (Region), E11 (Category), H11 (Payment), K11 (Segment), N11 (From), Q11 (To); strip below shows the selected view |
| 16 to 47 | **Trends and mix** | Monthly line, region bars, category doughnut, top products, payment pie; all respond to the filters |
| 49 to 61 | **Performance vs target** and **Top 10 customers** | Arrow icons on variance; customer tier counts; data bars on share |
| 62 to 82 | **Pivot band** | Five live PivotTables (header row 68, data from row 69) |
| 83 to 98 | **Pivot charts** | Bar (region), line (month), pie (category) linked to the pivots |

### Report
Eleven short sections written as formulas, so the numbers update with the data: big picture, who buys, what sells, where, when, trend, payments, cross-sell, What-If, data-quality note, recommendations.

### Raw Data
The original transactions as Excel Table `tblSales` (A1:Z251) with filter buttons and frozen header. Helper columns are in section 5.

### Analysis

| Rows | Section |
|---|---|
| 3 to 12 | 1. Date and time functions |
| 14 to 21 | 2. Timestamp with NOW() and iterative calculation |
| 23 to 31 | 3. Headline KPIs (revenue, orders, units, AOV, customers, repeat rate, top-5 share) |
| 33 to 51 | 4. Products table plus most-frequent product, units, revenue, buyer, top customer |
| 54 to 63 | 5. Regional sales with top and weakest region |
| 65 to 70 | 6. Category sales |
| 72 to 79 | 7. Payment methods |
| 81 to 86 | 8. Customer segment |
| 88 to 116 | 9. Monthly trend (rows 90 to 102) and regression formulas (rows 106 to 115) |
| 118 to 131 | 10. What-If model |

### Customers

| Columns | Content |
|---|---|
| A to J | One row per customer: Total_Spend, Orders, Avg_Order, Spend_Rank, Tier, Laptop_Orders, Laptop_Seq, Repeat_Status |
| L to O | Top 10 customers with share of revenue |
| Q to T | Two-list comparison: List A (top 15 spenders), List B (laptop buyers), Matches |

Tier rule: **High-Value** = spend at or above the 80th percentile; **Mid** = 50th to 80th; **Low** = below the median.

### Filter
Pick a Region (B4) and a minimum Total_Amount (B5). The table lists up to 30 matching transactions, and B6 shows the match count. Raw Data column Y does the work, so it runs in every Excel version.

### Pivots
Seven PivotTables with two report filters each (see section 8).

### Regression Output, Scenario Summary
See sections 10 and 11.

### Visualizations
Five analysis charts linked to the Analysis tables: region, monthly trend, category share, product revenue, payment mix.

### Calc
The filter engine: dropdown lists, wildcard criteria (`*` means "All"), and the SUMIFS/COUNTIFS tables that feed the Dashboard charts. No manual edits needed.

---

## 7. Formula reference

### Date and time

```excel
=TODAY()
=NOW()
=EOMONTH(TODAY(),0)                      ' end of current month
=EOMONTH(TODAY(),-1)+1                   ' start of current month
=DATEDIF(MIN(Dates),MAX(Dates),"M")      ' dataset span in months
=DATEDIF(M2,B2,"Y")                      ' tenure in years
```

### Timestamp that freezes (iterative calculation)

Enable **File > Options > Formulas > Enable iterative calculation** (already saved ON in this file). Then, with the entry in A16 and the timestamp in B16:

```excel
=IF(A16="","",IF(OR(B16="",B16=0),NOW(),B16))
```

The formula refers to its own cell on purpose. It stamps the time once, when something is typed in column A, and then keeps that value.

### FILTER

Excel 365 one-cell version (shown on the Filter sheet):

```excel
=FILTER(tblSales[[Transaction_ID]:[Total_Amount]],
        (tblSales[Region]=B4)*(tblSales[Total_Amount]>=B5),
        "No matching records")
```

The workbook's own version uses a running match counter (Raw Data column Y) with INDEX/MATCH, so it works without dynamic arrays.

### GroupBy and ranking (Customers sheet)

```excel
Total spend : =SUMIFS(Amounts,CustomerIDs,A4)
Orders      : =COUNTIFS(CustomerIDs,A4)
Rank        : =RANK(C4,$C$4:$C$53,0)+COUNTIF($C$4:C4,C4)-1     ' unique, tie-safe
Tier        : =IF(C4>=PERCENTILE($C$4:$C$53,0.8),"High-Value",
                 IF(C4>=PERCENTILE($C$4:$C$53,0.5),"Mid","Low"))
```

### Most frequent product, repeat names

```excel
=INDEX(ProductNames,MATCH(MAX(OrderCounts),OrderCounts,0))
=IF(COUNTIF($D$2:D2,D2)>1,"Repeat","First")
```

### Compare two lists

```excel
=IF(COUNTIF($S$4:$S$53,R4)>0,R4,"")                  ' name appears in both lists
=IFERROR(INDEX(ListB,MATCH(R4,ListB,0)),"No match")  ' alternative
```

### Abbreviations (TEXT and string functions)

```excel
=LEFT(D2,1)&MID(D2,FIND(" ",D2)+1,1)                 ' Paul Baker -> PB
=UPPER(LEFT(F2,3))                                   ' Monitor -> MON
=UPPER(LEFT(K2,1))&"-"&TEXT(B2,"yymm")&"-"&RIGHT(A2,3)   ' C-2404-155
```

### Dashboard filter engine (Calc sheet)

```excel
Criterion:  =IF(Dashboard!$B$11="All","*",Dashboard!$B$11)
Revenue:    =SUMIFS(Amount, Region,$B$10, Category,$B$11, Payment,$B$12,
                    Segment,$B$13, Date,">="&$B$14, Date,"<="&$B$15)
```

---

## 8. Pivot tables and pivot charts

All PivotTables use the Raw Data table (A1:Z251) as their source.

### Pivots sheet (start row 6, header row 10)

| Name | Rows | Values | Report filters |
|---|---|---|---|
| P1_Region | Region | Sum of Total_Amount | Category, Payment_Method |
| P2_Monthly | Year_Month | Sum of Total_Amount | Region, Category |
| P3_Category | Category | Sum of Total_Amount | Region, Payment_Method |
| P4_Products | Product_Name | Sum of Total_Amount | Region, Category |
| P5_Payment | Payment_Method | Count of Transaction_ID | Region, Category |
| P6_Customers | Customer_Name | Sum of Total_Amount | Region, Category |
| P7_Region_x_Category | Region (columns: Category) | Sum of Total_Amount | Payment_Method |

### Dashboard pivots (start row 64, header row 68)

D1_Region, D3_Category, D5_Payment, D2_Monthly, D4_Products (same definitions as above).

### Pivot charts

| Chart | Pivot | Type |
|---|---|---|
| Revenue by Region | D1_Region | Bar (column) |
| Monthly Revenue | D2_Monthly | Line |
| Category Mix | D3_Category | Pie |

**Important:** these are charts linked to the PivotTable cell ranges. They update when the pivot refreshes, but they are not Excel "PivotChart" objects (no field buttons). If you change a pivot's layout or number of items, re-point the chart's data range.

**Refreshing:** right-click a pivot > Refresh, or **Data > Refresh All**. Rows added to the Raw Data table are picked up automatically because the table grows. The pivot source range is A1:Z251, so if you add rows, update it via PivotTable Analyze > Change Data Source (or point it at `tblSales`).

---

## 9. Interactive filters (slicer equivalent)

The Dashboard filter bar acts like slicers for the top charts:

| Filter | Cell | Options |
|---|---|---|
| Region | B11 | All, Central, East, North, South, West |
| Category | E11 | All, Appliances, Electronics, Furniture |
| Payment | H11 | All, Cash, Credit Card, Debit Card, PayPal |
| Segment | K11 | All, Basic, Premium, Standard |
| From date | N11 | any date (default 1 Apr 2024) |
| To date | Q11 | any date (default 30 Apr 2025) |

Filters combine (AND). The "Selected view" strip shows revenue, orders, units, average order and share of total. Verified example: Region = East, Segment = Premium, From = 1 Jul 2024 returns 11 orders and $15,269.60.

### Adding real Excel slicers and a timeline (about one minute)

1. Click any PivotTable on the **Pivots** sheet.
2. **PivotTable Analyze > Insert Slicer**, tick Region, Category, Customer_Segment, Payment_Method.
3. **PivotTable Analyze > Insert Timeline**, choose Date.
4. Right-click each slicer or timeline > **Report Connections** > tick all pivots.

Real slicers could not be written reliably outside Excel, which is why dropdown filters are used here.

---

## 10. What-If analysis: Scenario Manager and Goal Seek

### Model (Analysis section 10, rows 118 to 131)

| Row | Item | Cells |
|---|---|---|
| 120 | Price change % | B (live), C Base, D Optimistic, E Pessimistic |
| 121 | Volume change % | same |
| 122 | Discount % | same |
| 123 | Baseline revenue | `=$B$24` |
| 124 | Projected revenue | `=Baseline*(1+Price)*(1+Volume)*(1-Discount)` |
| 127 to 129 | Goal Seek target uplift, target revenue, closed-form price change needed | |

### Saved scenarios

| Scenario | Price | Volume | Discount | Projected revenue (rounded) |
|---|---|---|---|---|
| Base case | 0% | 0% | 0% | $229,192 |
| Optimistic | +5% | +10% | 0% | $264,717 |
| Pessimistic | -3% | -8% | 5% | $194,305 |

View them: **Data > What-If Analysis > Scenario Manager** on the Analysis sheet, select a scenario, click **Show**. The changing cells are B120:B122 and the result cell is B124. Click **Summary** to generate Excel's own scenario report.

> Scenario percentages are placeholder assumptions, not dataset values. Edit them to suit your case.

### Goal Seek

1. Data > What-If Analysis > Goal Seek.
2. **Set cell:** B124. **To value:** the number in B128 (target revenue). **By changing:** B120.
3. Result: a price change of **20.00%** reaches the default +20% revenue target (with volume and discount at 0%). This matches the closed-form check in B129 and is shown on **Scenario Summary** (difference to target = $0.00).
4. Reset B120 to 0 afterwards.

---

## 11. Regression analysis

**Question:** Is monthly revenue trending up or down?

**Setup:** Y = monthly revenue, X = month index 1 to 12 (Apr 2024 to Mar 2025). April 2025 is excluded because the data ends on 11 Apr 2025 (partial month).

**Where:** the **Regression Output** sheet uses the same layout as Data > Data Analysis > Regression: Summary Output, ANOVA, coefficients with 95% confidence limits, residual output, a line-fit plot with trendline and a residual plot. Analysis rows 106 to 115 give a quick formula cross-check (SLOPE, INTERCEPT, RSQ, STEYX, LINEST, FORECAST).

| Statistic | Value |
|---|---|
| Slope | +$159.50 per month |
| Intercept | $17,480.94 |
| R Square | 0.0213 |
| Standard error | $4,091.24 |
| t Stat (slope) | 0.466 |
| P-value / Significance F | 0.651 |
| 95% CI for slope | about -$603 to +$922 |

**Interpretation:** the trend is **not statistically significant**. Revenue is essentially flat, with large month-to-month swings (for example, strong Dec 2024 and Jan 2025, weak Nov 2024). Do not read the forecast as a reliable prediction.

**Note on the ToolPak:** the Regression Output sheet is built with live formulas (LINEST, FDIST, TDIST), so it reproduces ToolPak numbers and updates with the data. To produce the add-in's own output: File > Options > Add-ins > Analysis ToolPak, then Data > Data Analysis > Regression with Y = Analysis!E90:E101 and X = Analysis!C90:C101, labels off.

---

## 12. Key findings

| Topic | Finding |
|---|---|
| Revenue | $229,192 from 250 orders; average order $916.77 |
| Customers | 50 active; 47 (94%) bought more than once |
| Concentration | Top 5 customers generate about 24.6% of revenue (moderate concentration) |
| Top customer | Mark Carter, $15,659.65 |
| Most frequently bought | Bookshelf |
| Highest revenue | Laptop ($67,499), then Smartphone ($67,199) |
| Category mix | Electronics 75%, Furniture 21%, Appliances 4% |
| Regions | East strongest ($59,288), South weakest ($36,399), a gap of about 63% |
| Payments | Cash is most used (76 orders) |
| Cross-sell | 12 of the top 15 spenders also bought a Laptop |
| Trend | Flat; no significant growth or decline (p = 0.65) |
| Targets | Revenue about 4.5% below the placeholder target; AOV about 1.9% above |

**Recommended actions**

1. Upsell laptops and smartphones to high-value customers who have not bought them.
2. Run targeted promotions in the South region.
3. Keep Bookshelf stocked as the volume driver and protect Laptop and Smartphone margins as the revenue drivers.
4. Review the revenue trend monthly from the Dashboard.
5. Fix the customer-attribute inconsistencies described below at the source.

---

## 13. Data quality notes

- **Totals are consistent.** Quantity x Unit_Price equals Total_Amount on all 250 rows (Check column).
- **No duplicate Transaction_IDs**, and no customer joined after the date of a purchase.
- **Customer attributes vary per transaction.** The same Customer_ID appears with several Regions, Segments and Customer_Since dates (on average about three regions per customer, up to five). Each customer ID has exactly one name, and each product ID exactly one name.
  - Consequence: Region and Segment behave as per-transaction fields. Customer-level analysis groups by **Customer_ID**; regional analysis uses the transaction's Region.
  - Customer_Since and Tenure_Years are therefore indicative only.
- **April 2025 is a partial month** (data ends 11 Apr). It is shown in charts but excluded from the regression.
- **Today-based values change over time.** Days_Since_Sale and "Days since last transaction" use TODAY(), so they grow every day.

---

## 14. Assumptions and limitations

- **Targets** on the Dashboard (revenue $240,000; orders 260; average order $900; repeat rate 95%) are placeholder assumptions, not dataset values. The "last full-month revenue" target is the average of earlier months.
- **Scenario inputs** and the 20% Goal Seek uplift are placeholders.
- **Slicers and timelines** are not included as native Excel objects (see section 9).
- **Pivot charts** are ranges linked to PivotTables, not native PivotChart objects (see section 8).
- **Regression** is a simple one-variable trend model on 12 data points; treat it as descriptive.
- **Top 10 filtering inside a pivot** is not set; the Top 10 list comes from the Customers sheet. To add it in a pivot: row field dropdown > Value Filters > Top 10.
- **Compatibility:** Scenario Manager, PivotTables with slicers, and the ToolPak need desktop Excel. Google Sheets and some online viewers do not support these.
- **Testing:** the workbook was checked with LibreOffice and validated against the Office Open XML schema (zero formula errors, all key figures cross-checked against an independent calculation). It was not opened in desktop Excel during creation. If Excel shows a repair prompt, please report it.

---

## 15. How to maintain and extend the workbook

### Add new transactions
1. Paste new rows directly below the last row of the `tblSales` table in Raw Data; the table and its helper-column formulas extend automatically.
2. Total_Amount must equal Quantity x Unit_Price (the Check column flags mismatches).
3. The Analysis, Customers and Calc formulas use fixed ranges (rows 2 to 251). Extend them (or convert them to `tblSales[...]` references) when the data grows beyond row 251.
4. Update the PivotTable source range and click **Data > Refresh All**.
5. New customers, products, regions or payment methods need a new row in the matching Analysis/Customers/Calc lists.

### Change a target or scenario
Edit the blue-on-yellow cells: Dashboard F51:G54 (targets), Analysis B120:E122 (scenarios) and B127 (uplift target).

### Add a KPI card or chart
Copy an existing card or chart on the Dashboard and point it at a new cell or range in Analysis or Calc.

### Print the Dashboard
The Dashboard is set to landscape, fit to one page wide. Use File > Print.

---

## 16. Troubleshooting

| Symptom | Cause and fix |
|---|---|
| Timestamps show 0 or an error | Turn on iterative calculation: File > Options > Formulas |
| PivotTables look unchanged after editing data | Data > Refresh All |
| Dashboard charts empty or zero | A filter combination returns no rows; set the dropdowns back to All and the date range to 1 Apr 2024 to 30 Apr 2025 |
| Filter sheet shows no rows | Region and minimum amount produce no matches; lower the minimum in B5 |
| `#NAME?` on a FILTER formula | FILTER needs Excel 365 or 2021; use the Filter sheet table instead |
| Pivot chart range is wrong after changing a pivot | Select the chart > Chart Design > Select Data and re-point the ranges |
| Scenarios missing | Data > What-If Analysis > Scenario Manager must be run on the Analysis sheet |
| Excel shows a repair prompt on opening | Accept the repair, then please report it; the workbook data is not affected |
| Goal Seek result differs | Reset Analysis B120:B122 to 0 first, then rerun |

---

## 17. How the file was built

- **Data and formulas:** generated with Python (pandas and openpyxl), then recalculated with LibreOffice.
- **PivotTables:** created through LibreOffice's UNO interface, then relabelled and restyled (Grand Total labels, Medium 2 pivot style, "Revenue ($)" and "Orders" captions).
- **Scenario Manager scenarios:** written directly into the Analysis sheet XML and validated against the Office Open XML schema.
- **Charts:** openpyxl, then normalised by LibreOffice, with data-label formats corrected to `$#,##0`.
- **Verification:** every headline number was recomputed in pandas; the dropdown filter engine was tested against pandas for an East, Premium, from-July-2024 filter; pivot and regression outputs were compared against independent calculations.

---

## 18. Glossary

| Term | Meaning |
|---|---|
| AOV | Average order value: revenue divided by orders |
| GroupBy | Summarising rows by a key (here Customer_ID) using SUMIFS and COUNTIFS |
| High-Value customer | Spend at or above the 80th percentile of customer spend |
| Repeat rate | Share of customers with two or more orders |
| MoM | Month over month change |
| Pivot band | The section of the Dashboard holding live PivotTables and their charts |
| Scenario Manager | Excel tool that stores named sets of input values and compares results |
| Goal Seek | Excel tool that finds the input value needed to reach a target result |
| ToolPak | Excel's Analysis ToolPak add-in, which includes the Regression tool |
| Iterative calculation | Excel setting that allows deliberate circular references (used for timestamps) |
# final-project-excel
