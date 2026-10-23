# E-Commerce Sales Performance Analysis

An Excel project analysing **1,000 online orders (Jan 2024 – Jun 2025)** to understand where revenue comes from and how well orders are being fulfilled. The workbook includes a KPI dashboard, pivot-style summary tables and a cleaned sales dataset.

## Objectives

- Measure total, shipped and average order revenue
- Identify top-performing regions, categories, products and sales representatives
- Track monthly revenue trends
- Assess order fulfilment (shipped vs pending vs cancelled)

## Dataset

| Item | Detail |
|---|---|
| Orders | 1,000 |
| Period | Jan 2024 – Jun 2025 |
| Products | 10 (Laptop, Monitor, Printer, Router, SSD Drive, etc.) |
| Categories | 6 (Electronics, Peripherals, Gaming, Networking, Accessories, Storage) |
| Regions | 5 (Central, East, North, South, West) |
| Sales representatives | 6 |
| Order statuses | Shipped, Pending, Cancelled |

### Columns

| Column | Description |
|---|---|
| Order ID | Unique order identifier |
| Order Date | Date and time of the order |
| Product Name / Category | Product sold and its category |
| Region | Sales region |
| Customer ID | Customer identifier |
| Quantity / Unit Price | Units sold and price per unit |
| Status | Shipped, Pending or Cancelled |
| Sales Representative | Rep who handled the order |
| Total Revenue | Calculated: `Quantity × Unit Price` |
| Order Type | Calculated: `High Value` if revenue ≥ 500, otherwise `Standard Value` |

## Workbook structure

| Sheet | Contents |
|---|---|
| **Overview** | Project summary, key insights, sheet guide and data-cleaning log |
| **Dashboard** | 6 KPI cards and 6 charts |
| **Pivot Analysis** | 6 pivot-style summary tables with totals, % share and reconciliation checks |
| **Sales Data** | Cleaned source data with filters and calculated columns |

### Analyses included

1. Revenue by Region
2. Revenue by Category
3. Order Status Breakdown
4. Orders by Product and Status
5. Revenue by Sales Representative and Order Type
6. Monthly Revenue Trend

## Key findings

- **Gross revenue** is about **$1.45M**, with an average order value of about **$1,451**.
- **Electronics** drives roughly **87%** of revenue; the other five categories share the remaining 13%.
- **Central** is the top region (about $311K, 21.4%), followed closely by West and East. South is the lowest.
- **Monitor** is the most ordered product (136 orders).
- **Charlie** is the top sales representative by revenue (about $267K).
- **May 2024** was the best month (about $106K).
- Only **60.5%** of orders are shipped; **20%** are cancelled and **19.5%** are pending, so shipped revenue is about **$887K** of the $1.45M total.
- **High Value** orders (≥ $500) account for about 94% of revenue.

*Figures reflect the cleaned data and update automatically in the workbook if the data changes.*

## Tools and techniques

- Microsoft Excel
- Formulas: `SUMIFS`, `COUNTIFS`, `IF`, `EDATE`, `INDEX/MATCH`
- Pivot-style summary tables
- Charts and dashboard design

## Data cleaning notes

- **Revenue formula corrected:** in the original file, Total Revenue and Order Type referenced the wrong row for 999 of 1,000 orders. Every row now uses its own `Quantity × Unit Price`.
- **Duplicates removed:** three sheets that repeated the sales data (two exact copies and one Central-region subset) were removed. They can be recreated with the filters on the Sales Data sheet.
- **Layout cleaned:** removed 14 empty columns, added filters and a frozen header row, and consolidated the pivot sheets into one Pivot Analysis sheet.
- **Currency:** the source data has no currency symbol, so US dollars (`$`) are assumed.

## How to use

1. Download `Ecommerce_Sales_Performance_Analysis.xlsx`.
2. Open it in Microsoft Excel (or LibreOffice Calc).
3. Start on the **Overview** sheet, then explore the **Dashboard**.
4. To create native Excel PivotTables, select the **Sales Data** sheet and use *Insert → PivotTable*.