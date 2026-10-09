<div align="center">

# 📊 Sales Analyzer Dashboard
### *An 8-Sheet Excel Project for Sales, Profit, Customer & Trend Analysis*

[![Excel](https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Sheets](https://img.shields.io/badge/Worksheets-8-2E7D32?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Records](https://img.shields.io/badge/Orders-200-1565C0?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Analysis](https://img.shields.io/badge/Analysis-Dashboard%20%7C%20What--If%20%7C%20Regression-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)

<br/>

> *"A dashboard doesn't show you numbers — it shows you where to look next."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Workbook Workflow](#-workbook-workflow)
- [🧾 Sheet-by-Sheet Explanation](#-sheet-by-sheet-explanation)
- [🖥️ Sample Output](#️-sample-output)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [⚠️ Known Issues](#️-known-issues)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

**Sales Analyzer Dashboard** (titled *PR.2 Analyzer — Sales Dashboard* inside the workbook) is an Excel analytics project built on **200 sales orders** placed between **April 2024 and April 2025**. A single `Data` sheet feeds seven analysis sheets, each answering a different business question — from "which region sells the most?" to "what happens to profit if we change the discount?".

The workbook covers **KPI cards and charts**, **conditional aggregation (`SUMIF`, `SUMIFS`)**, **a formula-built pivot table**, **month-over-month growth**, **customer ranking**, **descriptive statistics**, **what-if scenario analysis**, and **linear regression** of profit against sales.

This project is designed to:
- Turn raw order data into a one-page, chart-driven dashboard
- Summarise sales by region, category, month, and customer
- Measure month-over-month sales and profit growth
- Test how a different discount rate would change total profit
- Quantify the relationship between sales and profit with regression
- Keep every figure formula-driven, so results update when data changes

---

## 🎯 Problem Statement

> **Objective:** Build a single Excel workbook that converts 200 raw sales records into clear business insights — totals, trends, rankings, statistics, a discount scenario, and a profit prediction — without macros or external tools.

| 📂 Category | 📄 Sheets | 🔍 Description |
|-------------|-----------|----------------|
| Reporting | Dashboard | KPI cards plus Region, Month, and Category summaries with charts |
| Source Data | Data | 200 orders with helper columns for month and scenario profit |
| Aggregation | Monthly Summary, Region Product Pivot | Time-based and cross-tab sales totals |
| Customer Analysis | Customer Summary | Purchase value, profit, order count, and value rank per customer |
| Statistics | Descriptive Stats, Regression | Spread of the data and the profit-vs-sales relationship |
| Scenario Planning | What-If | Profit impact of a different discount rate |

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 📈 **Live Dashboard** | Four KPI cards (Sales, Profit, Quantity, Orders) with a bar, line, and pie chart |
| 🧮 **Formula-Driven** | Totals pull from `Data` through `SUM`, `SUMIF`, and `SUMIFS`, so edits flow through |
| 🔀 **Formula-Built Pivot** | Region × Category grid using `SUMIFS` with mixed references (`$A2`, `B$1`) |
| 📆 **Growth Tracking** | Month-over-month sales and profit growth with `IFERROR` protection |
| 🏅 **Customer Ranking** | `RANK` orders 30 customers by total purchase value, with conditional formatting |
| 🎛️ **What-If Input** | Change one discount cell and the scenario profit recalculates for all 200 orders |
| 📉 **Regression Model** | Slope, intercept, R², and correlation, plus a predictor for any sales value |
| 🔒 **Frozen Headers & Filters** | Frozen panes on key sheets and an auto-filter on the `Data` sheet |

---

## 🏗️ Project Structure

```
📦 sales-analyzer-dashboard/
│
├── 📄 Project_2_excel.xlsx   ← Main workbook with all 8 worksheets (entry point)
│
└── 📄 README.md              ← Project documentation
```

**Worksheets inside `Project_2_excel.xlsx`:**

```
1. Dashboard              5. Descriptive Stats
2. Data                   6. Monthly Summary
3. Customer Summary       7. Region Product Pivot
4. What-If                8. Regression
```

**Columns in the `Data` sheet:**

| Column | Field | Type |
|--------|-------|------|
| A | `Customer_ID` | Input |
| B | `Customer_Name` | Input |
| C | `Region` | Input (Central, East, North, South, West) |
| D | `Product_Category` | Input (Books, Clothing, Electronics, Furniture, Office Supplies) |
| E | `Sales` | Input |
| F | `Quantity` | Input |
| G | `Discount` | Input (0% – 20%) |
| H | `Order_Date` | Input |
| I | `Profit` | Input |
| J | `Timestamp` | Formula (`=NOW()`) |
| K | `Month` | Formula (`=TEXT(H2,"mmm-yyyy")`) |
| L | `Profit at Scenario Discount` | Formula (reads the What-If input) |

---

## 🔄 Workbook Workflow

```
                  Raw Orders (Data sheet, 200 rows)
                              │
        ┌─────────────┬───────┴───────┬──────────────┐
        ▼             ▼               ▼              ▼
┌──────────────┐ ┌────────────┐ ┌───────────┐ ┌────────────────┐
│ Monthly      │ │ Region ×   │ │ Descriptive│ │ Regression     │
│ Summary      │ │ Product    │ │ Stats      │ │ (Profit vs     │
│ (growth)     │ │ Pivot      │ │            │ │  Sales)        │
└──────┬───────┘ └─────┬──────┘ └───────────┘ └────────────────┘
       │               │
       └───────┬───────┘
               ▼
      ┌──────────────────┐        ┌──────────────────┐
      │    Dashboard     │        │  What-If         │
      │ KPIs + 3 charts  │        │ (Discount → Profit)│
      └──────────────────┘        └──────────────────┘

   Customer Summary → ranks customers by total purchase value
```

---

## 🧾 Sheet-by-Sheet Explanation

### 📊 1. Dashboard — `Dashboard`

The one-page summary. Four KPI cards sit at the top, with three small tables below that drive a **bar chart**, a **line chart**, and a **pie chart**.

| Formula | Purpose |
|---------|---------|
| `=SUM(Data!E2:E201)` | Total Sales |
| `=SUM(Data!I2:I201)` | Total Profit |
| `=SUM(Data!F2:F201)` | Total Quantity |
| `=COUNTA(Data!A2:A201)` | Total Orders |
| `=SUMIF(Data!$C$2:$C$201,A7,Data!$E$2:$E$201)` | Sales per Region |
| `=SUMIF(Data!$D$2:$D$201,G7,Data!$E$2:$E$201)` | Sales per Category |
| `='Monthly Summary'!B2` | Monthly sales pulled from the Monthly Summary sheet |

---

### 🗂️ 2. Source Data — `Data`

Holds the 200 orders. Three helper columns extend the raw fields:

```excel
Timestamp                    =NOW()
Month                        =TEXT(H2,"mmm-yyyy")
Profit at Scenario Discount  =I2-E2*('What-If'!$B$2-G2)
```

Headers are frozen and an auto-filter covers `A1:L201`.

---

### 👥 3. Customer Summary — `Customer Summary`

Lists 30 customers with total purchase value, total profit, and order count, ranked from highest to lowest spender.

```excel
Value Rank   =RANK(C2,$C$2:$C$31,0)
```

---

### 🎛️ 4. What-If Analysis — `What-If`

Type a **Scenario Discount** into `B2` (default 10%) and the sheet shows how total profit would change.

| Cell | Formula | Purpose |
|------|---------|---------|
| `B3` | `=SUM(Data!I2:I201)` | Current total profit |
| `B4` | `=SUM(Data!L2:L201)` | Total profit under the scenario discount |
| `B5` | `=B4-B3` | Profit change (Scenario − Current) |

**Assumption:** `Adjusted Profit = Profit − Sales × (Scenario Discount − Original Discount)`

---

### 📐 5. Descriptive Statistics — `Descriptive Stats`

Eight metrics across four fields (Sales, Profit, Quantity, Discount).

| Metric | Formula |
|--------|---------|
| Count | `=COUNT(Data!E2:E201)` |
| Sum | `=SUM(Data!E2:E201)` |
| Average | `=AVERAGE(Data!E2:E201)` |
| Median | `=MEDIAN(Data!E2:E201)` |
| Minimum / Maximum | `=MIN(...)` / `=MAX(...)` |
| Standard Deviation | `=STDEV.S(Data!E2:E201)` |

---

### 📆 6. Monthly Summary — `Monthly Summary`

Totals sales and profit for each of the 13 months and calculates growth against the previous month.

```excel
Total Sales   =SUMIFS(Data!$E$2:$E$201,Data!$K$2:$K$201,TEXT(A2,"mmm-yyyy"))
Sales Growth  =IFERROR(B3/B2-1,0)
Profit Growth =IFERROR(C3/C2-1,0)
```

The first month shows `N/A` because it has no earlier month to compare against.

---

### 🔀 7. Region × Product Pivot — `Region Product Pivot`

A pivot-style grid built with formulas instead of a PivotTable, so it recalculates instantly.

```excel
Cell        =SUMIFS(Data!$E$2:$E$201,Data!$C$2:$C$201,$A2,Data!$D$2:$D$201,B$1)
Grand Total =SUM(B2:F2)
```

---

### 📉 8. Regression — `Regression`

Fits a straight line of **Profit (Y) on Sales (X)** and lets you predict profit for any sales amount typed into `B7`.

| Formula | Purpose |
|---------|---------|
| `=SLOPE(Data!I2:I201,Data!E2:E201)` | Profit gained per extra ₹1 of sales |
| `=INTERCEPT(Data!I2:I201,Data!E2:E201)` | Baseline profit at zero sales |
| `=RSQ(Data!I2:I201,Data!E2:E201)` | How much of profit the model explains |
| `=CORREL(Data!I2:I201,Data!E2:E201)` | Strength of the sales–profit link |
| `=B3+B2*B7` | Predicted profit for the sales value in `B7` |

---

## 🖥️ Sample Output

```
Dashboard — KPI Cards
+-------------+--------------+----------------+--------------+
| Total Sales | Total Profit | Total Quantity | Total Orders |
|  195,217.76 |    68,287.01 |          1,999 |          200 |
+-------------+--------------+----------------+--------------+

Region × Category Sales (Region Product Pivot)
+-----------+--------+----------+-------------+-----------+-----------------+-------------+
| Region    | Books  | Clothing | Electronics | Furniture | Office Supplies | Grand Total |
+-----------+--------+----------+-------------+-----------+-----------------+-------------+
| Central   | 11,297 |   12,200 |       6,621 |     4,877 |           8,473 |      43,468 |
| East      | 11,389 |    2,401 |       2,870 |     7,592 |           8,189 |      32,442 |
| North     |  8,037 |    8,423 |       8,192 |     6,978 |           3,346 |      34,975 |
| South     |  9,588 |    8,394 |      11,989 |     7,983 |           2,295 |      40,249 |
| West      |  6,691 |   10,509 |      11,372 |     9,183 |           6,328 |      44,084 |
| Total     | 47,002 |   41,928 |      41,044 |    36,613 |          28,631 |     195,218 |
+-----------+--------+----------+-------------+-----------+-----------------+-------------+

Descriptive Stats (Sales)
+--------------------+---------+
| Average            |  976.09 |
| Median             |  976.72 |
| Minimum / Maximum  | 258.51 / 1,923.66 |
| Standard Deviation |  299.27 |
+--------------------+---------+

What-If (Scenario Discount = 10%)
+--------------------------+------------+
| Current Total Profit     |  68,287.01 |
| Scenario Total Profit    |  68,037.99 |
| Profit Change            |    -249.02 |
+--------------------------+------------+

Regression (Profit vs Sales)
+-----------------------------+----------+
| Slope                       |   0.3509 |
| Intercept                   |  -1.0316 |
| R-squared                   |   0.5982 |
| Correlation                 |   0.7734 |
| Predicted Profit @ ₹1,000   |   349.82 |
+-----------------------------+----------+
```

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | 2016 or later / Microsoft 365 | Spreadsheet application |
| ➕ **Aggregate Functions** | Built-in | `SUM`, `SUMIF`, `SUMIFS`, `COUNT`, `COUNTA`, `AVERAGE` |
| 📐 **Statistical Functions** | Built-in | `MEDIAN`, `MIN`, `MAX`, `STDEV.S`, `RANK` |
| 📉 **Regression Functions** | Built-in | `SLOPE`, `INTERCEPT`, `RSQ`, `CORREL` |
| 🛡️ **Logical Functions** | Built-in | `IFERROR` |
| 🔤 **Text & Date Functions** | Built-in | `TEXT`, `NOW` |
| 📊 **Charts** | Built-in | Bar, line, and pie charts on the Dashboard |
| 🎨 **Conditional Formatting** | Built-in | Highlighting on the Customer Summary |

---

## 📈 Results & Insights

After opening the workbook, the following results are produced:

- 💰 **Total Sales of ₹1,95,217.76** across 200 orders and 1,999 units, with **₹68,287.01 profit** (about a **35% profit margin**)
- 🗺️ **West is the top region** at ₹44,083.58, followed closely by Central (₹43,468.33); **East is the lowest** at ₹32,441.64
- 📚 **Books is the top category** at ₹47,002.46; **Office Supplies is the lowest** at ₹28,630.82
- 🏆 **Best-selling combination:** South + Electronics (₹11,989) and West + Electronics (₹11,372)
- 📆 **Strongest month:** Sep 2024 (₹20,533.50); the sharpest rise was +103% over Aug 2024, and the sharpest fall was Jul 2024 (−42%)
- 👑 **Top customer:** CUST015 (James Wilson) with ₹11,611.06 across 11 orders
- 🎛️ **What-If:** Setting the discount to 10% across all orders lowers total profit by only **₹249.02**
- 📉 **Regression:** Each extra ₹1 of sales adds about **₹0.35 profit**. With **R² = 0.598** and **correlation = 0.77**, sales explain roughly 60% of profit variation

---

## ⚠️ Known Issues

| Issue | Detail |
|-------|--------|
| 📅 **Partial last month** | The data ends on 7 Apr 2025, so Apr 2025 (₹3,212) looks like a steep drop of about −82%, when it is simply an incomplete month. The data also starts on 10 Apr 2024 |
| 🔄 **Volatile `NOW()`** | The `Timestamp` column recalculates on every open, so it shows the current time rather than the order time. It also marks the file as changed each time |
| 📌 **Static Customer Summary** | Totals, profit, and order counts are stored as typed values; only the rank is a formula. They will not update if `Data` changes |
| 👤 **Repeated customer names** | Several different IDs share a name (for example, "James Wilson" appears under CUST003, CUST008, CUST015, CUST021, and CUST028). Always use `Customer_ID` as the unique key |
| 🔗 **Fixed ranges** | Formulas point to rows 2–201. Adding orders below row 201 needs the ranges extended (or the data converted to an Excel Table) |
| 🧮 **Simple What-If model** | The scenario assumes sales volume stays the same when the discount changes, so it shows margin impact only, not demand response |
| 🔤 **Text-based month key** | Monthly formulas match on `TEXT(date,"mmm-yyyy")`, which depends on the system's language and date settings |

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 📊 **One-Glance Reporting** | KPIs and three charts summarise the whole dataset on a single sheet |
| 🔗 **Fully Connected** | One `Data` sheet drives every analysis, with no copy-paste between sheets |
| 🎛️ **Interactive** | Change the scenario discount or the sales value and see results update instantly |
| 🖥️ **No Setup Needed** | Opens directly in Excel with no macros, add-ins, or external connections |
| 🧪 **Extensible** | Easy to add a `PivotTable`, `XLOOKUP` customer search, or `FORECAST` projections |
| 📖 **Readable Structure** | Clearly named sheets with a consistent layout make navigation simple |

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

<div align="center">

### Priya Vadsak

[![GitHub](https://img.shields.io/badge/GitHub-23se02ml213--design-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/23se02ml213-design)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/priya-vadsak-4347294300/)

> *"Every great analysis starts with a well-built formula — get it right, and the data speaks for itself."*

**🎓 Role:** Excel Learner \
**📍 Location:** India \
**🛠️ Skills:** Excel, Data Analysis

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📚 [Microsoft Excel Support](https://support.microsoft.com/excel) — Official function reference
- 🔍 [W3Schools Excel](https://www.w3schools.com/excel/) — Beginner formula reference
- 🧠 [GeeksForGeeks — Excel Tutorial](https://www.geeksforgeeks.org/excel/) — Formula concepts explained
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support
- 📖 [Kaggle Learn](https://www.kaggle.com/learn) — Data analysis courses

---

<div align="center">

---

*Made with ❤️ and ☕ — Last updated: October 2026*

</div>
