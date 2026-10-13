<div align="center">

# 📊 Data Intelligence Dashboard
### *A 6-Sheet Excel Project: Sales Analysis, Customer Segmentation, Data Quality Checks & Dashboard Charts*

[![Excel](https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Sheets](https://img.shields.io/badge/Worksheets-6-2E7D32?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Records](https://img.shields.io/badge/Transactions-250-1976D2?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Formulas](https://img.shields.io/badge/Formulas-SUMIF%20%7C%20INDEX--MATCH%20%7C%20DATEDIF%20%7C%20PERCENTILE-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)

<br/>

> *"Raw data tells you what happened. A dashboard tells you what to do about it."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Workbook Workflow](#-workbook-workflow)
- [🗂️ Dataset Description](#️-dataset-description)
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

**Data Intelligence Dashboard** is an end-to-end Excel analytics project built on a sales dataset of **250 transactions** from **50 customers**, covering **11 April 2024 to 11 April 2025**. The workbook takes raw transaction data and turns it into KPIs, region / product / customer summaries, a data-quality report, and a chart-based dashboard, using only built-in Excel formulas.

The workbook contains **six worksheets**: `Dashboard`, `Raw Data`, `Analysis`, `Formula Guide`, `Data Quality`, and `Visualizations`.

This project is designed to:
- Calculate headline KPIs (revenue, transactions, units, average order value) with `SUM`, `COUNTA`, and `AVERAGE`
- Summarise sales by region, product, month, and customer using `SUMIF`, `SUMIFS`, and `COUNTIF`
- Segment customers into **High Value** and **Regular** using the 75th percentile (`PERCENTILE`)
- Find the top product and top customer with `INDEX` + `MATCH` + `MAX`
- Add helper columns for tenure, month-end, repeat customers, and name initials (`DATEDIF`, `EOMONTH`, `COUNTIF`, `LEFT`, `MID`, `FIND`)
- Validate the data with a dedicated quality-check sheet
- Present the results as three charts on a dashboard

---

## 🎯 Problem Statement

> **Objective:** Turn a flat file of sales transactions into a decision-ready dashboard that shows where revenue comes from (region, product, month), who the most valuable customers are, and whether the underlying data can be trusted.

| 📂 Business Question | 📄 Sheet | 🔍 How It Is Answered |
|----------------------|----------|-----------------------|
| How is the business doing overall? | Dashboard, Analysis | KPI cards: revenue, transactions, units, average transaction |
| Which regions earn the most? | Analysis, Visualizations | `SUMIF` / `COUNTIF` by region + bar and pie chart |
| Which products sell and earn the most? | Analysis | Units and revenue per product + `INDEX-MATCH` for the top seller |
| How does revenue change month to month? | Visualizations | `SUMIFS` with date boundaries + line chart |
| Who are the best customers? | Analysis | Revenue per customer + 75th-percentile High Value flag |
| Is the data reliable? | Data Quality | Blanks, duplicates, amount mismatches, negative values |

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🧮 **KPI Cards** | Four live dashboard tiles linked to the `Analysis` sheet |
| 🌍 **Regional Summary** | Revenue and transaction count for Central, East, North, South, and West |
| 📦 **Product Summary** | Units sold and revenue for all 10 products |
| 👥 **Customer Segmentation** | 50 customers tagged **High Value** (top quartile by revenue) or **Regular** |
| 🎨 **Colour Scale** | Conditional-formatting heat map on customer revenue (`C33:C82`) |
| 📅 **Monthly Trend** | 13 months of revenue calculated with `SUMIFS` and `EOMONTH` |
| 🔁 **Repeat Customer Flag** | Running `COUNTIF` labels each row "First time" or "Repeat" |
| ✅ **Data Quality Sheet** | Six automated checks: record count, blanks, unique IDs, amount mismatches, negative quantity and price |
| 📈 **Six Charts** | Bar, line, and pie charts on both `Dashboard` and `Visualizations` |
| 📘 **Formula Guide** | Reference sheet with formulas and menu paths for PivotTables, Slicers, Goal Seek, Scenario Manager, and Regression |

---

## 🏗️ Project Structure

```
📦 data-intelligence-dashboard/
│
├── 📄 final_project.xlsx   ← Main workbook with all 6 worksheets (entry point)
│
└── 📄 README.md            ← Project documentation
```

**Worksheets inside `final_project.xlsx`:**

```
1. Dashboard        ← KPI cards, key insights, 3 charts
2. Raw Data         ← 250 transactions + 5 helper formula columns
3. Analysis         ← KPIs, region / product / customer summaries
4. Formula Guide    ← Formula reference and how-to steps
5. Data Quality     ← 6 automated data-validation checks
6. Visualizations   ← Chart source tables + 3 charts
```

---

## 🔄 Workbook Workflow

```
           Raw Data (250 transactions)
                      │
                      ▼
┌────────────────────────────────────────┐
│  Helper Columns (Raw Data, cols O–S)   │  ← Tenure, month-end, repeat flag, initials
└───────────────────┬────────────────────┘
                    │
        ┌───────────┴────────────┐
        ▼                        ▼
┌────────────────────┐   ┌────────────────────┐
│  Data Quality      │   │  Analysis          │  ← KPIs, region, product, customer
│  (6 checks)        │   │  (SUMIF / COUNTIF) │
└────────────────────┘   └─────────┬──────────┘
                                   │
                    ┌──────────────┴──────────────┐
                    ▼                             ▼
          ┌──────────────────┐         ┌──────────────────┐
          │  Visualizations  │         │    Dashboard     │
          │  (chart tables)  │         │  (KPIs + charts) │
          └──────────────────┘         └──────────────────┘
                                   │
                                   ▼
                       Insights & Decisions ✅
```

---

## 🗂️ Dataset Description

The `Raw Data` sheet holds **250 rows** with the following columns:

| Column | Name | Type | Description |
|--------|------|------|-------------|
| A | `Transaction_ID` | Text | Unique ID, e.g. `TRX0155` |
| B | `Date` | Date | Transaction date (11 Apr 2024 – 11 Apr 2025) |
| C | `Customer_ID` | Text | Customer code, e.g. `CUST023` |
| D | `Customer_Name` | Text | Full customer name (50 unique customers) |
| E | `Product_ID` | Text | Product code, e.g. `P009` |
| F | `Product_Name` | Text | One of 10 products |
| G | `Category` | Text | Electronics, Furniture, or Appliances |
| H | `Quantity` | Number | Units purchased |
| I | `Unit_Price` | Number | Price per unit |
| J | `Payment_Method` | Text | Cash, Credit Card, Debit Card, or PayPal |
| K | `Region` | Text | Central, East, North, South, or West |
| L | `Customer_Segment` | Text | Basic, Standard, or Premium |
| M | `Customer_Since` | Date | Date the customer joined |
| N | `Total_Amount` | Number | Quantity × Unit Price |

**Helper columns added with formulas:**

| Column | Name | Formula |
|--------|------|---------|
| O | `Customer_Tenure_Years` | `=DATEDIF(M2,TODAY(),"Y")` |
| P | `Transaction_Month_End` | `=EOMONTH(B2,0)` |
| Q | `Repeat_Customer_Name` | `=IF(COUNTIF($D$2:D2,D2)>1,"Repeat","First time")` |
| R | `Customer_Initials` | `=IFERROR(LEFT(D2,1)&MID(D2,FIND(" ",D2)+1,1),"")` |
| S | `Static_Timestamp` | `=IF(S2="",NOW(),S2)` (needs iterative calculation) |

---

## 🧾 Sheet-by-Sheet Explanation

### 🖥️ 1. Dashboard — `Dashboard`

The front page of the project. Four KPI cards pull their values from the `Analysis` sheet, and three charts (revenue by region, monthly revenue trend, and revenue share by region) sit below, followed by a short list of key insights.

```excel
Total Revenue        =Analysis!B2
Transactions         =Analysis!B3
Units Sold           =Analysis!B4
Average Transaction  =Analysis!B5
```

---

### 🗃️ 2. Raw Data — `Raw Data`

The source table of 250 transactions plus five helper columns (see [Dataset Description](#️-dataset-description)). Every other sheet reads from the ranges `'Raw Data'!A2:N251`.

---

### 📊 3. Analysis — `Analysis`

**KPIs and dates**

| Formula | Purpose |
|---------|---------|
| `=SUM('Raw Data'!N2:N251)` | Total revenue |
| `=COUNTA('Raw Data'!A2:A251)` | Total transactions |
| `=SUM('Raw Data'!H2:H251)` | Total units sold |
| `=AVERAGE('Raw Data'!N2:N251)` | Average transaction value |
| `=MAX(...)` / `=MIN(...)` | Highest and lowest transaction value |
| `=MAX('Raw Data'!B2:B251)` | Latest transaction date |
| `=EOMONTH(B10,0)` | Month end of the latest transaction |
| `=TODAY()` / `=NOW()` | Current date and date-time |

**Region summary** (rows 14–18)

```excel
Revenue       =SUMIF('Raw Data'!K$2:K$251,A14,'Raw Data'!N$2:N$251)
Transactions  =COUNTIF('Raw Data'!K$2:K$251,A14)
```

**Product summary** (rows 21–30)

```excel
Units Sold  =SUMIF('Raw Data'!F$2:F$251,A21,'Raw Data'!H$2:H$251)
Revenue     =SUMIF('Raw Data'!F$2:F$251,A21,'Raw Data'!N$2:N$251)
```

**Customer summary and segmentation** (rows 33–82)

```excel
Transaction Count =COUNTIF('Raw Data'!D$2:D$251,A33)
Revenue           =SUMIF('Raw Data'!D$2:D$251,A33,'Raw Data'!N$2:N$251)
Value Group       =IF(C33>=PERCENTILE($C$33:$C$82,0.75),"High Value","Regular")
```

A colour scale highlights customer revenue in `C33:C82`.

**Top performers**

```excel
Most Purchased Product =INDEX(A21:A30,MATCH(MAX(B21:B30),B21:B30,0))
Most Frequent Customer =INDEX(A33:A82,MATCH(MAX(B33:B82),B33:B82,0))
```

---

### 📘 4. Formula Guide — `Formula Guide`

A reference sheet listing the formulas and Excel menu paths used or required by the project.

| Requirement | Formula / Menu Path |
|-------------|---------------------|
| TODAY / NOW | `=TODAY()`, `=NOW()` |
| DATEDIF tenure | `=DATEDIF(M2,TODAY(),"Y")` |
| EOMONTH | `=EOMONTH(B2,0)` |
| Repeat customer names | `=IF(COUNTIF($D$2:D2,D2)>1,"Repeat","First time")` |
| Compare two lists | `=IF(COUNTIF($A$2:$A$251,A2)>0,"Match","No Match")` |
| Abbreviations | `=IFERROR(LEFT(D2,1)&MID(D2,FIND(" ",D2)+1,1),"")` |
| Most frequently purchased product | `INDEX + MATCH + MAX` |
| Conditional formatting | Home > Conditional Formatting > Data Bars / Color Scales |
| Scenario Manager / Goal Seek | Data > What-If Analysis |
| Linear Regression | Data > Data Analysis > Regression (Analysis ToolPak) |
| PivotTables / Pivot Charts | Insert > PivotTable / PivotChart |
| Slicers and Timeline | PivotTable Analyze > Insert Slicer / Insert Timeline |

---

### ✅ 5. Data Quality — `Data Quality`

| Check | Formula |
|-------|---------|
| Transaction records | `=COUNTA('Raw Data'!A2:A251)` |
| Blank cells in source data | `=COUNTBLANK('Raw Data'!A2:N251)` |
| Unique Transaction_ID count | `=SUMPRODUCT((A2:A251<>"")/COUNTIF(A2:A251,A2:A251))` |
| Amount mismatch rows | `=SUMPRODUCT(--(ABS(N2:N251-H2:H251*I2:I251)>0.01))` |
| Negative quantity rows | `=COUNTIF('Raw Data'!H2:H251,"<0")` |
| Negative unit price rows | `=COUNTIF('Raw Data'!I2:I251,"<0")` |

---

### 📈 6. Visualizations — `Visualizations`

Three small tables feed three charts:

| Table | Formula | Chart |
|-------|---------|-------|
| Monthly Revenue (Apr 2024 – Apr 2025) | `=SUMIFS(N, B, ">="&A3, B, "<"&EOMONTH(A3,0)+1)` | Line chart — *Monthly Revenue Trend* |
| Revenue by Region | `=SUMIF(K, D3, N)` | Bar chart — *Revenue by Region* and pie chart — *Revenue Share by Region* |
| Units Sold by Product | `=SUMIF(F, G3, H)` | Product units reference table |

---

## 🖥️ Sample Output

```
Dashboard — KPI Cards
+---------------+--------------+------------+---------------------+
| Total Revenue | Transactions | Units Sold | Average Transaction |
+---------------+--------------+------------+---------------------+
|    229,192.47 |          250 |        753 |              916.77 |
+---------------+--------------+------------+---------------------+

Analysis — Revenue by Region
+---------+-----------+--------------+
| Region  |   Revenue | Transactions |
+---------+-----------+--------------+
| Central | 41,288.34 |           54 |
| East    | 59,288.39 |           48 |
| North   | 50,808.31 |           59 |
| South   | 36,398.75 |           42 |
| West    | 41,408.68 |           47 |
+---------+-----------+--------------+

Analysis — Product Summary
+--------------+-------+-----------+
| Product      | Units |   Revenue |
+--------------+-------+-----------+
| Laptop       |    75 | 67,499.25 |
| Smartphone   |    96 | 67,199.04 |
| Desk         |    78 | 23,399.22 |
| Monitor      |    88 | 21,999.12 |
| Bookshelf    |   102 | 15,298.98 |
| Office Chair |    52 | 10,399.48 |
| Keyboard     |    89 |  8,009.11 |
| Headphones   |    47 |  7,049.53 |
| Blender      |    87 |  5,219.13 |
| Coffee Maker |    39 |  3,119.61 |
+--------------+-------+-----------+

Data Quality
+-----------------------------+--------+
| Transaction records         |    250 |
| Blank cells in source data  |      0 |
| Unique Transaction_ID count |    250 |
| Amount mismatch rows        |      0 |
| Negative quantity rows      |      0 |
| Negative unit price rows    |      0 |
+-----------------------------+--------+
```

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | 2016 or later | Spreadsheet application |
| 📊 **Aggregate Functions** | Built-in | `SUM`, `AVERAGE`, `MAX`, `MIN`, `COUNTA`, `COUNTIF`, `SUMIF`, `SUMIFS`, `PERCENTILE` |
| 🔎 **Lookup Functions** | Built-in | `INDEX`, `MATCH` |
| 🧠 **Logical Functions** | Built-in | `IF`, `IFERROR` |
| 📅 **Date Functions** | Built-in | `TODAY`, `NOW`, `DATEDIF`, `EOMONTH` |
| 🔤 **Text Functions** | Built-in | `LEFT`, `MID`, `FIND` |
| ✅ **Validation Functions** | Built-in | `SUMPRODUCT`, `COUNTBLANK` |
| 🎨 **Visuals** | Built-in | Bar, line, and pie charts; colour-scale conditional formatting |

---

## 📈 Results & Insights

| Metric | Result |
|--------|--------|
| 💰 **Total revenue** | 229,192.47 across 250 transactions and 753 units |
| 🧾 **Transaction value** | Average 916.77; highest 4,499.95; lowest 59.99 |
| 🌍 **Top region** | **East** leads with 59,288.39 (about 26% of revenue); **South** is lowest at 36,398.75 |
| 🏢 **Busiest region** | **North** has the most transactions (59), but East earns more from fewer sales (48) |
| 💻 **Top products by revenue** | **Laptop** (67,499.25) and **Smartphone** (67,199.04) together make up about 59% of revenue |
| 📦 **Most purchased product** | **Bookshelf** with 102 units, yet only 15,298.98 in revenue |
| 🔌 **Category split** | Electronics 74.9%, Furniture 21.4%, Appliances 3.6% of revenue |
| 👑 **Top customer** | **Mark Carter** with 12 transactions and 15,659.65 in revenue |
| ⭐ **High Value customers** | 13 of 50 customers sit at or above the 75th percentile of customer revenue |
| 📅 **Monthly trend** | Revenue peaks in **January 2025** (25,799.15); the lowest full month is **November 2024** (12,309.44) |
| ✅ **Data quality** | All six checks pass: no blanks, no duplicate IDs, no amount mismatches, no negative values |

---

## ⚠️ Known Issues

| Issue | Detail |
|-------|--------|
| 🔁 **Circular reference** | `Raw Data` column S (`=IF(S2="",NOW(),S2)`) refers to itself and needs **File > Options > Formulas > Enable iterative calculation**; otherwise Excel shows a circular-reference warning |
| 📆 **Volatile dates** | `TODAY()` and `NOW()` recalculate on every open, so customer tenure values change over time |
| 🧭 **Wrong reference in Formula Guide** | `Formula Guide!B4` (`=DATEDIF(M2,TODAY(),"Y")`) points at its own sheet's `M2`, not `Raw Data!M2`, so it returns 126 instead of a real tenure |
| 📋 **Missing Excel Table** | The Dashboard notes mention a `TransactionsTable`, but the data is a plain range; convert it with **Insert > Table** (and name it `TransactionsTable`) before creating PivotTables |
| 🕳️ **Unfilled guide rows** | The `FILTER multi-value` row in `Formula Guide` has no formula yet, and the PivotTable, Slicer, Scenario Manager, Goal Seek, and Regression items are instructions only |
| 📉 **Partial last month** | April 2025 contains only transactions up to 11 April, so its revenue (6,979.83) looks like a drop in the trend line |
| 🔢 **Hard-coded lists** | Region, product, and customer names in `Analysis` are typed in, so new values in `Raw Data` will not appear automatically |
| 📐 **Fixed ranges** | Formulas use `2:251`; rows added below row 251 are not counted unless the ranges are extended |

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🧠 **End-to-End Workflow** | Covers cleaning checks, analysis, segmentation, and presentation in one file |
| 🔗 **Fully Formula-Driven** | Every KPI and chart traces back to `Raw Data`; update the data and results refresh |
| 🖥️ **No Setup Needed** | Opens directly in Excel with no macros, add-ins, or external connections |
| 📊 **Business-Ready Insights** | Answers real questions on regions, products, months, and customers |
| ✅ **Built-In Trust Checks** | The `Data Quality` sheet shows at a glance whether the numbers can be relied on |
| 🧪 **Extensible** | Ready to grow with PivotTables, Slicers, Goal Seek, and Regression, as outlined in `Formula Guide` |

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

> *"Every great analysis starts with clean data — get it right, and the dashboard tells the story."*

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
