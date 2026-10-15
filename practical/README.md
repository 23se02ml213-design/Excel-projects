<div align="center">

# 📊 Excel Retail Sales Dashboard
### *An Interactive 7-Sheet Project Covering Data Cleaning, Summary Tables, Charts & Dropdown-Driven KPIs*

[![Excel](https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Sheets](https://img.shields.io/badge/Worksheets-7-2E7D32?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Formulas](https://img.shields.io/badge/Formulas-SUMIF%20%7C%20SUMIFS%20%7C%20SUMPRODUCT-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Charts](https://img.shields.io/badge/Charts-Bar%20%7C%20Pie%20%7C%20Line-1976D2?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Version](https://img.shields.io/badge/Excel-2021%20%2F%20365-9C27B0?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)

<br/>

> *"A dashboard doesn't show you more data — it shows you the answer."*

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

**Excel Retail Sales Dashboard** is an end-to-end Excel project that takes a raw retail sales dataset through **cleaning, summarisation, visualisation and an interactive dashboard**, all inside a single workbook. The data holds **120 orders from January to December 2025** across four regions, three product categories, twelve products and six salespeople.

The workbook is organised as a small pipeline: a `Raw Data` sheet is cleaned by formulas into `Clean Data`, summarised in `Analysis`, and presented on a `Dashboard` where **Region** and **Category** dropdowns update the KPI cards instantly.

This project is designed to:
- Import a dataset and clean it with formulas instead of manual edits
- Handle missing values and recalculate derived columns (`Sales`, `Profit`)
- Build pivot-style summary tables with `SUMIF`, `COUNTIF`, and `SUMIFS`
- Visualise results with bar, pie, and line charts plus data-bar conditional formatting
- Create an interactive dashboard using **Data Validation dropdowns** and `SUMPRODUCT`
- Document sheets, formulas, and methods clearly for anyone who opens the file

---

## 🎯 Problem Statement

> **Objective:** Build an interactive Excel dashboard on a retail sales dataset, completed in the order of a fixed task plan within a 3-hour limit, covering data cleaning, summary analysis, advanced formulas, charts, interactivity, and documentation.

| 📂 Task | ⏱️ Time | ⚖️ Weightage | 🔍 Skills Tested |
|---------|---------|--------------|------------------|
| 1. Dataset Selection & Import | 15 min | 5% | Data import, initial cleaning |
| 2. Data Cleaning & Preparation | 25 min | 10% | Remove duplicates, handle missing values, format data types |
| 3. Create Pivot Tables | 30 min | 15% | Data analysis, summarisation |
| 4. Advanced Formulas | 35 min | 15% | VLOOKUP, INDEX-MATCH, IF, nested functions |
| 5. Data Visualization | 40 min | 20% | Charts, conditional formatting |
| 6. Interactive Dashboard | 45 min | 25% | Slicers, timeline, dynamic ranges, dropdown filters |
| 7. Documentation | 20 min | 10% | Sheet organisation, naming conventions, instructions |

> ⏳ Total planned time: **210 minutes** (the Tasks sheet plan; the instruction sheet sets a 3-hour limit).

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 📑 **7 Organised Sheets** | Instructions → Tasks → Raw Data → Clean Data → Analysis → Dashboard → Documentation |
| 🧹 **Formula-Based Cleaning** | `Clean Data` references `Raw Data`, so the original data is never overwritten |
| 🩹 **Missing-Value Handling** | Missing units become `0`, missing salesperson becomes `"Unknown"` |
| 🔁 **Recalculated Columns** | `Sales = Units × Unit Price` and `Profit = Sales − Cost` rebuilt on clean data |
| 📊 **Pivot-Style Summaries** | Sales, profit and order counts by region, category and month |
| 🎛️ **Interactive Filters** | Region and Category dropdowns drive the Total Sales, Total Profit and Units Sold cards |
| 📈 **Three Charts** | Region bar chart, category pie chart, and monthly trend line chart |
| 🎨 **Conditional Formatting** | Data bars on the region and category sales columns |
| 📖 **Built-In Documentation** | A `Documentation` sheet explains every area, formula and method used |

---

## 🏗️ Project Structure

```
📦 excel-retail-sales-dashboard/
│
├── 📄 priya_ExcelDashboard.xlsx   ← Main workbook with all 7 worksheets (entry point)
│
└── 📄 README.md                   ← Project documentation
```

**Worksheets inside `priya_ExcelDashboard.xlsx`:**

```
Instructions      Analysis
Tasks             Dashboard
Raw Data          Documentation
Clean Data
```

---

## 🔄 Workbook Workflow

```
Import Retail Sales Dataset (120 orders + 1 duplicate)
      │
      ▼
┌──────────────────────────────┐
│   Raw Data                   │  ← Original data, kept untouched
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│   Clean Data                 │  ← Fill blanks, recalculate Sales & Profit
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│   Analysis                   │  ← SUMIF / COUNTIF / SUMIFS summaries
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│   Dashboard                  │  ← Dropdown filters + KPI cards + 3 charts
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│   Documentation              │  ← Methods, formulas, and notes
└─────────────┬────────────────┘
              │
              ▼
       Insights Delivered ✅
```

---

## 🧾 Sheet-by-Sheet Explanation

### 📘 1. Instructions & Tasks — `Instructions`, `Tasks`

`Instructions` lists the project rules (3-hour limit, choose any Kaggle/UCI dataset, build an interactive dashboard, work in order, save every 15 minutes, submit as `YourName_ExcelDashboard.xlsx`). `Tasks` holds the seven-step plan with skills tested, time estimate, and weightage.

```excel
Total time (minutes) =SUM(D4:D10)
```

---

### 📥 2. Raw Data — `Raw Data`

The original dataset with **12 columns**: Order ID, Order Date, Region, Category, Product, Units, Unit Price, Sales, Cost, Profit, Salesperson, Customer Type. It deliberately contains **one duplicate order** (`ORD0004`) and **missing values** (one blank unit count and one blank salesperson) for cleaning practice.

---

### 🧹 3. Data Cleaning — `Clean Data`

Every cell points back to `Raw Data`, with fixes applied through formulas. The duplicate row sits outside the cleaned range, so the sheet holds exactly 120 unique orders.

| Formula | Purpose |
|---------|---------|
| `='Raw Data'!A2` | Carries Order ID, date, region, category, product, price, cost and customer type across |
| `=IF('Raw Data'!F2="",0,'Raw Data'!F2)` | Replaces missing units with `0` |
| `=IF('Raw Data'!K2="","Unknown",'Raw Data'!K2)` | Replaces missing salesperson with `"Unknown"` |
| `=F2*G2` | Recalculates Sales as Units × Unit Price |
| `=H2-I2` | Recalculates Profit as Sales − Cost |

---

### 📊 4. Analysis — `Analysis`

Pivot-style summary tables built entirely with formulas, with **data bars** applied to the sales columns.

| Table | Formula | Purpose |
|-------|---------|---------|
| Sales by Region | `=SUMIF('Clean Data'!$C$2:$C$121,A5,'Clean Data'!$H$2:$H$121)` | Total sales per region |
| Sales by Region | `=SUMIF('Clean Data'!$C$2:$C$121,A5,'Clean Data'!$J$2:$J$121)` | Total profit per region |
| Sales by Region | `=COUNTIF('Clean Data'!$C$2:$C$121,A5)` | Number of orders per region |
| Sales by Category | `=SUMIF('Clean Data'!$D$2:$D$121,F5,'Clean Data'!$H$2:$H$121)` | Total sales per category |
| Monthly Sales | `=SUMIFS('Clean Data'!$H$2:$H$121,'Clean Data'!$B$2:$B$121,">="&A13,'Clean Data'!$B$2:$B$121,"<"&EDATE(A13,1))` | Sales and profit for each month of 2025 |

---

### 🎛️ 5. Interactive Dashboard — `Dashboard`

Two **Data Validation dropdowns** — **Region** (`All, North, South, East, West`) in `B3` and **Category** (`All, Electronics, Furniture, Office Supplies`) in `E3` — filter three KPI cards. Choosing `All` switches a filter off.

```excel
Total Sales  =SUMPRODUCT((IF($B$3="All",1,'Clean Data'!$C$2:$C$121=$B$3))*(IF($E$3="All",1,'Clean Data'!$D$2:$D$121=$E$3))*'Clean Data'!$H$2:$H$121)
Total Profit =  same pattern, multiplied by 'Clean Data'!$J$2:$J$121
Units Sold   =  same pattern, multiplied by 'Clean Data'!$F$2:$F$121
```

Below the cards sit three charts linked to the `Analysis` tables:

| Chart | Type | Shows |
|-------|------|-------|
| Sales by Region | 📊 Bar | Compares total sales across the four regions |
| Sales by Category | 🥧 Pie | Share of sales per product category |
| Monthly Trend | 📈 Line | Sales and profit month by month through 2025 |

---

### 📖 6. Documentation — `Documentation`

A formula guide covering the dataset, cleaning steps, summary tables, advanced formulas, visualisation, interactive controls, and submission notes — so the workbook explains itself.

---

## 🖥️ Sample Output

```
Dashboard — Default View (Region = All, Category = All)
+----------------+---------------+--------------+
| Total Sales    | Total Profit  | Units Sold   |
+----------------+---------------+--------------+
|     127,626.85 |     41,731.40 |          920 |
+----------------+---------------+--------------+

Dashboard — Filtered View (Region = South, Category = Electronics)
+----------------+---------------+--------------+
| Total Sales    | Total Profit  | Units Sold   |
+----------------+---------------+--------------+
|      34,538.10 |     11,530.88 |          103 |
+----------------+---------------+--------------+

Analysis — Sales by Region
+--------+-------------+--------------+--------+
| Region | Total Sales | Total Profit | Orders |
+--------+-------------+--------------+--------+
| North  |   32,981.26 |    10,259.40 |     27 |
| South  |   44,165.54 |    15,047.12 |     38 |
| East   |   20,943.76 |     7,399.18 |     30 |
| West   |   29,536.29 |     9,025.70 |     25 |
+--------+-------------+--------------+--------+

Analysis — Sales by Category
+-----------------+-------------+--------------+
| Category        | Total Sales | Total Profit |
+-----------------+-------------+--------------+
| Electronics     |   84,037.89 |    27,714.43 |
| Furniture       |   40,857.30 |    13,160.06 |
| Office Supplies |    2,731.66 |       856.91 |
+-----------------+-------------+--------------+
```

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | 2021 / Microsoft 365 | Spreadsheet application |
| 🧠 **Logical Functions** | Built-in | `IF` |
| 📊 **Aggregate Functions** | Built-in | `SUMIF`, `COUNTIF`, `SUMIFS`, `SUMPRODUCT`, `SUM` |
| 📅 **Date Functions** | Built-in | `EDATE` |
| ✅ **Data Validation** | Built-in | Dropdown filters for Region and Category |
| 📈 **Charts** | Built-in | Bar, pie, and line charts |
| 🎨 **Conditional Formatting** | Built-in | Data bars on summary tables |

---

## 📈 Results & Insights

After opening the workbook, the following results are produced:

- ✅ **120 Clean Orders** — One duplicate and two missing values handled without touching the raw data
- 💰 **Total Sales of 127,626.85** with **Total Profit of 41,731.40**, an overall margin of about **32.7%**
- 📦 **920 Units Sold** across 12 products and 6 salespeople
- 🌍 **South Leads** — Highest sales (44,165.54) and the most orders (38); **East is lowest** in sales (20,943.76)
- 💻 **Electronics Dominates** — About **65.8%** of all sales; Office Supplies adds only about 2.1%
- 🏆 **Laptop Is the Top Product** — 62,780.18 in sales, close to half of total revenue
- 📅 **Peak Month** — October 2025 (18,951.10); **weakest month** — April 2025 (718.97)
- 🧑‍💼 **Top Salesperson** — Asha, with 34,430.66 in sales

---

## ⚠️ Known Issues

| Issue | Detail |
|-------|--------|
| 🧪 **Sample dataset** | The data is a generated retail practice dataset; replace it with an approved Kaggle/UCI dataset if required |
| 🎛️ **Charts don't follow filters** | The dropdowns update only the KPI cards; the charts always show overall summaries from `Analysis` |
| 🔢 **Missing units affect profit** | In `Clean Data`, a blank unit count becomes `0`, so Sales becomes `0` while Cost stays — order `ORD0024` shows a profit of −476.14 |
| 🔁 **Fixed cleaning range** | The duplicate `ORD0004` is excluded because `Clean Data` stops at row 121; adding rows to `Raw Data` means extending the formulas and `Analysis` ranges |
| 🧮 **Array formula behaviour** | The `IF` inside `SUMPRODUCT` works directly in Excel 2021 / 365; older versions may need `Ctrl + Shift + Enter` |
| 🍰 **No slicers or timeline** | Native slicers need PivotTables, so dropdowns are used as simple interactive filters |
| 📆 **Fixed months** | Monthly summary rows are typed for January–December 2025 and don't adjust to other years |

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Beginner Friendly** | Follows a clear order, from raw data to a finished dashboard |
| 🔗 **Fully Formula-Driven** | Every number traces back to `Raw Data`, with no pasted values |
| 🛡️ **Non-Destructive Cleaning** | Original data stays intact beside the cleaned version |
| 🖥️ **No Setup Needed** | Opens directly in Excel with no macros or add-ins |
| 📊 **Business-Ready Insights** | Answers region, category, product, and trend questions at a glance |
| 🧪 **Extensible** | Easy to add PivotTables, slicers, `XLOOKUP`, or new KPI cards |
| 📖 **Self-Documented** | Instructions, task plan, and formula guide live inside the workbook |

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

> *"Every great dashboard starts with clean data — get that right, and the insights follow."*

**🎓 Role:** Excel Learner \
**📍 Location:** India \
**🛠️ Skills:** Excel, Data Analysis, Dashboards

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📚 [Microsoft Excel Support](https://support.microsoft.com/excel) — Official function reference
- 🔍 [W3Schools Excel](https://www.w3schools.com/excel/) — Beginner formula reference
- 🧠 [GeeksForGeeks — Excel Tutorial](https://www.geeksforgeeks.org/excel/) — Formula concepts explained
- 📊 [Kaggle Datasets](https://www.kaggle.com/datasets) — Source for practice retail datasets
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support

---

<div align="center">

---

*Made with ❤️ and ☕ — Last updated: October 2026*

</div>
