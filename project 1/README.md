<div align="center">

# 🧮 Excel Functions Workbook
### *A 16-Sheet Hands-On Project Covering Logical, Lookup, Text, Date & Dynamic-Array Formulas*

[![Excel](https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Sheets](https://img.shields.io/badge/Worksheets-16-2E7D32?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Formulas](https://img.shields.io/badge/Formulas-IF%20%7C%20LOOKUP%20%7C%20DATE%20%7C%20FILTER-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)
[![Version](https://img.shields.io/badge/Excel-2021%20%2F%20365-9C27B0?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-in/microsoft-365/excel)

<br/>

> *"Spreadsheets don't calculate for you — formulas do. Learn the formula, and the data starts to answer back."*

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

**Excel Functions Workbook** is a beginner-friendly Excel project that practises the most useful spreadsheet formulas through **16 focused worksheets**, one concept per sheet. Each sheet holds a small, realistic dataset (sales, students, employees, products) and one or more formulas that solve a real question about that data.

The workbook covers **relative and absolute referencing**, **nested IF and IF with AND/OR**, **conditional counting and summing**, **VLOOKUP, INDEX-MATCH, XLOOKUP and XMATCH**, **text, date and math functions**, the **reference functions INDIRECT and OFFSET**, and the modern dynamic-array function **FILTER**.

This project is designed to:
- Understand the difference between relative (`B2`) and absolute (`$C$2`) cell references
- Build decision logic using `IF`, nested `IF`, `AND`, and `OR`
- Summarise data with `COUNTIFS`, `SUMIFS`, and `AVERAGEIFS`
- Look up values with `VLOOKUP`, `INDEX` + `MATCH`, `XLOOKUP`, and `XMATCH`
- Clean and transform data with text, date, and math functions
- Use dynamic-array formulas such as `FILTER` to extract matching rows automatically

---

## 🎯 Problem Statement

> **Objective:** Build a single Excel workbook that demonstrates the core formula families used in everyday data work, with one clear worked example per function, so that each concept can be learned, tested, and reused independently.

| 📂 Category | 📄 Sheets | 🔍 Description |
|-------------|-----------|----------------|
| Referencing | 01 | Tax, final price, and total amount using absolute references |
| Logical | 02, 03 | Grades, discounts, and eligibility using IF / AND / OR |
| Conditional Aggregation | 04 | Counts, sums, and averages based on conditions |
| Lookup | 05, 06, 07, 10, 11 | Fetching values with VLOOKUP, INDEX-MATCH, XLOOKUP, XMATCH |
| Text, Date & Math | 08, 09, 14, 15 | Name cleanup, tenure, age, rounding |
| Reference & Dynamic | 12, 13, 16 | INDIRECT, OFFSET, and FILTER |

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 📑 **16 Focused Sheets** | Numbered sheets, each dedicated to one formula concept |
| 📥 **Realistic Sample Data** | 10–60 rows per sheet using sales, marks, salary, and product data |
| 🔒 **Absolute References** | Single tax rate in `$C$2` reused across every row |
| 🧠 **Decision Logic** | Nested `IF` grading (A / B / C / Fail) and tiered discounts of 20% / 10% / 0% |
| 📊 **Conditional Reports** | `COUNTIFS`, `SUMIFS`, and `AVERAGEIFS` summary cells |
| 🔎 **Multiple Lookup Styles** | Classic `VLOOKUP` to modern `XLOOKUP` on the same kind of data |
| 🗝️ **Two-Key Lookup** | Combines Salesperson + Month into a helper key for `INDEX` + `MATCH` |
| 🌊 **Dynamic Arrays** | `FILTER` returns all students scoring above 80 in one formula |

---

## 🏗️ Project Structure

```
📦 excel-functions-workbook/
│
├── 📄 Project_1_excel.xlsx   ← Main workbook with all 16 worksheets (entry point)
│
└── 📄 README.md              ← Project documentation
```

**Worksheets inside `Project_1_excel.xlsx`:**

```
01_Relative_Absolute      09_TEXT
02_Nested_IF              10_XLOOKUP
03_IF_AND_OR              11_XMATCH
04_COUNTIFS_SUMIFS        12_INDIRECT
05_VLOOKUP_Student        13_OFFSET
06_VLOOKUP_Product        14_Date_Time
07_INDEX_MATCH            15_Math
08_Employee               16_FILTER
```

---

## 🔄 Workbook Workflow

```
Enter Raw Data in Each Sheet
      │
      ▼
┌──────────────────────────────┐
│   Referencing & Arithmetic   │  ← Sheet 01 (absolute references)
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│   Logical Decisions          │  ← Sheets 02–03 (IF, AND, OR)
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│   Conditional Summaries      │  ← Sheet 04 (COUNTIFS, SUMIFS)
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│   Lookups & Matching         │  ← Sheets 05–07, 10–11
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│   Text, Date & Math          │  ← Sheets 08, 09, 14, 15
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│   Reference & Dynamic Arrays │  ← Sheets 12, 13, 16
└─────────────┬────────────────┘
              │
              ▼
       Results Calculated ✅
```

---

## 🧾 Sheet-by-Sheet Explanation

### 💰 1. Relative & Absolute Referencing — `01_Relative_Absolute`

Calculates tax, final price, and total amount for 20 customers. The tax rate is stored once in `C2` and locked with `$C$2`, so it stays fixed while the price reference moves down each row.

```excel
Tax Amount   =B2*$C$2
Final Price  =B2+E2
Total Amount =B2*D2
```

---

### 🎓 2. Nested IF — `02_Nested_IF`

Two exercises: converting marks to grades, and applying a tiered discount by price.

| Formula | Purpose |
|---------|---------|
| `=IF(C2>=80,"A",IF(C2>=60,"B",IF(C2>=40,"C","Fail")))` | Grades marks into A / B / C / Fail |
| `=IF(C25>=5000,C25*20%,IF(C25>=3000,C25*10%,0))` | 20% discount at ₹5,000+, 10% at ₹3,000+, otherwise none |

---

### ⚖️ 3. IF with AND / OR — `03_IF_AND_OR`

| Formula | Purpose |
|---------|---------|
| `=IF(OR(C2>=5000,D2>=10),"Eligible","Not Eligible")` | Eligible if price is ₹5,000+ **or** quantity is 10+ |
| `=IF(AND(C27>80,D27>80),"Yes","No")` | "Yes" only if **both** Math and Science are above 80 |

---

### 📊 4. COUNTIFS, SUMIFS & AVERAGEIFS — `04_COUNTIFS_SUMIFS`

| Formula | Purpose |
|---------|---------|
| `=COUNTIFS(C26:C45,">50")` | Counts students with Math above 50 |
| `=SUMIFS(E2:E21,C2:C21,"North",D2:D21,"Laptop")` | Total North-region laptop sales |
| `=AVERAGEIFS(C26:C45,C26:C45,">60")` | Average Math score among those above 60 |

---

### 🔎 5. VLOOKUP — `05_VLOOKUP_Student` & `06_VLOOKUP_Product`

| Formula | Purpose |
|---------|---------|
| `=VLOOKUP(D14,A2:B21,2,FALSE)` | Finds a student name from a Student ID |
| `=VLOOKUP(G14,A2:C21,3,FALSE)` | Finds a product price from a Product Code |

---

### 🧩 6. INDEX + MATCH — `07_INDEX_MATCH`

A helper column joins Salesperson and Month into a key such as `Rahul|Feb`, enabling a **two-condition lookup**.

```excel
Key          =A2&"|"&B2
Sales Value  =INDEX($C$2:$C$61,MATCH(F2&"|"&G2,$D$2:$D$61,0))
```

---

### 👥 7. Employee Analysis — `08_Employee`

| Formula | Purpose |
|---------|---------|
| `=DATEDIF(E2,DATE(2026,9,30),"Y")` | Completed years of service |
| `=XLOOKUP(J1,A2:A11,D2:D11,"Not Found")` | Employee salary from Employee ID |
| `=XLOOKUP(J1,A2:A11,E2:E11,"Not Found")` | Joining date from Employee ID |

---

### 🔤 8. Text Functions — `09_TEXT`

| Formula | Purpose |
|---------|---------|
| `=LEFT(A2,FIND(" ",A2)-1)` | Extracts the first name from a full name |
| `=UPPER(A2)` | Converts the name to capitals |
| `=LOWER(A2)` | Converts the name to lowercase |

---

### 🚀 9. XLOOKUP & XMATCH — `10_XLOOKUP` & `11_XMATCH`

| Formula | Purpose |
|---------|---------|
| `=XLOOKUP(J12,A2:A11,C2:C11,"Not Found")` | Salary lookup with a built-in "not found" message |
| `=XLOOKUP(1,(E2:E11=J15)*(F2:F11=J16),G2:G11,"Not Found")` | Multi-condition lookup on Salesperson and Month |
| `=XMATCH(E1,A2:A21,0)` | Returns the position of the first matching product |

---

### 🧭 10. INDIRECT & OFFSET — `12_INDIRECT` & `13_OFFSET`

| Formula | Purpose |
|---------|---------|
| `=INDIRECT(E1)` | Reads the value of the cell whose address is typed in `E1` |
| `=OFFSET(B2,4,0)` | Returns the value 4 rows below the starting cell |
| `=SUM(OFFSET(B2,0,0,5,1))` | Sums a 5-row dynamic range starting at `B2` |

---

### 📅 11. Date, Time & Math — `14_Date_Time` & `15_Math`

| Formula | Purpose |
|---------|---------|
| `=DATEDIF(B2,DATE(2026,9,30),"Y")` | Calculates age in completed years |
| `=D2-C2` | Number of days between start and end dates |
| `=ROUND(B2,2)` | Rounds to 2 decimal places |
| `=CEILING(B2,1)` / `=FLOOR(B2,1)` | Rounds up / down to the nearest whole number |

---

### 🌊 12. FILTER — `16_FILTER`

```excel
=FILTER(A2:C21,C2:C21>80,"No students")
```

Returns every student scoring above 80 as a spilled table, updating automatically if the scores change.

---

## 🖥️ Sample Output

```
Sheet 01 — Relative & Absolute Referencing
+--------+-------+----------+----------+------------+-----------+---------------+
| Name   | Price | Tax Rate | Quantity | Tax Amount | Final     | Total Amount  |
+--------+-------+----------+----------+------------+-----------+---------------+
| Priya  |  6000 |     0.18 |       12 |       1080 |      7080 |         72000 |
| Nikita |  4000 |     0.18 |        5 |        720 |      4720 |         20000 |
+--------+-------+----------+----------+------------+-----------+---------------+

Sheet 04 — Conditional Summaries
+---------------------+-----------+
| Count Math > 50     |        17 |
| North Laptop Sales  |    225000 |
| Average Math > 60   |     79.07 |
+---------------------+-----------+

Sheet 07 — Two-Key Lookup (Rahul | Feb)
+-------------+-------+
| Sales Value | 52000 |
+-------------+-------+

Sheet 16 — FILTER (Score > 80)
+------------+--------+-------+
| Student ID | Name   | Score |
+------------+--------+-------+
|        101 | Priya  |    85 |
|        105 | Amit   |    91 |
|        109 | Pooja  |    88 |
|        111 | Meera  |    83 |
|        114 | Anjali |    94 |
|        116 | Isha   |    87 |
|        118 | Komal  |    81 |
|        120 | Hetal  |    90 |
+------------+--------+-------+
```

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | 2021 / Microsoft 365 | Spreadsheet application |
| 🧠 **Logical Functions** | Built-in | `IF`, `AND`, `OR` |
| 📊 **Aggregate Functions** | Built-in | `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `SUM` |
| 🔎 **Lookup Functions** | Built-in | `VLOOKUP`, `INDEX`, `MATCH`, `XLOOKUP`, `XMATCH` |
| 🔤 **Text Functions** | Built-in | `LEFT`, `FIND`, `UPPER`, `LOWER` |
| 📅 **Date Functions** | Built-in | `DATE`, `DATEDIF` |
| 🔢 **Math Functions** | Built-in | `ROUND`, `CEILING`, `FLOOR` |
| 🌊 **Dynamic Arrays** | Excel 365 / 2021 | `FILTER` |

---

## 📈 Results & Insights

After opening the workbook, the following results are produced:

- ✅ **16 Worksheets** — Each demonstrating a different formula concept with live, working formulas
- 🎓 **Grade Distribution** — Of 20 students: 8 scored A, 7 scored B, 4 scored C, and 1 failed
- ⚖️ **Eligibility** — 11 of 20 customers qualify under the OR rule (price ≥ ₹5,000 or quantity ≥ 10)
- 🏅 **Top Performers** — 8 students score above 80 in both Math and Science
- 💻 **North Laptop Sales** — Total of ₹2,25,000 across the North region
- 🔍 **Lookups** — Employee 202 (Nikita) returns a salary of ₹45,000 and joining date 20-Aug-2019
- 🌊 **FILTER** — 8 students with scores above 80 extracted automatically

---

## ⚠️ Known Issues

| Issue | Detail |
|-------|--------|
| 🆕 **Excel version** | `XLOOKUP`, `XMATCH`, and `FILTER` need Excel 2021 or Microsoft 365; older versions show `#NAME?` |
| 📆 **Fixed reference date** | Age and tenure formulas use `DATE(2026,9,30)` instead of `TODAY()`, so results do not update daily |
| 🔁 **Duplicate keys** | In `07_INDEX_MATCH`, keys like `Rahul|Feb` repeat; `MATCH` always returns the first occurrence |
| 📎 **Input cells** | Lookup inputs (for example the Search ID) are typed values; changing them updates results, but blank inputs return errors |

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Beginner Friendly** | One concept per sheet, easy to follow in order |
| 🔗 **Progressive Learning** | Moves from basic references to modern dynamic-array formulas |
| 📊 **Business-Ready Examples** | Covers sales, salary, grading, and discount scenarios |
| 🖥️ **No Setup Needed** | Opens directly in Excel with no macros or add-ins |
| 🧪 **Extensible** | Easy to add new sheets such as `SUMPRODUCT`, `PivotTable`, or `IFERROR` |
| 📖 **Readable Structure** | Numbered, clearly named sheets make navigation simple |

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

*Made with ❤️ and ☕ — Last updated: September 2026*

</div>
