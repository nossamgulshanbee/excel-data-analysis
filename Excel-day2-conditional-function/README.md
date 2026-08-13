# Sales Performance Analysis

## 📌 Project Overview

This project is part of my Excel learning journey for Data Analytics.

The objective of this project is to analyze employee sales performance using Excel conditional functions and apply the concepts learned during Day 2 practice.

---

## 📊 Dataset

The dataset contains employee-level sales information with the following columns:

- Employee
- Region
- Sales
- Rating
- Experience
- Performance
- Eligibility

The dataset contains 10 employees from three regions: South, North, and West.

---

## 🎯 Project Objectives

The main objectives of this project are:

- Classify employee performance based on sales.
- Identify eligible employees based on sales and rating.
- Analyze sales by region.
- Calculate conditional counts, sums, and averages.
- Answer business questions using Excel formulas.
- Practice applying multiple conditions to real-world-style data.

---

## 🧠 Excel Concepts Practiced

The following Excel concepts and functions were practiced:

- `IF()`
- `IFS()`
- `AND()`
- `OR()`
- Nested `IF()`
- `COUNTIF()`
- `COUNTIFS()`
- `SUMIF()`
- `SUMIFS()`
- `AVERAGEIF()`
- `AVERAGEIFS()`
- Single-cell references
- Range references
- Multiple-condition analysis

---
IF()          → Make decisions
IFS()         → Multiple conditions
AND()         → All conditions must be TRUE
OR()          → At least one condition must be TRUE
COUNTIF()     → Count with one condition
COUNTIFS()    → Count with multiple conditions
SUMIF()       → Sum with one condition
SUMIFS()      → Sum with multiple conditions
AVERAGEIF()   → Average with one condition
AVERAGEIFS()  → Average with multiple conditions

## 🧮 Key Excel Formulas

1. Performance Classification
=IF(C2>=80000,"Excellent",IF(C2>=60000,"Good",IF(C2>=40000,"Average","Poor")))

2. Performance Classification using IFS
=IFS(C2>=80000,"Excellent",C2>=60000,"Good",C2>=40000,"Average",TRUE,"Poor")

3. Employee Eligibility
=IF(AND(C2>=60000,D2>=4),"Eligible","Not Eligible")

4. Count South Employees
=COUNTIF(B2:B11,"South")

5. Count Employees with Sales >= 70000
=COUNTIF(C2:C11,">=70000")

6. Count South Employees with Sales >= 50000
=COUNTIFS(B2:B11,"South",C2:C11,">=50000")

7. Total South Sales
=SUMIF(B2:B11,"South",C2:C11)

8. Total West Sales
=SUMIF(B2:B11,"West",C2:C11)

9. Total North Sales >= 70000
=SUMIFS(C2:C11,B2:B11,"North",C2:C11,">=70000")

10. Average South Sales
=AVERAGEIF(B2:B11,"South",C2:C11)

11. Average West Sales with Rating >= 3
=AVERAGEIFS(C2:C11,B2:B11,"West",D2:D11,">=3")

12. Total North Sales
=SUMIF(B2:B11,"North",C2:C11)

13. Average North Sales
=AVERAGEIF(B2:B11,"North",C2:C11)

14. Average West Sales
=AVERAGEIF(B2:B11,"West",C2:C11)

15. Count Eligible Employees
=COUNTIF(G2:G11,"Eligible")

16. Count Excellent Employees
=COUNTIF(F2:F11,"Excellent")

17. Eligibility Percentage
=COUNTIF(G2:G11,"Eligible")/COUNTA(G2:G11)*100

18. South + West + Sales >= 50000 + Rating >= 4
=COUNTIFS(C2:C11,">=50000",B2:B11,"South",D2:D11,">=4")+COUNTIFS(C2:C11,">=50000",B2:B11,"West",D2:D11,">=4")
