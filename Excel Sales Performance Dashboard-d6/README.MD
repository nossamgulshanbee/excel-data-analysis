# 📊 Excel Day 6 — Sales Performance Analysis Dashboard

## 📌 Project Overview

This project focuses on **advanced Excel data analysis and dashboard creation** using a sales dataset.

The goal was to transform raw sales data into meaningful business insights using **PivotTables, PivotCharts, KPI calculations, percentage analysis, running totals, monthly comparisons, and dashboard design**.

This project represents the practical application of the Excel concepts learned during **Day 6 of my Data Analytics learning journey**.

---

## 🎯 Objectives

* Analyze sales performance across different dimensions.
* Calculate important business KPIs.
* Analyze monthly sales trends.
* Calculate percentage contribution to total sales.
* Create running totals.
* Identify top-performing products.
* Identify top-performing employees.
* Calculate month-to-month sales changes.
* Build an interactive and professional sales dashboard.

---

## 🗂️ Dataset

The dataset contains sales transactions with the following columns:

| Column   | Description                       |
| -------- | --------------------------------- |
| Date     | Date of the sales transaction     |
| Employee | Employee responsible for the sale |
| Region   | Sales region                      |
| Product  | Product sold                      |
| Category | Product category                  |
| Sales    | Revenue generated                 |
| Cost     | Cost associated with the sale     |
| Profit   | Profit generated                  |

### Dataset Size

* **51 transactions**
* **5 employees**
* **4 regions**
* **8 products**
* **3 categories**
* **2 months of sales data**

---

# 📚 Concepts Covered

## 1. Percentage of Grand Total

Used PivotTable:

**Show Values As → % of Grand Total**

This shows how much each region contributes to the overall sales.

Example:

```text
Region Sales Contribution
-------------------------
East   → 27.59%
North  → 21.37%
South  → 27.62%
West   → 23.42%
```

### Purpose

Helps identify which regions contribute the largest share of total sales.

---

## 2. Running Total

Used PivotTable:

**Show Values As → Running Total In**

The running total continuously adds sales as the dates progress.

Example:

```text
Date        Sales       Running Total
01-01-2026  45,000      45,000
02-01-2026  30,000      75,000
03-01-2026  25,000      100,000
```

### Purpose

Used to understand cumulative sales growth over time.

---

## 3. Monthly Sales Analysis

A `Month` column was created from the Date column.

Formula:

```excel
=TEXT(A2,"mmm-yyyy")
```

The month field was then used in a PivotTable.

### Monthly Sales

| Month     |          Sales |
| --------- | -------------: |
| Jan-2026  |       ₹728,500 |
| Feb-2026  |       ₹503,200 |
| **Total** | **₹1,231,700** |

---

## 4. Month-to-Month Difference

Formula:

```excel
=February Sales-January Sales
```

Result:

```text
₹503,200 - ₹728,500
= -₹225,300
```

Therefore, sales decreased by:

**₹225,300**

---

## 5. Percentage Change

Formula:

```excel
=(February Sales-January Sales)/January Sales
```

Result:

**-30.93%**

### Insight

Sales decreased by approximately **30.93% from January to February 2026**.

---

# 🏆 Top 5 Products by Sales

A PivotTable was created using:

* Rows → Product
* Values → Sales
* Values → Profit

The products were sorted from largest to smallest based on sales.

| Rank | Product |    Sales |   Profit |
| ---: | ------- | -------: | -------: |
|    1 | Laptop  | ₹391,000 | ₹129,000 |
|    2 | Phone   | ₹273,000 |  ₹87,000 |
|    3 | Tablet  | ₹196,000 |  ₹64,000 |
|    4 | Monitor | ₹168,000 |  ₹55,500 |
|    5 | Printer | ₹117,000 |  ₹41,000 |

### Key Insight

**Laptop** was the highest-performing product by both sales and profit.

A **Bar Chart** was created to visualize the Top 5 Products by Sales.

---

# 👨‍💼 Top 5 Employees by Profit

A PivotTable was created using:

* Rows → Employee
* Values → Profit

| Rank | Employee |   Profit |
| ---: | -------- | -------: |
|    1 | Rahul    | ₹105,700 |
|    2 | Priya    |  ₹93,300 |
|    3 | Sneha    |  ₹93,200 |
|    4 | Kiran    |  ₹60,800 |
|    5 | Arjun    |  ₹60,700 |

### Key Insight

**Rahul** generated the highest profit with **₹105,700**.

A **Bar Chart** was created to visualize employee profit performance.

---

# 📊 KPI Analysis

Six important KPIs were calculated.

| KPI                   |      Value |
| --------------------- | ---------: |
| Total Sales           | ₹1,231,700 |
| Total Cost            |   ₹818,000 |
| Total Profit          |   ₹413,700 |
| Profit Margin         |     33.59% |
| Number of Orders      |         51 |
| Average Sales / Order | ₹24,150.98 |

---

## KPI Formulas

### Total Sales

```excel
=SUM(F2:F52)
```

### Total Cost

```excel
=SUM(G2:G52)
```

### Total Profit

```excel
=SUM(H2:H52)
```

### Profit Margin

```excel
=Total Profit/Total Sales*100
```

Result:

**33.59%**

### Number of Orders

```excel
=COUNTA(A2:A52)
```

Result:

**51**

### Average Sales per Order

```excel
=AVERAGE(F2:F52)
```

Result:

**₹24,150.98**

---

# 📈 Dashboard

A final **Sales Performance Dashboard** was created containing:

### KPI Cards

* Total Sales
* Total Cost
* Total Profit
* Profit Margin
* Orders
* Average Sales per Order

### Charts

* Monthly Sales Trend
* Top 5 Products by Sales
* Top 5 Employees by Profit

The dashboard combines the major findings into one visual summary.

---

# 🔍 Key Business Insights

### 1. Overall Performance

Total sales were:

**₹1,231,700**

with total profit of:

**₹413,700**

and a profit margin of:

**33.59%**

### 2. Regional Performance

**South** contributed the highest percentage of total sales at **27.62%**, closely followed by **East at 27.59%**.

### 3. Product Performance

**Laptop** was the highest-selling product with:

**₹391,000 sales**

and:

**₹129,000 profit**

### 4. Employee Performance

**Rahul** generated the highest profit:

**₹105,700**

### 5. Monthly Performance

January sales:

**₹728,500**

February sales:

**₹503,200**

Sales decreased by:

**₹225,300**

or approximately:

**30.93%**

---

# 🛠️ Excel Features Used

* Excel Tables
* PivotTables
* PivotCharts
* Slicers
* PivotTable Connections
* Percentage of Grand Total
* Running Total
* Sorting
* Monthly Analysis
* KPI Calculations
* Bar Charts
* Line Charts
* Number Formatting
* Dashboard Design

---

# 📁 Project Structure

```text
excel-sales-performance-dashboard/
│
├── Sales_Performance_Analysis.xlsx
│
└── README.md
```

---

# 🎓 Learning Outcome

Through this project, I learned how to move from **raw Excel data to meaningful business analysis**.

I practiced:

```text
Raw Data
   ↓
PivotTables
   ↓
Advanced Analysis
   ↓
KPIs
   ↓
Charts
   ↓
Dashboard
   ↓
Business Insights
```

This project strengthened my practical Excel skills for **Data Analytics and Business Intelligence**.

---

## 🚀 Skills Demonstrated

**Microsoft Excel | Data Analysis | PivotTables | PivotCharts | KPI Analysis | Dashboard Development | Data Visualization | Business Insights**

---

## 📌 Project

**Excel Day 6 — Sales Performance Analysis Dashboard**

**GitHub Repository:** `excel-sales-performance-dashboard`
