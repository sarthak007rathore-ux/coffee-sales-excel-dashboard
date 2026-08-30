# ☕ Coffee Sales Analysis & Interactive Excel Dashboard

An interactive **Coffee Sales Analysis Dashboard** built using Microsoft Excel to analyze sales performance across time, countries, coffee types, roast types, product sizes, and customers.

The project demonstrates practical Excel data-analysis techniques including **XLOOKUP, INDEX-MATCH, IF statements, calculated fields, Excel Tables, PivotTables, PivotCharts, and Slicers**.

<img width="773" height="396" alt="image" src="https://github.com/user-attachments/assets/1d0827b4-b4f6-4402-954d-61572a231139" />


---

## 📊 Project Overview

The objective of this project is to transform raw coffee sales transaction data into an interactive business dashboard that can be used to identify sales trends, top-performing markets, customer performance, and product-level patterns.

The workbook contains three main data components:

- **Orders** — transaction-level sales data
- **Customers** — customer master data
- **Products** — product master data

These datasets are combined using lookup functions to create a structured analysis dataset, which is then used to build PivotTables and an interactive dashboard.

---

## 🎯 Business Questions

The dashboard is designed to answer questions such as:

- How are sales changing over time?
- Which country generates the highest sales?
- Which customers contribute the most revenue?
- Which coffee types generate the most sales?
- Which roast type performs best?
- How does package size affect sales?
- How do loyalty-card customers compare with non-loyalty customers?
- How can sales data be filtered interactively?

---

## 🗂️ Workbook Structure

| Sheet | Description |
|---|---|
| `Dashboard` | Main interactive sales dashboard |
| `orders` | Transaction-level sales dataset with calculated fields |
| `customers` | Customer master dataset |
| `products` | Product master dataset |
| `TotalSales` | PivotTable showing monthly sales by coffee type |
| `Country BarChart` | PivotTable summarizing sales by country |
| `Top5Customers` | PivotTable showing the top customers by sales |

---

## 📈 Dashboard Features

### 1. Total Sales Over Time

A line chart showing monthly sales trends across the available period.

The analysis covers data from:

**January 2019 – August 2022**

---

### 2. Sales by Country

A bar chart comparing sales across:

- United States
- Ireland
- United Kingdom

The United States is the largest sales market in the dataset.

---

### 3. Top Customers

A bar chart displaying the highest-value customers based on total sales.

---

### 4. Interactive Slicers

The dashboard includes interactive slicers for:

- **Size**
- **Roast Type**
- **Loyalty Card**

These slicers allow users to dynamically filter the PivotTables and dashboard visualizations.

---

# 🧮 Excel Formulas & Techniques

## XLOOKUP

XLOOKUP was used to retrieve customer information from the customer master dataset.

### Customer Name

```excel
=XLOOKUP(C2,customers!$A$1:$A$1001,customers!$B$1:$B$1001,0)
```

This matches the Customer ID from the Orders table with the Customer ID in the Customers table and returns the corresponding Customer Name.

### Country

```excel
=XLOOKUP(C2,customers!$A$1:$A$1001,customers!$G$1:$G$1001,0)
```

This retrieves the customer's country.

### Loyalty Card

```excel
=XLOOKUP(Orders[[#This Row],[Customer ID]],customers!$A$1:$A$1001,customers!$I$1:$I$1001,,0)
```

This demonstrates the use of an Excel Table structured reference together with XLOOKUP.

---

## INDEX + MATCH

INDEX and MATCH were used to dynamically retrieve product attributes based on the Product ID and column header.

Example:

```excel
=INDEX(products!$A$1:$G$49,
MATCH(orders!$D2,products!$A$1:$A$49,0),
MATCH(orders!I$1,products!$A$1:$G$1,0))
```

The formula performs two lookups:

1. `MATCH` identifies the product row.
2. `MATCH` identifies the required product attribute column.
3. `INDEX` returns the value at the intersection.

This approach was used for:

- Coffee Type
- Roast Type
- Size
- Unit Price

---

## IF Statements

IF statements were used to convert coded values into readable categories and handle missing email values.

### Coffee Type

```excel
=IF(I2="Rob","Robusta",
IF(I2="Exc","Excelsa",
IF(I2="Ara","Arabica",
IF(I2="Lib","Liberica",""))))
```

### Roast Type

```excel
=IF(J2="M","Medium",
IF(J2="L","Light",
IF(J2="D","Dark","")))
```

---

## Sales Calculation

Sales were calculated using:

```excel
=L2*E2
```

Where:

- `L2` = Unit Price
- `E2` = Quantity

Therefore:

**Sales = Quantity × Unit Price**

---

# 📊 Key Dataset Statistics

| Metric | Value |
|---|---:|
| Order Records | 1,000 |
| Unique Order IDs | 957 |
| Unique Customers in Orders | 913 |
| Products | 48 |
| Countries | 3 |
| Coffee Types | 4 |
| Roast Types | 3 |
| Package Sizes | 4 |
| Total Quantity Sold | 3,551 |
| Total Sales | 45,134.255 |
| Data Start Date | 02-Jan-2019 |
| Data End Date | 19-Aug-2022 |

---

# 🔎 Key Findings

### Sales by Country

| Country | Sales |
|---|---:|
| United States | 35,638.885 |
| Ireland | 6,696.865 |
| United Kingdom | 2,798.505 |

The United States represents approximately **78.9% of total sales**.

### Sales by Coffee Type

| Coffee Type | Sales |
|---|---:|
| Excelsa | 12,306.440 |
| Liberica | 12,054.075 |
| Arabica | 11,768.495 |
| Robusta | 9,005.245 |

Excelsa generated the highest sales among the four coffee types.

### Sales by Roast Type

| Roast Type | Sales |
|---|---:|
| Light | 17,354.465 |
| Medium | 14,600.475 |
| Dark | 13,179.315 |

Light roast generated the highest sales.

### Sales by Size

| Size | Sales |
|---:|---:|
| 2.5 | 23,785.565 |
| 1.0 | 11,010.750 |
| 0.5 | 7,029.990 |
| 0.2 | 3,307.950 |

The 2.5-size products generated the highest sales in the dataset.

### Sales by Year

| Year | Sales |
|---|---:|
| 2019 | 12,187.165 |
| 2020 | 12,117.545 |
| 2021 | 13,766.110 |
| 2022 | 7,063.435 |

**Note:** 2022 contains data only through August 19 and therefore should not be directly compared with the full-year figures.

---

# 🛠️ Tools & Technologies

- **Microsoft Excel**
- XLOOKUP
- INDEX-MATCH
- IF
- Excel Tables
- PivotTables
- PivotCharts
- Slicers
- Data Cleaning
- Data Transformation
- Business Intelligence / Data Visualization

---

# 🧠 Skills Demonstrated

This project demonstrates practical skills in:

- Data cleaning and preparation
- Data lookup and integration
- Relational data concepts
- Excel formulas
- Data aggregation
- PivotTable analysis
- Interactive dashboard design
- Data visualization
- Customer analysis
- Sales trend analysis
- Business insight generation

---

# 📁 Project Workflow

```text
Raw Data
   │
   ├── Customers
   │
   ├── Products
   │
   └── Orders
        │
        ▼
Data Integration
(XLOOKUP + INDEX-MATCH)
        │
        ▼
Calculated Fields
(Sales + Category Names)
        │
        ▼
PivotTables
        │
        ▼
PivotCharts + Slicers
        │
        ▼
Interactive Dashboard
        │
        ▼
Business Insights
```

---
