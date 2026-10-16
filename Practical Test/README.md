<div align="center">

# 📊 E-Commerce Sales Dashboard

### 🛍️ Sales Analytics • Profitability • Order Performance • Business Insights

<p>
<img src="https://img.shields.io/badge/Microsoft_Excel-Sales_Analytics-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white">
<img src="https://img.shields.io/badge/Orders-799_Records-4472C4?style=for-the-badge">
<img src="https://img.shields.io/badge/Pivot_Tables-Business_Analysis-6C63FF?style=for-the-badge">
<img src="https://img.shields.io/badge/Dashboard-KPI_Reporting-FF6B35?style=for-the-badge">
</p>

<p>
<img src="https://img.shields.io/badge/Profit_Margin-23.12%25-16A085?style=for-the-badge">
<img src="https://img.shields.io/badge/Revenue_Analysis-2025-8E44AD?style=for-the-badge">
<img src="https://img.shields.io/badge/Excel_Functions-VLOOKUP%20%7C%20INDEX%20MATCH%20%7C%20SUMIF-E67E22?style=for-the-badge">
</p>

<br>

**📥 Order Data → 🧮 Formula Analysis → 🔄 Pivot Tables → 📊 Dashboard → 💡 Business Insights**

</div>

---

## ✨ Project Overview

The **E-Commerce Sales Dashboard** is a practical Microsoft Excel project designed to analyze sales, profit, order status, product performance, customer activity, payment methods, and monthly trends.

The workbook contains **799 order records** for 2025 and combines formula-based analysis, PivotTables, KPI reporting, and a dashboard-style summary.

| 🧩 Sheet | 📌 Purpose |
|---|---|
| 📋 **Data** | Stores order, customer, product, location, sales, profit, and payment details |
| 🧮 **Analysis** | Practices lookup formulas, conditional calculations, and key sales metrics |
| 🔄 **Pivot Table** | Summarizes sales by month, category, state, payment mode, status, product, and customer |
| 🎯 **Dashboard** | Presents sales KPIs, filters, charts, and business insights |

> 💡 **Main Excel File:** `Ayush_ExcelDashboard.xlsx`

This project is useful for **Excel formula practice, sales reporting, business analysis, PivotTables, interactive dashboards, and portfolio development**. 🚀

---

# 🗂️ Workbook Structure

```text
📦 Ayush_ExcelDashboard.xlsx
│
├── 📋 Data
│   └── E-Commerce Order Records
│
├── 🧮 Analysis
│   └── Lookup Formulas & Sales KPIs
│
├── 🔄 Pivot Table
│   └── Monthly, Product, State, Payment & Customer Analysis
│
└── 🎯 Dashboard
    └── Interactive Sales Dashboard & Key Insights
```

---

# 📋 Data Sheet

The `Data` sheet contains the source records used throughout the workbook.

### 🧾 Main Columns

| Column | Purpose |
|---|---|
| `Order ID` | Unique order identifier |
| `Order Date` | Date of the order |
| `Customer ID` | Customer identifier |
| `Customer Name` | Customer name |
| `City` | Customer city |
| `State` | Customer state |
| `Product ID` | Product identifier |
| `Product Name` | Product name |
| `Category` | Product category |
| `Quantity` | Units purchased |
| `Unit Price` | Price per unit |
| `Sales Amount` | Total sales value for the order |
| `Profit` | Profit generated |
| `Payment Mode` | Payment method |
| `Order Status` | Completed, Pending, Cancelled, or Returned |
| `Sales Category` | High Value, Medium Value, or Low Value |

### 🛍️ Product Categories

```text
💻 Electronics
🪑 Furniture
🏠 Home & Kitchen
👗 Fashion
💄 Beauty
🏏 Sports
📚 Stationery
```

### 💳 Payment Methods

```text
📱 UPI
💳 Credit Card
💳 Debit Card
🏦 Net Banking
💵 Cash
```

The data also supports comparisons across Indian states, cities, customers, products, and order dates.

---

# 🧮 Analysis Sheet

The `Analysis` sheet demonstrates how Excel formulas can retrieve information and calculate sales metrics.

### 🔍 VLOOKUP

**Purpose:** Find a product name using a Product ID.

```excel
=IFERROR(VLOOKUP(B2,data[[Product ID]:[Product Name]],2,FALSE),"Not Found")
```

Uses the Product ID entered in `B2` to return the matching product name.

### 🔗 INDEX + MATCH

**Purpose:** Find a customer name using a Customer ID.

```excel
=IFERROR(INDEX(data[Customer Name],MATCH(Analysis!E2,data[Customer ID],0)),"Not Found")
```

`MATCH` finds the matching customer row, while `INDEX` returns the corresponding customer name.

### 💰 SUMIF

**Purpose:** Calculate total sales for a selected product.

```excel
=SUMIF(data[Product Name],A7,data[Sales Amount])
```

Adds the sales amounts for records matching the product name in `A7`.

### 📍 COUNTIF

**Purpose:** Count orders from a selected state.

```excel
=COUNTIF(data[State],D7)
```

Counts the orders where the state matches the value in `D7`.

### 📌 Main KPIs

```text
💰 Total Sales
📈 Total Profit
📊 Average Sale
📦 Total Quantity Sold
🎯 Profit Margin
✅ Completed Orders
⬆️ Highest Sale
⬇️ Lowest Sale
```

### 📈 Profit Margin

```excel
=IFERROR(SUM(data[Profit])/SUM(data[Sales Amount]),0)
```

Calculates total profit as a proportion of total sales.

### 🔎 Highest and Lowest Sale

The Analysis sheet uses `MAX()` and `MIN()` to identify the highest and lowest sales amounts, with `XLOOKUP()` used to retrieve related product details.

> 💡 This section is useful for practising lookup functions, conditional error handling, aggregation, and formula references.

---

# 🔄 Pivot Table Analysis

The `Pivot Table` sheet summarizes sales performance from different business perspectives.

### 📅 Monthly Sales

The monthly summary compares sales from January through December 2025.

| Month | Total Sales |
|---|---:|
| January | ₹5.69 lakh |
| March | ₹6.83 lakh |
| July | ₹6.61 lakh |
| November | ₹6.22 lakh |
| December | ₹7.57 lakh |

🏆 **December** has the highest monthly sales in the workbook's summary, at approximately **₹7.57 lakh**.

### 🛍️ Orders by Product Category

| Category | Orders |
|---|---:|
| Stationery | 72 |
| Sports | 85 |
| Furniture | 91 |
| Beauty | 92 |
| Fashion | 118 |
| Home & Kitchen | 159 |
| Electronics | 182 |

🏆 **Electronics** has the highest number of orders in the category summary.

### 📍 Sales by State

The state summary compares total sales across locations including Gujarat, Maharashtra, Rajasthan, Karnataka, and other states.

| State | Total Sales |
|---|---:|
| Maharashtra | ₹12.54 lakh |
| Gujarat | ₹11.37 lakh |
| Karnataka | ₹8.94 lakh |
| Rajasthan | ₹7.54 lakh |
| Madhya Pradesh | ₹7.41 lakh |

🏆 **Maharashtra** leads the state-wise sales summary.

### 💳 Orders by Payment Mode

The PivotTable compares order counts for Cash, Credit Card, Debit Card, Net Banking, and UPI.

This analysis helps identify the payment methods used most frequently by customers.

### 📦 Order Status

The workbook summarizes the number of orders in each status:

```text
✅ Completed
🕒 Pending
❌ Cancelled
↩️ Returned
```

The status breakdown can help highlight completed orders and identify orders that may need follow-up.

### 🏆 Top 10 Products

The product summary ranks products by total sales.

**Study Desk** is the leading product in the displayed Top 10 product sales summary.

### 👥 Top 10 Customers

The customer summary ranks customers by total sales, helping identify customers with the highest recorded purchase value.

---

# 🎯 E-Commerce Sales Dashboard

The `Dashboard` sheet brings the main sales metrics and insights into one view.

### 📌 Dashboard KPIs

| KPI | Workbook Value |
|---|---:|
| 💰 Total Sales | ₹67.37 lakh |
| 📈 Total Profit | ₹15.58 lakh |
| 🎯 Profit Margin | 23.12% |
| 📦 Total Orders | 799 |
| 🧾 Average Order Value | ₹8,431 |
| ✅ Completed Orders | 588 |

### 🔍 Key Insights

| 💡 Area | Finding |
|---|---|
| 📍 Top Market | Maharashtra leads state-wise sales with approximately ₹12.54 lakh |
| 📅 Seasonality | December has the highest monthly sales in the PivotTable summary |
| 🏆 Top Product | Study Desk leads the displayed Top 10 product sales summary |
| 💰 Product Sales | Office Chair records approximately ₹10.72 lakh in sales |
| 📦 Order Health | Completed, pending, cancelled, and returned orders can be monitored separately |
| 💳 Payment Mix | The dashboard highlights UPI and Net Banking as popular payment methods |

### 🎛️ Interactive Analysis

The dashboard includes a filter area for exploring the sales data. Use the available filters and slicers in Excel to compare different views of the business.

> 💡 When the source data changes, refresh PivotTables and verify formula ranges so the dashboard reflects the latest records.

---

# 🧠 Excel Concepts Covered

This project provides practical experience with:

```text
🔎 VLOOKUP
🔗 INDEX + MATCH
🔍 XLOOKUP
🧮 SUMIF
🔢 COUNTIF
➕ SUM
📊 AVERAGE
⬆️ MAX
⬇️ MIN
🛡️ IFERROR
📋 Excel Table References
🔄 Pivot Tables
🎛️ Filters & Slicers
📈 Sales and Profit Analysis
🎯 KPI Dashboard Design
📅 Monthly Trend Analysis
```

---

# 🧪 Recommended Practice Flow

```text
📥 Review Order Data
        ↓
🔍 Understand Products & Customers
        ↓
🧮 Practise Lookup Formulas
        ↓
💰 Calculate Sales & Profit Metrics
        ↓
📅 Compare Monthly Sales
        ↓
🛍️ Analyse Product Categories
        ↓
📍 Compare State Performance
        ↓
💳 Review Payment Methods
        ↓
📦 Analyse Order Status
        ↓
🎛️ Explore Dashboard Filters
        ↓
💡 Summarize Business Insights
```

> 💡 **Best way to practise:** Change the lookup input values, compare formula outputs, refresh PivotTables after data updates, and observe how the dashboard insights change.

---

# ▶️ How to Use

## 💻 Requirements

- 🟢 **Microsoft Excel**
- 📄 **`Ayush_ExcelDashboard.xlsx`**
- 📊 Basic knowledge of Excel formulas and PivotTables

## 🚀 Steps

**1️⃣ Open the workbook**

Open `Ayush_ExcelDashboard.xlsx` in Microsoft Excel.

**2️⃣ Start with `Data`**

Review orders, customers, products, locations, sales amounts, profits, payment methods, and order status.

**3️⃣ Explore `Analysis`**

Test the VLOOKUP and INDEX + MATCH inputs, and review the calculated KPIs.

**4️⃣ Review `Pivot Table`**

Compare monthly sales, categories, states, payment modes, order status, top products, and top customers.

**5️⃣ Open `Dashboard`**

Review the KPI cards, filters, and key insights.

**6️⃣ Practise and validate**

If you update the source data, check formula ranges and refresh PivotTables so the analysis remains up to date.

---

# 📌 Quick Project Summary

| 📌 Property | 💻 Details |
|---|---|
| 📄 Workbook | `Ayush_ExcelDashboard.xlsx` |
| 📊 Sheets | 4 |
| 📦 Order Records | 799 |
| 💰 Sales Analysis | Included |
| 📈 Profit Analysis | Included |
| 🔎 Lookup Functions | Included |
| 🔄 Pivot Tables | Included |
| 📅 Monthly Sales Analysis | Included |
| 🛍️ Product Analysis | Included |
| 📍 State-wise Analysis | Included |
| 💳 Payment Mode Analysis | Included |
| 📦 Order Status Analysis | Included |
| 🎛️ Dashboard Filters | Included |
| 🎯 KPI Dashboard | Included |

---

# 🎯 Learning Outcomes

After completing this project, you will have practical experience with:

- 📋 Working with structured e-commerce order data
- 🔎 Retrieving information using lookup formulas
- 🧮 Calculating sales and profit metrics
- 📊 Summarizing data with `SUMIF()` and `COUNTIF()`
- 📅 Comparing monthly sales performance
- 🛍️ Identifying strong product categories and products
- 📍 Comparing sales across states
- 💳 Reviewing payment method preferences
- 📦 Monitoring order status
- 🔄 Creating and interpreting PivotTables
- 🎛️ Exploring dashboard filters and slicers
- 🎯 Presenting business metrics in an Excel dashboard

---

## 👨‍💻 Author

<div align="center">

### 🌟 Ayush Donga 🌟

**B.Sc IT Student | Aspiring AI/ML Engineer 🤖**

`📊 Excel` · `📈 Data Analysis` · `🐍 Python` · `🤖 AI/ML`

**📄 Project:** `Ayush_ExcelDashboard.xlsx`

</div>

---

<div align="center">

### ⭐ Explore the data. Understand the trends. Make informed decisions. 🚀

**🛍️ E-Commerce • 📊 Excel • 🔄 PivotTables • 🎯 Dashboards**

</div>
