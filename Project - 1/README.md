<div align="center">

# 📊 Excel Project - 1

### 🧮 Excel Formula Practice • Data Analysis • Real-World Calculations

<p>
<img src="https://img.shields.io/badge/Microsoft_Excel-Workbook-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white">
<img src="https://img.shields.io/badge/Formulas-Practice-4472C4?style=for-the-badge">
<img src="https://img.shields.io/badge/Logical_Functions-IF%20%7C%20AND%20%7C%20OR-6C63FF?style=for-the-badge">
<img src="https://img.shields.io/badge/Date_Functions-YEAR%20%7C%20TODAY-16A085?style=for-the-badge">
</p>

<p>
<img src="https://img.shields.io/badge/Text_Functions-UPPER%20%7C%20LOWER-E67E22?style=for-the-badge">
<img src="https://img.shields.io/badge/Rounding-ROUND%20%7C%20CEILING%20%7C%20FLOOR-8E44AD?style=for-the-badge">
<img src="https://img.shields.io/badge/Absolute_Reference-%24I%243-FF6B35?style=for-the-badge">
</p>

<br>

**📥 Enter Data → 🧮 Apply Formula → 🔍 Calculate Results → 📊 Analyze Data → 🚀 Practice Excel**

</div>

---

## ✨ Project Overview

**Excel Project - 1** is a practical Microsoft Excel workbook created to practice commonly used Excel formulas through structured, real-world style datasets.

The workbook contains **3 different sheets**, each focused on a specific type of Excel calculation:

| 🧩 Sheet | 📌 Purpose |
|---|---|
| 🎓 **Sheet1** | Student marks, averages, grades, qualification status, and text functions |
| 💰 **Sheet2** | Sales amounts, discounts, product eligibility, and rounding calculations |
| 👨‍💼 **Sheet3** | Employee salary, joining dates, service duration, and salary increment |

> 💡 **Main Excel File:** `Project - 1.xlsx`

This project is useful for **Excel practice, assignments, practical examinations, formula learning, and beginner-level data analysis**. 🚀

---

# 🗂️ Workbook Structure

```text
📦 Project - 1.xlsx
│
├── 🎓 Sheet1
│   └── Student Performance & Formula Practice
│
├── 💰 Sheet2
│   └── Sales Analysis & Rounding Functions
│
└── 👨‍💼 Sheet3
    └── Employee Salary & Date Calculations
```

---

# 🎓 Sheet 1 - Student Performance

The first sheet contains student records with subject marks, enrollment dates, and calculated results.

### 📋 Columns

| Column | Purpose |
|---|---|
| `Student ID` | Unique student identifier |
| `Name` | Student name |
| `Math` | Mathematics marks |
| `Science` | Science marks |
| `English` | English marks |
| `Enrollment Date` | Student enrollment date |
| `Avg` | Average marks |
| `Grade` | Grade based on average |
| `Status` | Qualification status |
| `Name (Upper)` | Name converted to uppercase |
| `Name (Lower)` | Name converted to lowercase |

### 🧮 Formulas Practiced

#### 📊 Average

```excel
=AVERAGE(C2:E2)
```

Calculates the average marks of Math, Science, and English.

#### 🏆 Grade

```excel
=IF(G2>=90,"A",IF(G2>=75,"B",IF(G2>=60,"C",IF(G2>=50,"D","F"))))
```

Assigns a grade based on the student's average marks.

```text
90 or above  → A
75–89        → B
60–74        → C
50–59        → D
Below 50     → F
```

#### ✅ Qualification Status

```excel
=IF(AND(C2>75,D2>75),"Qualified","Not Qualified")
```

Uses `AND` to check whether both Math and Science marks are greater than 75.

#### 🔠 Uppercase Text

```excel
=UPPER(B2)
```

Converts the student name to uppercase.

#### 🔡 Lowercase Text

```excel
=LOWER(B2)
```

Converts the student name to lowercase.

---

# 💰 Sheet 2 - Sales Analysis

The second sheet contains sales records for different products, regions, and salespersons.

It focuses on **conditional calculations, discounts, logical functions, and rounding values**.

### 📋 Columns

| Column | Purpose |
|---|---|
| `Sales ID` | Unique sales identifier |
| `Product` | Product sold |
| `Region` | Sales region |
| `Salesperson` | Person responsible for the sale |
| `Amount` | Sales amount |
| `Date` | Sales date |
| `Discount` | Discount calculated from amount |
| `IF + OR` | Product eligibility check |
| `Round of` | Rounded sales amount |
| `Ceiling` | Value rounded upward |
| `Floor` | Value rounded downward |

### 💸 Discount Calculation

```excel
=IF(E2>=40000,E2*10%,IF(E2>=20000,E2*5%,0))
```

The discount is calculated according to the sales amount:

```text
💰 Amount >= 40,000 → 10% Discount
💰 Amount >= 20,000 → 5% Discount
💰 Amount < 20,000  → 0 Discount
```

### 🔀 IF + OR

```excel
=IF(OR(B2="Laptop",B2="keyboard"),"Eligible","Not Eligible")
```

Checks whether the product matches one of the specified products.

> 💡 This also provides practice with combining `IF` and `OR` in one formula.

### 🔢 ROUND

```excel
=ROUND(E2,-3)
```

Rounds the sales amount to the nearest thousand.

### ⬆️ CEILING

```excel
=CEILING(E2,1000)
```

Rounds the sales amount upward to the nearest multiple of 1,000.

### ⬇️ FLOOR

```excel
=FLOOR(E2,1000)
```

Rounds the sales amount downward to the nearest multiple of 1,000.

---

# 👨‍💼 Sheet 3 - Employee Analysis

The third sheet contains employee information and focuses on **date calculations, service duration, salary increments, and absolute cell references**.

### 📋 Columns

| Column | Purpose |
|---|---|
| `Employee ID` | Unique employee identifier |
| `Name` | Employee name |
| `Department` | Employee department |
| `Salary` | Current salary |
| `Joining Date` | Employee joining date |
| `Joing Year` | Year extracted from joining date |
| `Days of Service` | Number of service days |
| `Increment (10%)` | Calculated salary increment |
| `Salary Hike (Absolute Reference)` | Fixed salary hike percentage |

### 📅 Joining Year

```excel
=YEAR(E2)
```

Extracts the year from the employee's joining date.

### ⏱️ Days of Service

```excel
=IF(E2>TODAY(),0,TODAY()-E2)
```

Calculates the number of days between the joining date and today's date.

If the joining date is in the future, the formula returns `0`.

### 💵 Salary Increment

```excel
=D2*$I$3
```

Calculates the salary increment using the percentage stored in cell `I3`.

---

# 🔒 Absolute Cell Reference

One of the important concepts practiced in **Sheet3** is the absolute reference:

```excel
$I$3
```

The salary hike percentage is stored in cell `I3`.

When the formula is copied down:

```text
D2 × $I$3
D3 × $I$3
D4 × $I$3
D5 × $I$3
...
```

The employee salary reference changes, but `$I$3` remains fixed.

### 💡 Why `$I$3`?

```text
$I$3
 ↑  ↑
 |  └── Fixed Row
 └───── Fixed Column
```

This is useful when one fixed percentage or value needs to be applied to many rows.

---

# 🧠 Excel Concepts Covered

This project provides practical practice with:

```text
📊 AVERAGE
🧠 IF
🔗 AND
🔀 OR
🔠 UPPER
🔡 LOWER
📅 YEAR
⏰ TODAY
🔢 ROUND
⬆️ CEILING
⬇️ FLOOR
🔒 Absolute Cell Reference
📋 Formula Fill Down
🧮 Nested IF
📈 Conditional Calculations
📅 Date Calculations
```

---

# 🧪 Recommended Practice Flow

```text
📥 Enter / Review Data
        ↓
🔍 Understand Columns
        ↓
🧮 Apply Basic Formula
        ↓
🧠 Add Logical Conditions
        ↓
📅 Work with Dates
        ↓
🔢 Apply Rounding
        ↓
🔒 Practice Absolute Reference
        ↓
📋 Copy Formulas Down
        ↓
📊 Check Results
        ↓
🚀 Create Your Own Formula
```

> 💡 **Best way to learn Excel:** Change the input value, observe the result, then change the formula and compare the output.

---

# ▶️ How to Use

## 💻 Requirements

- 🟢 **Microsoft Excel**
- 📄 **`Project - 1.xlsx`**
- 🧠 Basic understanding of cells, rows, columns, and formulas

## 🚀 Steps

**1️⃣ Open the workbook**

Open:

```text
Project - 1.xlsx
```

**2️⃣ Select a sheet**

Start with `Sheet1`, then practice `Sheet2` and `Sheet3`.

**3️⃣ Select a calculated cell**

Click a formula cell and check the formula bar to understand how the calculation works.

**4️⃣ Change input values**

Modify marks, sales amounts, salaries, or dates and observe how the results change.

**5️⃣ Practice copying formulas**

Drag formulas down to apply the same calculation to additional records.

**6️⃣ Practice absolute reference**

In `Sheet3`, observe how `$I$3` remains fixed when the salary increment formula is copied down.

---

# 📌 Quick Project Summary

| 📌 Property | 💻 Details |
|---|---|
| 📄 Workbook | `Project - 1.xlsx` |
| 📊 Sheets | 3 |
| 🎓 Student Analysis | Included |
| 💰 Sales Analysis | Included |
| 👨‍💼 Employee Analysis | Included |
| 🧠 Logical Functions | `IF`, `AND`, `OR` |
| 📊 Statistical Function | `AVERAGE` |
| 🔤 Text Functions | `UPPER`, `LOWER` |
| 📅 Date Functions | `YEAR`, `TODAY` |
| 🔢 Rounding Functions | `ROUND`, `CEILING`, `FLOOR` |
| 🔒 Absolute Reference | `$I$3` |
| 🧮 Nested Formula | Included |
| 📋 Formula Practice | Included |

---

# 🎯 Learning Outcomes

After completing this project, you will have practical experience with:

- 📊 Calculating averages and grades
- 🧠 Using `IF` with multiple conditions
- 🔗 Combining conditions with `AND`
- 🔀 Using `OR` for multiple criteria
- 🔠 Converting text using `UPPER` and `LOWER`
- 📅 Extracting years from dates
- ⏰ Calculating service duration using `TODAY`
- 💰 Applying discount conditions
- 🔢 Rounding numerical values
- 🔒 Understanding absolute cell references
- 📋 Copying formulas across multiple records
- 🧮 Building formulas for real-world datasets

---

## 👨‍💻 Author

<div align="center">

### 🌟 Ayush Donga 🌟

**B.Sc IT Student | Aspiring AI/ML Engineer 🤖**

`📊 Excel` · `💻 Data Analysis` · `🐍 Python` · `🤖 AI/ML`

**📄 Project:** `Project - 1.xlsx`

</div>

---

<div align="center">

### ⭐ Learn by practicing. Build by applying. 🚀

**📊 Excel • 🧮 Formulas • 📈 Data Analysis • 💻 Technology**

</div>
