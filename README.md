# 🍕 Pizza Sales Dashboard (Excel)

An interactive Excel dashboard that analyzes 2015 pizza sales. It covers the full workflow: data preparation, pivot tables, KPI cards, charts, and slicers.

---

## 📌 Table of Contents
1. [Project Overview](#-project-overview)
2. [Business Questions](#-business-questions)
3. [Dataset](#-dataset)
4. [Tools Used](#-tools-used)
5. [Step 1: Data Preparation](#-step-1-data-preparation)
6. [Step 2: Pivot Tables](#-step-2-pivot-tables)
7. [Step 3: Building the Dashboard](#-step-3-building-the-dashboard)
8. [Key Insights](#-key-insights)
9. [Repository Structure](#-repository-structure)
10. [How to Use](#-how-to-use)
11. [Future Improvements](#-future-improvements)

---

## 📖 Project Overview

This project turns raw pizza order data into a one-page, interactive dashboard that helps a pizza business understand **what sells, when it sells, and how much it earns**. Everything was built in **Microsoft Excel** using Power Query / data cleaning, PivotTables, PivotCharts, and Slicers.

### Dashboard Preview
![Pizza Sales Dashboard](images/dashboard.png)
<img width="792" height="302" alt="dashboard1" src="https://github.com/user-attachments/assets/da4c9485-e15c-4df2-85ac-5777033066ee" />

---

## ❓ Business Questions

- What are the total sales, orders, customers, and quantity sold?
- Which pizza categories generate the most revenue?
- Which pizza sizes drive sales?
- What are the best-selling and least-selling pizzas?
- How do orders change month to month?
- Which days of the week are the busiest?

---

## 🗂 Dataset

| Field | Description |
|---|---|
| `order_id` | Unique ID of each order |
| `order_date` | Date of the order (2015) |
| `Customer Name` | Name of the customer |
| `pizza_name` | Name of the pizza |
| `pizza_category` | Chicken, Classic, Supreme, Veggie |
| `pizza_size` | S, M, L, XL |
| `quantity` | Number of pizzas ordered |
| `Sales` | Revenue of the order line (price × quantity) |

> The source columns include Arabic and English headers.

---

## 🛠 Tools Used

- **Microsoft Excel**: data cleaning, formulas, PivotTables, PivotCharts, Slicers
- **Excel Timeline**: date filtering
- **Excel Shapes and Icons**: dashboard layout and design

---

## 🧹 Step 1: Data Preparation

Before analysis, the raw data was cleaned and structured in Excel:

1. **Import the data** into Excel and convert the range into an **Excel Table** (`Ctrl + T`) so formulas and pivots update automatically.
2. **Check data quality**
   - Remove duplicate rows (`Data → Remove Duplicates`)
   - Find blank cells and handle missing values
   - Trim extra spaces (`TRIM`) and fix inconsistent text
3. **Fix data types**: set dates as *Date*, quantities and prices as *Number*, and text fields as *Text*.
4. **Create helper columns**
   - `Month`: `=TEXT([@order_date],"mm-mmmm")` (the number prefix keeps months in calendar order)
   - `Weekday`: `=TEXT([@order_date],"dddd")`
   - `Sales`: `=[@quantity]*[@price]` (if not already in the data)
5. **Rename columns** consistently (for example `pizza_size`, `pizza_category`) so slicers and pivots are easy to read.

---

## 📊 Step 2: Pivot Tables

Each pivot table feeds one KPI or chart on the dashboard.

| # | Pivot Table | Rows | Values | Used For |
|---|---|---|---|---|
| 1 | Sales by Category | `pizza_category` | Sum of `Sales` | Pie chart |
| 2 | Sales by Size | `pizza_size` | Sum of `Sales` | Donut chart |
| 3 | Orders by Month | `Month` | Count of `order_id` | Line chart |
| 4 | Orders by Weekday | `Weekday` | Count of `order_id` | Column chart |
| 5 | Top Pizzas | `pizza_name` | Sum of `quantity` (Top N filter, descending) | Top best-selling bar chart |
| 6 | Bottom Pizzas | `pizza_name` | Sum of `quantity` (Bottom N filter) | Least-selling bar chart |
| 7 | KPI Cards | n/a | Sum of `Sales`, Sum of `quantity`, Count of `order_id`, Count of `Customer Name` | KPI cards |

### Pivot Tables Preview
![Pivot Tables](images/pivot-tables.png)
<img width="792" height="316" alt="pivot" src="https://github.com/user-attachments/assets/3c0f05b8-9cba-42bb-9ed4-a571771ba6d9" />

**Key pivot tips used**
- **Value Filters → Top 10 / Bottom 10** to rank pizzas
- Sorted values so the charts read from largest to smallest
- Number formats set for currency and thousands separators

---

## 🎨 Step 3: Building the Dashboard

1. **Create a new sheet** named `Dashboard` and hide gridlines (`View → uncheck Gridlines`).
2. **Design the layout**
   - Dark green header banner with the title and pizza icons
   - KPI strip under the header
   - Charts in a grid with rounded-corner cards
   - Filter panel on the right
3. **Add KPI cards** linked to the KPI pivot cells: **Total Sales, Total Customers, Total Orders, Total Quantity**.
4. **Insert PivotCharts** from the pivot tables:
   - 🥧 Pie chart: Total Sales per Category
   - 🍩 Donut chart: Sales Distribution by Pizza Size
   - 📈 Line chart: Monthly Order Volume
   - 📊 Column chart: Weekly Order Distribution
   - 📉 Bar charts: Top 5 Best-Selling and Bottom 5 Least-Selling Pizzas
5. **Format the charts**: consistent green palette, data labels on, chart titles and axis clutter removed.
6. **Add Slicers and a Timeline**
   - Slicers: `pizza_size`, `pizza_category`
   - Timeline: `order_date`
7. **Connect the filters to all pivots** (`Slicer → Report Connections`) so one click updates every chart and KPI.
8. **Final polish**: align objects, group shapes, lock the layout, and hide the pivot/helper sheets if needed.

---

## 💡 Key Insights

| Metric | Result |
|---|---|
| **Total Sales** | **$8,436.05** |
| **Total Orders** | **499** |
| **Total Quantity** | **511** pizzas |
| **Top category by revenue** | Veggie (~28%), followed by Classic (~27%) |
| **Lowest category by revenue** | Chicken (~21%) |
| **Top size by revenue** | Large (~48% of sales) |
| **Smallest size** | XL (~3% of sales) |
| **Busiest day** | Friday (91 orders) |
| **Slowest day** | Wednesday (48 orders) |
| **Peak month** | June (54 orders) |
| **Best-selling pizza** | The Classic Deluxe Pizza (30 units) |
| **Least-selling pizza** | The Brie Carre Pizza (2 units) |

---

## 📁 Repository Structure

```
pizza-sales-excel-dashboard/
│
├── data/
│   └── pizza_sales_raw.xlsx          # original dataset
│
├── dashboard/
│   └── Pizza_Sales_Dashboard.xlsx    # cleaned data, pivots, and dashboard
│
├── images/
│   ├── dashboard.png                 # dashboard screenshot
│   └── pivot-tables.png              # pivot tables screenshot
│
└── README.md
```

---

## ▶️ How to Use

1. Download or clone the repository:
   ```bash
   git clone https://github.com/<your-username>/pizza-sales-excel-dashboard.git
   ```
2. Open `dashboard/Pizza_Sales_Dashboard.xlsx` in **Microsoft Excel** (2016 or later recommended).
3. Go to the **Dashboard** sheet.
4. Use the **slicers** (size, category) and the **timeline** (date) to filter the data.
5. If you update the data, go to `Data → Refresh All`.

---

## 🚀 Future Improvements

- Add a **distinct customer count** (Data Model) so Total Customers is not the same as Total Orders
- Add **Average Order Value** as a KPI
- Add **revenue by month** and **top 5 by revenue** charts
- Use the same axis scale on Top 5 and Bottom 5 charts
- Replace the 3D pie chart with a flat donut or bar chart for more accurate comparison
- Rebuild in **Power BI** for a web-shareable version

---

## 👤 Author

**<Your Name>**
- LinkedIn: <your-link>
- GitHub: <your-link>

⭐ If you found this project useful, please give it a star!
