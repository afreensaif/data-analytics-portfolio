# 🛒 Amazon Sales Automated Report Generator

A Python automation script that transforms raw Amazon order data into a fully formatted, executive-ready Excel KPI report — replacing hours of manual weekly analysis with a single command.

---

## 📌 The Problem This Solves

Sales and operations teams often spend **3–5 hours every week** manually pulling Amazon order data, calculating KPIs in Excel, and building summary reports for leadership.

This script automates that entire workflow:
- Reads raw order data (128,000+ rows)
- Cleans and processes it automatically
- Calculates 10+ KPIs across revenue, fulfilment, and cancellations
- Outputs a branded, multi-tab Excel report in seconds

**Result: What took hours now takes one command.**

---

## 📊 What the Report Contains

The generated Excel file has 6 formatted sheets:

| Sheet | Contents |
|---|---|
| 📊 Executive Summary | 10 top-line KPIs: revenue, AOV, delivery rate, cancellation rate, B2B vs B2C |
| 📈 Monthly Trend | Revenue, orders, units, AOV, and MoM growth % by month + bar chart |
| 🏷️ Category Performance | Revenue share, AOV, and order volume by product category |
| 📍 Top States | Top 10 states by revenue |
| 🚚 Fulfilment Breakdown | Amazon vs Merchant fulfilment comparison |
| ❌ Cancellation Analysis | Cancellation rate by category — flags problem areas |

---

## 🗂️ Dataset

**Source:** [E-Commerce Sales Dataset — Kaggle](https://www.kaggle.com/datasets/thedevastator/unlock-profits-with-e-commerce-sales-data)

**File used:** `Amazon Sale Report.csv`

| Property | Detail |
|---|---|
| Rows | 128,975 orders |
| Date range | March 2022 – June 2022 |
| Market | Amazon India (INR) |
| Categories | Sets, Kurta, Western Dress, Top, Ethnic Dress, Blouse, Bottom, Saree |

**Key columns used:**

```
Order ID | Date | Status | Category | Qty | Amount |
Fulfilment | ship-state | ship-city | B2B | Courier Status
```

---

## ⚙️ How It Works

The script runs in 4 sequential steps — logged in the terminal as it runs:

**Step 1 — Load & Clean**
- Reads the CSV with pandas
- Parses dates, standardises amounts, strips whitespace
- Classifies each order as Delivered / Cancelled / Returned

**Step 2 — Calculate KPIs**
- Filters to delivered orders for revenue metrics
- Computes monthly aggregations, category breakdowns, state rankings
- Calculates MoM growth, cancellation rates, fulfilment splits

**Step 3 — Write to Excel**
- Uses `pandas` ExcelWriter to write each analysis to a separate sheet
- Structures data with appropriate headers and row positioning

**Step 4 — Apply Formatting**
- Uses `openpyxl` to apply branded styling: teal headers, alternating row colours, borders
- Auto-sizes columns
- Adds a bar chart to the Monthly Trend sheet

---

## 🚀 How to Run

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/data-analytics-portfolio.git
cd data-analytics-portfolio/project-1-amazon-automation
```

### 2. Install dependencies
```bash
pip install pandas openpyxl
```

### 3. Add your data file
Place `Amazon_Sale_Report.csv` in the same folder as the script.

> Download from [Kaggle](https://www.kaggle.com/datasets/thedevastator/unlock-profits-with-e-commerce-sales-data) → `Amazon Sale Report.csv`

### 4. Run
```bash
python amazon_sales_report.py
```

### 5. Open your report
The output file `Amazon_Sales_KPI_Report.xlsx` will appear in the same directory.

---

## 🛠️ Tech Stack

| Tool | Purpose |
|---|---|
| `pandas` | Data loading, cleaning, aggregation |
| `openpyxl` | Excel writing, styling, charting |
| Python 3.8+ | Core language |

---

## 📁 File Structure

```
project-1-amazon-automation/
│
├── amazon_sales_report.py       # Main automation script
├── Amazon_Sale_Report.csv       # Raw input data (download from Kaggle)
├── Amazon_Sales_KPI_Report.xlsx # Generated output report
└── README.md                    # This file
```

---

## 💡 Key Insights From This Dataset

Running the script on the Amazon India dataset surfaces some immediately actionable findings:

- **14.2% cancellation rate** — higher than industry average; worth investigating by category
- **Sets and Kurtas** account for the majority of revenue — concentration risk
- **Merchant fulfilment** vs Amazon fulfilment shows meaningful differences in order volume
- **Maharashtra and Karnataka** are the highest-revenue states — useful for geo-targeting

These are the kinds of insights this script is designed to surface automatically, every week, without manual effort.

---

## 🔄 Adapting to Your Own Data

This script is designed to be reusable. To run it on your own Amazon sales export:

1. Replace `INPUT_FILE` at the top of the script with your filename
2. Ensure your CSV has these columns (standard Amazon Seller Central export format):
   `Order ID`, `Date`, `Status`, `Category`, `Qty`, `Amount`, `Fulfilment`, `ship-state`, `B2B`
3. Run the script — all KPIs and formatting apply automatically

---

## 👩‍💻 About

Built by **Afreen Saif** — Senior Data Analyst specialising in Python automation, BigQuery pipelines, and Power BI dashboards across real estate, e-commerce, and fintech.

🔗 [LinkedIn](https://www.linkedin.com/in/afreen-saif) | 📧 afreen.saif95@gmail.com
