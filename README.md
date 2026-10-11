# Sales-Data-Analysis-in-MS-EXCEL-
covers data modelling, data cleaning, KPI design, and business insights across sales, profitability, customers, products, employees, inventory, expenses and targets.
Project Objectives
Build a clean star-schema data model for BI reporting (Excel / Power BI / SQL).
Identify and fix data quality issues in the raw data.
Track revenue, profit and margin over time and by channel, product, customer segment and region.
Compare employee performance vs. sales targets.
Monitor inventory risk (stock-outs, reorder levels) and operating expenses.
🗂️ Data Model
Sheet	Type	Rows	Description
Fact_Sales	Fact	30,250	Order lines: date, customer, product, employee, channel, payment method, status, quantity, price, discount, gross/net sales, cost, profit, margin
Fact_Inventory	Fact	5,000	Stock movements per product/warehouse: opening, in, out, reorder level, closing
Fact_Expenses	Fact	4,000	Departmental expenses by type with approval status
Fact_Targets	Fact	1,000	Monthly sales and customer-count targets per employee
Dim_Customers	Dimension	2,500	Segment, state, city, region, join date, rating
Dim_Products	Dimension	250	Product name, category, brand, unit cost
Dim_Employees	Dimension	80	Department, job level, region
Dim_Date	Dimension	1,096	Calendar 2024-01-01 → 2026-12-31 (year, quarter, month, week, weekday, weekend flag)

Relationships

Dim_Customers (Customer_ID) ──┐
Dim_Products  (Product_ID)  ──┼──< Fact_Sales >── Dim_Employees (Employee_ID)
Dim_Date      (Date)        ──┘          │
                                         └── Order_Date → Dim_Date
Dim_Products  ──< Fact_Inventory      Dim_Employees ──< Fact_Expenses / Fact_Targets
Calculated columns in Fact_Sales
Column	Formula
Gross_Sales	Quantity × Unit_Price
Discount_Amount	Gross_Sales × Discount
Net_Sales	Gross_Sales − Discount_Amount
UNT COST	VLOOKUP(Product_ID, Dim_Products, Unit_Cost)
TOTAL COST	Quantity × Unit Cost
PROFIT	Net_Sales − TOTAL COST
Margin	PROFIT ÷ Gross_Sales × 100 (note: based on gross, not net, sales)
🧹 Data Quality Findings & Cleaning Steps
Issue	Where	Count	Fix
Exact duplicate rows (repeated Order_IDs)	Fact_Sales	250	Remove duplicates
Missing Customer_ID	Fact_Sales	101	Flag as "Unknown" or exclude from customer analysis
Missing Discount	Fact_Sales	120	Treat as 0% (or impute); verify with source
Inconsistent text (" online", "transfer", "sales")	Channel, Payment_Method, Department	~182 / 120 / many	TRIM + proper-case standardisation
Gross_Sales ≠ Quantity × Unit_Price	Fact_Sales	80	Recalculate
Missing City / Brand	Dim_Customers / Dim_Products	70 / 20	Fill as "Unknown"
"Transfer" vs "Bank Transfer" overlap	Payment_Method	–	Confirm if they are the same method and merge
Duplicate Month + Employee targets	Fact_Targets	150	Keep one target per employee per month
Negative closing stock	Fact_Inventory	242	Investigate stock-out / data-entry errors
Future-dated orders (after the analysis date)	Fact_Sales	2,234	Confirm whether data is forecast/synthetic
Margin column definition	Fact_Sales	–	Documented as profit ÷ gross sales; consider adding profit ÷ net sales
📊 Key Metrics (raw data, all order statuses)
KPI	Value
Total orders (rows)	30,250
Gross sales	₦195.5 B
Discounts given	₦12.4 B
Net sales	₦183.1 B
Total cost	₦140.4 B
Profit	₦42.7 B
Profit margin (on net sales)	23.3 %
Operating expenses	₦1.69 B

Figures are computed on the raw file before removing duplicates and including Pending, Returned and Cancelled orders. Re-run after cleaning, and filter to Completed orders for realised revenue.

🔍 Key Insights

Order status

Only 72 % of orders are Completed (₦132.1 B net sales). Pending (₦21.6 B), Returned (₦16.3 B) and Cancelled (₦13.1 B) represent roughly 28 % of order value, which is a notable leakage and a risk to revenue recognition.

Trend

Sales are flat year over year (2024 ₦60.8 B → 2025 ₦61.8 B → 2026 ₦60.5 B net) with no strong seasonality; growth is a business opportunity.

Channels & payments

Walk-in leads (₦37.8 B), but all five channels sit within ~7 % of each other (Social Media, Corporate Sales, Online, Referral).
Payment methods are evenly distributed (Transfer, Cash, Bank Transfer, Card, POS).

Products

Top categories: Computers (₦47.0 B), Storage (₦40.8 B), Networking (₦32.9 B).
Best sellers: Flash Drive, Laptop, Mini PC. Samsung and Canon lead by brand.
Margins are very uniform (22.7 %–23.8 %) across categories and brands, so pricing/cost is not differentiating profitability.

Customers

Retail is the largest segment (₦74.5 B, ~41 %), followed by Corporate (₦40.0 B) and SME (₦36.6 B).
Region performance is balanced (South South, South East and South West lead; North Central is lowest).
Top states by net sales: Kano, Edo, Abia, Imo, Delta.

Discounts

765 orders are loss-making, all of them with 15 % or 20 % discounts. Capping discounts at 10 % (or requiring approval above it) protects margin.

Targets

Of the employee-months that have a target, ~97 % have actual completed sales above target. Targets may be set too low and should be recalibrated.

Expenses

₦1.69 B total; spread evenly across 7 expense types and 5 departments. ₦0.31 B is still Pending and ₦0.12 B Rejected approval.

Inventory

579 of 5,000 stock records (~12 %) closed below their reorder level, including 242 with negative stock, so replenishment and stock-recording controls need attention.
🛠️ Tools & Techniques
Excel: Power Query (cleaning), data model, VLOOKUP/XLOOKUP, SUMIFS, PivotTables, slicers, charts
Power BI / DAX (optional extension): relationships, measures, dashboards
Python (pandas) for validation and the figures above
Suggested DAX measures
DAX
Net Sales   = SUM(Fact_Sales[Net_Sales])
Total Cost  = SUM(Fact_Sales[TOTAL COST])
Profit      = [Net Sales] - [Total Cost]
Profit Margin % = DIVIDE([Profit], [Net Sales])
Avg Order Value = DIVIDE([Net Sales], DISTINCTCOUNT(Fact_Sales[Order_ID]))
Net Sales YoY % = DIVIDE([Net Sales] - CALCULATE([Net Sales], SAMEPERIODLASTYEAR(Dim_Date[Date])),
                         CALCULATE([Net Sales], SAMEPERIODLASTYEAR(Dim_Date[Date])))
Target Attainment % = DIVIDE([Net Sales], SUM(Fact_Targets[Sales_Target]))
Cancellation Rate = DIVIDE(COUNTROWS(FILTER(Fact_Sales, Fact_Sales[Order_Status]="Cancelled")), COUNTROWS(Fact_Sales))
Python validation script
python
import pandas as pd

FILE = "BUSINESS_INTELLIGENCE-SALES_DATA_ANALYSIS.xlsx"
sheets = pd.read_excel(FILE, sheet_name=None)
sales, prod, cust = sheets["Fact_Sales"], sheets["Dim_Products"], sheets["Dim_Customers"]

# --- Cleaning ---
sales = sales.drop_duplicates()
for col in ["Channel", "Payment_Method"]:
    sales[col] = sales[col].str.strip().str.title()
sales["Discount"] = sales["Discount"].fillna(0)
sales["Customer_ID"] = sales["Customer_ID"].fillna("UNKNOWN")
sales["Gross_Sales"] = sales["Quantity"] * sales["Unit_Price"]
sales["Net_Sales"] = sales["Gross_Sales"] * (1 - sales["Discount"])
sales = sales.merge(prod[["Product_ID", "Category", "Brand", "Unit_Cost"]], on="Product_ID", how="left")
sales["Total_Cost"] = sales["Quantity"] * sales["Unit_Cost"]
sales["Profit"] = sales["Net_Sales"] - sales["Total_Cost"]

# --- KPIs ---
done = sales[sales["Order_Status"] == "Completed"]
print("Net sales:", done["Net_Sales"].sum())
print("Profit   :", done["Profit"].sum())
print("Margin % :", done["Profit"].sum() / done["Net_Sales"].sum() * 100)

# --- Breakdowns ---
print(done.groupby("Category")[["Net_Sales", "Profit"]].sum().sort_values("Net_Sales", ascending=False))
print(done.groupby(done["Order_Date"].dt.to_period("Q"))["Net_Sales"].sum())
print(sales[sales["Profit"] < 0].groupby("Discount").size())   # loss-making discount tiers
💡 Recommendations
Cap discounts at 10 % (approval required above) to eliminate loss-making orders.
Reduce Pending / Returned / Cancelled orders (≈28 % of value): tighten order confirmation and returns policy.
Recalibrate sales targets, as ~97 % attainment suggests they are too easy.
Fix inventory controls: investigate negative stock and automate reorder alerts.
Grow the top-performing segments (Retail, Corporate) and diversify beyond flat YoY sales.
Enforce data validation at entry (dropdowns for channel / payment method; mandatory Customer_ID and Discount).
📁 Repository Structure
├── data/
│   └── BUSINESS_INTELLIGENCE-SALES_DATA_ANALYSIS.xlsx
├── scripts/
│   └── analysis.py
├── dashboard/            # screenshots / Power BI file (optional)
└── README.md
▶️ How to Use
Clone the repo and open the workbook in Excel (or load it into Power BI).
Review the Dim_* and Fact_* sheets and the relationships above.
Run python scripts/analysis.py to reproduce the KPIs (pip install pandas openpyxl).
👤 Author

Your Name – GERALD B. VERALLO
