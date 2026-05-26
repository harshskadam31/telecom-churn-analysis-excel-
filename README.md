# 📊 Telecom Customer Churn Analysis — Excel Dashboard
### Q2 2022 | Business Analytics & Product Management Case Study

---

## Project Description

Every quarter, telecom companies lose thousands of customers silently — and most never understand exactly why until it is too late. This project tackles that problem head-on.

Using real customer data from a California-based telecom company, I built a full end-to-end churn analysis dashboard entirely in Microsoft Excel — from raw data cleaning all the way to an executive-level interactive dashboard with slicers, KPI cards, and business recommendations.

The project is structured the way a Product Manager or Business Analyst would approach it in a real company: identify the problem, quantify the impact, segment the affected users, find the root causes, and recommend what to do about it. Every chart and number on the dashboard is tied to a specific business decision, not just displayed for the sake of it.

This is not a template fill-in. Every pivot table, formula, chart, and design decision was built from scratch — with the goal of answering one central question: **why are customers leaving, and what should the business do about it?**

---

## Objective

To build an executive-level churn analysis dashboard that translates raw customer data into actionable business insights — the kind a Product Manager or Business Analyst would present to leadership.

---

##  Dataset Description

The dataset contains **3 files** that work together:

| File | Rows | Description |
|------|------|-------------|
| `telecom_customer_churn.csv` | 7,043 | Main customer dataset — 38 columns covering demographics, subscribed services, contract details, billing, and churn information |
| `telecom_data_dictionary.csv` | 38 | Column-by-column definitions and value descriptions for the main dataset |
| `telecom_zipcode_population.csv` | 1,671 | Population estimates per zip code — used for geographic market penetration context |

**Source:** Maven Analytics Data Playground
**Time Period:** Q2 2022
**Geography:** California, USA

### What the Main Dataset Covers (38 Columns)

The 38 columns across the main dataset fall into 6 logical groups:

**Customer Demographics** — Customer ID, Gender, Age (range: 19–80), Married, Number of Dependents, City, Zip Code, Latitude, Longitude. Customers are spread across California with the highest concentrations in Los Angeles (293), San Diego (285), San Jose (112), Sacramento (108), and San Francisco (104).

**Acquisition & Tenure** — Number of Referrals, Tenure in Months (range: 1–72 months), Offer (None, Offer A through E). 3,877 customers had no promotional offer, while the rest were acquired through one of 5 offer types.

**Services Subscribed** — Phone Service, Multiple Lines, Internet Service, Internet Type (Fiber Optic: 3,035 / DSL: 1,652 / Cable: 830 / No Internet: 1,526), Online Security, Online Backup, Device Protection Plan, Premium Tech Support, Streaming TV, Streaming Movies, Streaming Music, Unlimited Data, Avg Monthly GB Download, Avg Monthly Long Distance Charges.

**Contract & Billing** — Contract type (Month-to-Month: 3,610 / Two Year: 1,883 / One Year: 1,550), Paperless Billing, Payment Method (Bank Withdrawal: 3,909 / Credit Card: 2,749 / Mailed Check: 385).

**Revenue** — Monthly Charge, Total Charges, Total Refunds, Total Extra Data Charges, Total Long Distance Charges, Total Revenue.

**Churn Details** — Customer Status (Churned / Stayed / Joined), Churn Category (Competitor / Dissatisfaction / Attitude / Price / Other), Churn Reason (specific reason text). Note: Churn Category and Churn Reason are only populated for customers who churned — the remaining 5,174 stayed or joined customers have blank values in these columns, which was handled during data preparation.

---

## 📐 How I Built This

### Step 1 — Data Preparation

Opened the raw CSV in Excel and reviewed all 38 columns against the data dictionary to understand every field before touching anything.

**What I changed and why:**

**Created a Tenure Bucket column** — The raw dataset had tenure as a continuous number (1 to 72 months). For analysis and slicing, I needed a categorical version that groups customers into meaningful lifecycle stages. I created a new column called "Tenure Bucket" using a nested IF formula:

```excel
=IF(K2<4,"Less than 3 Months",IF(K2<7,"3 to 6 Months",IF(K2<25,"6 Months to 2 Years","More than 2 Years")))
```

The buckets were chosen deliberately based on business logic — the first 3 months is the critical early retention window, 3–6 months is the early engagement phase, 6 months to 2 years is mid-lifecycle, and 2 years+ represents loyal long-term customers. This column became both a slicer and a pivot table dimension.

**Handled blank Churn Category values** — 5,174 customers who Stayed or Joined had no Churn Category since they did not churn. These appeared as "(blank)" in slicers, which looks unprofessional. To resolve this cleanly without corrupting the data, the blank slicer item was hidden via Slicer Settings so only actual churn categories (Competitor, Dissatisfaction, Attitude, Price, Other) appear when filtering.

Applied the formula down all 7,043 rows using the double-click fill handle method on the bottom-right corner of the formula cell, which auto-fills to the last row of data instantly.

---

### Step 2 — Pivot Table Construction

Built 9 separate pivot tables, each on its own dedicated named sheet. Every pivot table was built with **Customer Status (Churned / Stayed / Joined)** as the Column field — this creates a consistent side-by-side comparison structure across all tables.

| # | Sheet Name | Rows Field | Values Field | Business Question |
|---|-----------|-----------|-------------|-------------------|
| 1 | Churn Rate By Contract Type | Contract | Count of Customer Status | Are short-term customers leaving more? |
| 2 | Churn Rate By Tenure Bucket | Tenure Bucket | Count of Customer Status | When in the lifecycle does churn peak? |
| 3 | Churn Reason By Volume | Churn Category | Count of Churn Category | What reasons are customers citing? |
| 4 | Lost Revenue By Category | Churn Category | Sum of Total Revenue | What is each churn driver costing us? |
| 5 | Monthly Charge Churned vs Stayed | Customer Status | Average of Monthly Charge | Are we losing high or low value customers? |
| 6 | Customer By Internet Type | Internet Type | Count of Customer ID | What is our product mix by internet type? |
| 7 | Churn Rate By Internet Type | Internet Type | Count of Customer Status | Which internet product churns the most? |
| 8 | Revenue By Contract Type | Contract Type | Sum of Total Revenue | Which contract generates most revenue? |
| 9 | Overall Customer Status | Customer Status | Count of Phone Service | What is the net customer movement in Q2? |

For the **Lost Revenue** pivot table, I renamed the default "Sum of Total Revenue" column header to "Lost Revenue by Each Category" by clicking on the field in Values → Value Field Settings → Custom Name. This makes the table self-explanatory when referenced on the dashboard.

For the **Revenue by Contract Type** table, I changed the calculation from Sum to **Average** (Value Field Settings → Summarize by Average) to show revenue per customer rather than total — a more meaningful per-unit metric for comparing contract types.

---

### Step 3 — Chart Creation

Created charts directly from each pivot table by selecting the pivot table and inserting charts. Key chart type decisions and the reasoning behind each:

**Churn Reason by Volume → Horizontal Bar Chart (sorted largest to smallest)**
Sorted descending so the biggest churn driver is immediately obvious at the top. Competitor-related reasons dominate the top 3 bars. Horizontal bars work better than vertical here because the churn reason labels are long text.

**Churn by Tenure Bucket → Line Chart**
Shows the trend of churn declining as tenure increases far more clearly than bars would. The drop from 597 churned in the first 3 months down to 55 in the 5+ year bucket reads as a clear downward trend on a line.

**Churn by Internet Type → Combo Chart (Clustered Column + Line)**
Added a secondary axis showing Churn Rate % as a line overlaid on the raw count columns. This lets you see both the volume (Fiber Optic has the most churned customers) and the rate (Fiber Optic also has the highest churn rate at 40.7%) in a single chart without needing two separate visuals.

**Customer Count by Internet Type → Donut Chart**
Four clean segments with percentage labels. Donut works well here because the four internet type categories are easy to read at a glance and the center space was used to show the total (7,043).

**Average Monthly Charge Churned vs Stayed → 2-bar Column Chart**
Intentionally kept to just two bars — Churned ($73) and Stayed ($62). No clutter. The gap between the two bars tells the whole story: the company is losing its higher-paying customers.

**Revenue by Contract Type → Column Chart**
Shows Two Year contract customers generate the most revenue per customer ($3,432 average) despite not being the largest segment. Directly supports the recommendation to convert Month-to-Month customers.

For every chart, I cleaned up the following after creation:
- Removed chart border (Format Chart Area → No border)
- Removed gray plot area background (Format Plot Area → No fill)
- Removed pivot chart field buttons (PivotChart Analyze → Field Buttons → Hide All Field Buttons) — this is important for a clean dashboard look as field buttons make charts look like raw pivot outputs
- Removed legends where bar/data labels made them redundant
- Added data labels directly on bars and columns

---

### Step 4 — KPI Cards

Built 4 KPI cards for the top of the dashboard: Total Customers, Total Churned, Overall Churn Rate, and Total Revenue Lost.

**Method — Paste as Linked Picture:**

Standard text boxes in Excel are static. The Linked Picture method makes KPI numbers dynamic — they update automatically when a slicer is applied, just like the charts do.

Process for each KPI card:
1. On the relevant pivot table sheet, formatted the KPI number in a single cell — large font size (28–32), bold, colored to match the card theme
2. Selected and copied that cell (Ctrl+C)
3. Went to the Dashboard sheet, clicked the target position
4. Clicked the small dropdown arrow below the Paste button on the Home tab → selected **Linked Picture**
5. The number appeared as a floating object that looks like an image but behaves like a live formula

Resized each linked picture to fit within its KPI card box. Because it is linked, any slicer selection that changes the underlying pivot number automatically updates the KPI card on the dashboard.

---

### Step 5 — Slicers

Added 4 interactive slicers to the dashboard:
- Contract Type
- Internet Type
- Churn Category
- Tenure Bucket

**How I created them:**
Clicked inside any pivot table → PivotTable Analyze tab → Insert Slicer → selected the required field → clicked OK. Each slicer was created separately from its own pivot table, then moved to the Dashboard sheet.

**How I connected each slicer to all pivot tables:**
By default, a slicer only controls the pivot table it was created from. To make it filter every chart on the dashboard simultaneously: Right clicked the slicer → Report Connections → checked every pivot table in the list → clicked OK. This was done for all 4 slicers individually.

**How I aligned them professionally:**
Selected all 4 slicers together using Ctrl+Click on each → went to Shape Format tab in the ribbon → Align → Align Top (makes all slicer tops level) → then Distribute Horizontally (spaces them evenly). This creates a clean, evenly spaced slicer row without manually adjusting each one.

Removed the blank "(blank)" option from the Churn Category slicer via Slicer Settings → Hide items with no data, since non-churned customers have no churn category value.

---

### Step 6 — Dashboard Assembly

Created a dedicated Dashboard sheet and set it up as the canvas:
- Removed gridlines: View tab → unchecked Gridlines
- Set zoom to 85% for better spatial awareness while building
- Set background fill to a light off-white/gray tone across the full working area

**Layout structure built in this order:**

```
[ Title | Q2 2022 Subtitle | Last Updated ]
[ Slicers Row — Contract Type | Internet Type | Churn Category | Tenure Bucket | Reset Button ]
[ KPI Cards — Total Customers | Total Churned | Churn Rate | Revenue Lost | Key Insight Box ]
[ Row 1 — Churn by Contract Type | Churn Reason by Volume | Churn by Internet Type ]
[ Row 2 — Churn by Tenure Bucket | Avg Monthly Charge | Revenue by Contract Type ]
[ Row 3 — Lost Revenue by Category | Customer Donut | Overall Customer Status | Recommendations ]
```

Moved each chart from its pivot table sheet to the dashboard using Ctrl+X (cut) on the chart → Dashboard sheet → Ctrl+V (paste), then resized and aligned to fit the layout grid.

---

### Step 7 — Key Insights & Recommendations Panels

Added two text panels on the right side of the dashboard using formatted text boxes:

**Key Insights panel** — 5 bullet points summarizing the most critical findings from the data, written in plain business language. This forces any reader to understand the story without having to interpret the charts themselves.

**Top Recommendations panel** — 5 actionable recommendations with supporting icons, each directly tied to a specific data finding. This is what separates a reporting dashboard from a strategic business document — it answers "so what should we do?" not just "here is what happened."

**Potential Revenue at Risk box** — added at the bottom of the panel showing $3.68M — the forward-looking revenue risk if current churn trends continue into the next quarter. This reframes the lost revenue number from a historical fact into a future warning.

---

## 📊 Key Findings

- **26.5% churn rate** in Q2 2022 — nearly double the industry benchmark of ~15%
- **Competition is the #1 driver** — 841 customers left for a competitor, causing **$1.69M in lost revenue**
- **Month-to-Month customers churn at 45.8%** — nearly 1 in 2 customers on this plan left in a single quarter
- **Fiber Optic has the highest churn rate at 40.7%** — the premium product is underdelivering on value
- **597 customers churned in the first 3 months** — early retention is the most critical intervention window
- **Churned customers paid $73/month vs $62 for retained customers** — the business is losing its highest-paying customers first
- **Net customer movement is negative** — 454 joined but 1,869 churned, a net loss of 194 customers in Q2

---

## 💡 Business Recommendations

1. **Improve Early Retention** — Launch a structured onboarding program for new customers. 597 churned in the first 3 months, representing the single largest retention opportunity.

2. **Address Competition Directly** — 841 customers cited competitor offers and devices as their reason for leaving. A competitive device upgrade program or price-match policy could directly reduce this.

3. **Fix the Fiber Optic Experience** — 40.7% churn rate on the premium product is a serious product-market fit signal. Investigate pricing, speeds, and service quality specifically for Fiber Optic users.

4. **Convert Month-to-Month Customers to Long-Term Contracts** — Offer incentives (discounts, free add-ons) to migrate Month-to-Month customers to annual contracts. Churn drops from 45.8% to 7.4% with a One Year contract.

5. **Improve Customer Service Quality** — 220 customers cited "Attitude of support person" as their churn reason. This is an internal, directly fixable problem causing hundreds of avoidable exits per quarter.

6. **Promote Stickiness Add-ons** — Customers with Online Security churn at only 14.6% and Premium Tech Support at 15.2% — well below the 26.5% average. Push these during onboarding to increase product stickiness.

---

## 📁 File Structure

```
telecom-churn-analysis/
│
├── README.md
├── data/
│   ├── telecom_customer_churn.csv
│   ├── telecom_data_dictionary.csv
│   └── telecom_zipcode_population.csv
├── excel/
│   └── Telecom_Churn_Dashboard.xlsx
└── screenshots/
    ├── dashboard_overview.png
    ├── kpi_cards.png
    ├── churn_by_contract.png
    └── churn_reasons.png
```

---

*Data Source: Maven Analytics | Tools: Microsoft Excel | Date: May 2026*
