# 📊 Excel Business Analytics & Performance Modeling Projects

A comprehensive collection of hands-on Excel business analytics case studies, demonstrating dynamic logical functions, statistical aggregation, conditional formatting, and multi-criteria workforce modeling.

---

## 📑 Table of Contents
1. [Project Overview](#-project-overview)
2. [Case Study 1: Customer Care Services Data Management](#-case-study-1-customer-care-services-data-management)
3. [Case Study 2: Cinema Ticket Sales & Revenue Contribution Analysis](#-case-study-2-cinema-ticket-sales--revenue-contribution-analysis)
4. [Case Study 3: Employee Multi-Criteria Bonus Eligibility & Workforce Diagnostics](#-case-study-3-employee-multi-criteria-bonus-eligibility--workforce-diagnostics)
5. [Formulas & Technical Reference](#-formulas--technical-reference)
6. [Author & Portfolio Details](#-author--portfolio-details)

---

## 🎯 Project Overview

This repository documents three practical business intelligence and analytics assignments built in Microsoft Excel:
* **Global Customer Care Routing & Triage:** Multi-condition logic, nested routing, and threshold-based escalation modeling.
* **Cinema Sales & Market Share Distribution:** City-level revenue aggregation, statistical profiling (min/max/average), and absolute referencing for contribution shares.
* **Workforce Bonus Diagnostics:** Complex Boolean logic (`AND`/`OR`), dynamic conditional formatting, and non-qualification percentage metrics across multi-attribute operational KPIs.

---
## 🏢 Case Study 1: Customer Care Services Data Management

### 📌 Business Objective
A global enterprise routes customer support requests to regional helpdesk hubs (China, India, USA) and flags high-value damaged product claims for priority handling.

### 🛠️ Key Implementation Steps
1. **Data Ingestion:** Imported tab-delimited/raw data from `Customers.txt` using Excel's **Get & Transform (Power Query / Data Import)** wizard.
2. **Helpdesk Assignment Routing:**
   - **China Helpdesk:** Customers from *South Korea* or *Japan*.
   - **India Helpdesk:** Customers from *Sri Lanka*, *New Zealand*, or *Wales*.
   - **USA Helpdesk:** All other origin countries.
3. **High-Priority Escalation Flagging:**
   - Flagged as `"High"` if `Reason = "Damaged Product"` **AND** `Purchase Value > AVERAGE(all purchase values)`.
   - Remaining cells left blank (`""`).
## Implementation Details

The data transformation pipeline is executed using Power Query (M Formula Language):

### Step 1: Import Data
1. Open Excel and navigate to the **Data** tab.
2. Click **Get Data** $\rightarrow$ **From File** $\rightarrow$ **From Text/CSV**.
3. Select `customers.txt`, set delimiter detection to **Comma**, and click **Transform Data**.

<img width="1916" height="971" alt="1" src="https://github.com/user-attachments/assets/c7c52dd8-6859-4c46-a4b4-8c94bb3af05b" />
<img width="1187" height="875" alt="2" src="https://github.com/user-attachments/assets/b6004a29-dd1a-4b32-8bb0-3dd865238180" />

---

### Step 2: Regional Helpdesk Assignment
A custom column titled **`Helpdesk`** is created to route customer issues according to regional operations:

<img width="1712" height="863" alt="3" src="https://github.com/user-attachments/assets/b65355bb-b3e8-4f3d-b9c1-f40be5be8f19" />

#### Routing Matrix:
* **China Helpdesk**: Customers residing in South Korea or Japan.
* **India Helpdesk**: Customers residing in Sri Lanka, New Zealand, or Wales.
* **USA Helpdesk**: Customers from all other countries.

---

### Step 3: Priority Calculation
A custom column titled **`Priority`** flags high-value support requests needing immediate resolution:

<img width="1716" height="862" alt="4" src="https://github.com/user-attachments/assets/3f2be130-6f47-4e4a-8145-19f70648ec59" />

#### Priority Logic:
* **"High"**: Assigned only if the reason is `"Damaged item"` **and** the customer's purchase value exceeds the global average purchase value across all records.
* **Empty string (`""`)**: Applied to all remaining tickets.

---

### Step 4: Load Data
From the Power Query **Home** tab, click **Close & Load** to export the transformed data table directly into the active Excel workbook.

<img width="1727" height="865" alt="5" src="https://github.com/user-attachments/assets/f34ade18-b1f2-4029-b5d7-705a9699d9f1" />

---

## Final Output Structure
The final table includes all original attributes along with two engineered columns:
1. `Customer`
2. `Gender`
3. `Country`
4. `Reason`
5. `Purchase value`
6. **`Helpdesk`** (*China Helpdesk / India Helpdesk / USA Helpdesk*)
7. **`Priority`** (*"High" or blank*)

---

## 🎬 Case Study 2: Cinema Ticket Sales & Revenue Contribution Analysis

### 📌 Business Objective
Evaluate cinema ticket sales across regional cities, compute revenue metrics based on a benchmark ticket price (₹250 / $250), establish descriptive statistics, and calculate percentage revenue contributions.

### 🛠️ Key Implementation Steps
1. **Ticket & Revenue Aggregations:**
   - Total Tickets Sold per city and combined nationwide.
   - Total Earnings per city: `Tickets Sold * Average Ticket Price`.
2. **Statistical Profiling:**
   - Minimum, Maximum, and Average tickets sold and revenue generated across all cities using `MIN()`, `MAX()`, and `AVERAGE()`.
3. **Percentage Contribution:**
   - Computed each city's revenue share against overall takings using absolute cell referencing (`$F$Total`).

### Key Parameters:
* **Average Ticket Cost:** ₹250.00
* **Total Cities Analyzed:** 12

  ## 🎯 Objectives & Visual Breakdown

### Objective 1: Total Earnings Per City & Overall Totals
* **Total Tickets Sold Across All Cities:** `13,796,690`
* **Total Revenue Generated:** `₹ 3,449,172,500`
* **City-Level Revenue Calculation:** Multiplied each city's ticket volume by the base ticket price (`₹ 250`).

<img width="941" height="527" alt="1" src="https://github.com/user-attachments/assets/aa45fad9-463b-46fc-85b8-243f02ed8e87" />


---

### Objective 2: Statistical Insights (Min, Max, Average)
Calculated the core distribution benchmarks across both ticket sales volume and gross revenue:

| Metric | Tickets Sold | Revenue Generated (₹) | Top / Bottom City |
| :--- | :---: | :---: | :--- |
| **Maximum** | `3,249,788` | `₹ 812,447,000` | Mumbai |
| **Minimum** | `299,596` | `₹ 74,899,000` | Noida |
| **Average** | `1,149,724` | `₹ 287,431,042` | — |

<img width="1022" height="571" alt="2" src="https://github.com/user-attachments/assets/69ee4daf-6ae7-40d8-8c43-ce5bab3054f0" />

---

### Objective 3: Percentage Revenue Share (% Takings Per City)
Determines the relative financial contribution of each individual city toward total collections:

$$\text{\% Revenue} = \left( \frac{\text{City Revenue}}{\text{Total Revenue}} \right) \times 100$$

* **Top Contributor:** Mumbai at **23.6%** (₹81.24 Cr)
* **Top 3 Contributors:** Mumbai (23.6%), Delhi (13.8%), and Chennai (13.2%) together account for over **50.6%** of total revenue.
* **Lowest Contributor:** Noida at **2.2%** (₹7.49 Cr)

<img width="1055" height="563" alt="3" src="https://github.com/user-attachments/assets/2116d54d-2e98-43b3-bde8-18331838012d" />

---

## 🧮 Excel Formulas Reference

* **City Revenue (Column E):** `=D3 * $H$13` *(or `=D3 * 250`)*
* **Total Tickets Sold (Cell D16):** `=SUM(D3:D14)`
* **Total Revenue (Cell E16):** `=SUM(E3:E14)`
* **Max Tickets / Revenue:** `=MAX(D3:D14)` / `=MAX(E3:E14)`
* **Min Tickets / Revenue:** `=MIN(D3:D14)` / `=MIN(E3:E14)`
* **Average Tickets / Revenue:** `=AVERAGE(D3:D14)` / `=AVERAGE(E3:E14)`
* **% Revenue (Column F):** `=E3 / $E$16` *(Formatted as Percentage `0.0%`)*

---

## 🏭 Case Study 3: Employee Multi-Criteria Bonus Eligibility & Workforce Diagnostics

### 📌 Business Objective
Evaluate factory floor employees across five distinct production incentive programs (Bonuses A–E) combining output volume, attendance, shift hours, and machinery wear metrics. Highlight qualifiers using automated conditional formatting and quantify workforce ineligibility percentages.

### 📋 Bonus Criteria Matrix

| Bonus Program | Qualification Criteria | Excel Formula Syntax |
| :--- | :--- | :--- |
| **Bonus A** | Units Produced $\ge$ 2,800 | `=IF(Units>=2800, "Bonus A", "")` |
| **Bonus B** | Days Absent $<$ 4 **AND** Units $\ge$ 2,600 | `=IF(AND(DaysAbsent<4, Units>=2600), "Bonus B", "")` |
| **Bonus C** | Wear Coeff $\le$ 0.40 **OR** Hours Worked $\ge$ 245 | `=IF(OR(Wear<=0.40, Hours>=245), "Bonus C", "")` |
| **Bonus D** | (Wear $\le$ 0.30 **AND** Days Absent $<$ 5) **OR** Hours $\ge$ 200 | `=IF(OR(AND(Wear<=0.30, DaysAbsent<5), Hours>=200), "Bonus D", "")` |
| **Bonus E** | (Gender = "Female" **AND** Wear $\le$ 0.30) **OR** Units $\ge$ 2,800 | `=IF(OR(AND(Gender="Female", Wear<=0.30), Units>=2800), "Bonus E", "")` |

### 🎨 Dynamic Conditional Formatting
* Configured rules under **Home > Conditional Formatting > Highlight Cells Rules > Text that Contains**:
  - `Bonus A` ➔ Soft Mint Green fill with dark green text (`#D4EDDA`)
  - `Bonus B` ➔ Soft Cornflower Blue fill with dark blue text (`#CCE5FF`)
  - `Bonus C` ➔ Soft Warm Amber fill with dark amber text (`#FFF3CD`)
  - `Bonus D` ➔ Soft Lavender fill with dark purple text (`#E2D9F3`)
  - `Bonus E` ➔ Soft Coral/Rose fill with dark red text (`#F8D7DA`)

### 📊 Ineligibility Metric Calculations
To calculate the percentage of workers who **did not qualify** for a given bonus:

```excel
=COUNTBLANK(G2:G101) / COUNTA($A$2:$A$101)
```
*Or via non-blank check:*
```excel
=(COUNTA($A$2:$A$101) - COUNTIF(G2:G101, "Bonus A")) / COUNTA($A$2:$A$101)
```
*(Formatted as Percentage `0.0%`)*

### 📸 Visual Documentation

#### 1. Multi-Criteria Formulas & Logic Application
![Bonus Criteria Logic](screenshots/task3/01_bonus_criteria_logic.png)
*Figure 3.1: Formula implementation evaluating composite performance criteria.*

#### 2. Conditional Formatting Grid
![Conditional Formatting Grid](screenshots/task3/02_conditional_formatting_applied.png)
*Figure 3.2: Automated color highlights highlighting qualifying staff across programs.*

#### 3. Ineligibility Summary Dashboard
![Ineligibility Summary Table](screenshots/task3/03_ineligibility_summary_table.png)
*Figure 3.3: Diagnostic table displaying the percentage of workers ineligible per bonus.*

---

## 💡 Skills & Functions Demonstrated
- **Logical Functions:** `IF`, `AND`, `OR`, `IFS`
- **Statistical Aggregation:** `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNTIF`, `COUNTBLANK`, `COUNTA`
- **Cell Addressing:** Strict absolute referencing (`$A$1`) vs. relative referencing (`A1`)
- **Data Hygiene:** Power Query text delimiter transformation and header promotion
- **Data Visualization:** Rule-driven conditional formatting and contribution formatting (`0.0%`)

---

## 👤 Author
- **LinkedIn:** [Your LinkedIn Profile URL](https://www.linkedin.com/in/ritik-panwar-01a67a24b/?lipi=urn%3Ali%3Apage%3Ad_flagship3_feed%3BzPr0zrYGSguyR5pNTkBafQ%3D%3D)
- **GitHub:** [Your GitHub Profile URL](https://github.com/ritikpanwar10/)
- **Portfolio:** [Your Portfolio / Project Link](https://ritikpanwar10.github.io/ritikpanwar.github.io/)
