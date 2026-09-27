# 📊 Excel Business Analytics & Performance Modeling Projects

A comprehensive collection of hands-on Excel business analytics case studies, demonstrating dynamic logical functions, statistical aggregation, conditional formatting, and multi-criteria workforce modeling.

---

## 📑 Table of Contents
1. [Project Overview](#-project-overview)
2. [Case Study 1: Customer Care Services Data Management](#-case-study-1-customer-care-services-data-management)
3. [Case Study 2: Cinema Ticket Sales & Revenue Contribution Analysis](#-case-study-2-cinema-ticket-sales--revenue-contribution-analysis)
4. [Case Study 3: Employee Multi-Criteria Bonus Eligibility & Workforce Diagnostics](#-case-study-3-employee-multi-criteria-bonus-eligibility--workforce-diagnostics)
5. [Formulas & Technical Reference](#-skills--functions-demonstrated)
6. [Author & Portfolio Details](#-author)

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

## 🎯 Objectives
1. **Determine Eligibility Dynamically:** Create five dedicated columns (`Bonus A` to `Bonus E`) and evaluate qualification using nested Excel logical functions (`IF`, `AND`, `OR`). Eligible employees receive the bonus label; ineligible employees are left blank (`""`).
2. **Visual Hierarchy with Conditional Formatting:** Apply unique color fills to each bonus column so qualifying employees can be identified at a glance.
3. **Ineligibility & Performance Gap Analysis:** Compute the percentage of employees who failed to meet the criteria for each bonus tier across the total workforce.

---

## 🧮 Bonus Criteria & Excel Formulas

Formulas are structured for row `2` and copied down through row `121`:

### 1. Bonus A
* **Criteria:** Awarded to employees who produced **2,800 units or more**.
* **Formula:**
  ```excel
  =IF(F2>=2800, "Bonus A", "")
  ```

### 2. Bonus B
* **Criteria:** Awarded to employees with **fewer than 4 days of absence** AND **at least 2,600 units produced**.
* **Formula:**
  ```excel
  =IF(AND(C2<4, F2>=2600), "Bonus B", "")
  ```

### 3. Bonus C
* **Criteria:** Awarded to employees with a **machinery wear coefficient of 0.40 or lower** OR who **worked 245 hours or more**.
* **Formula:**
  ```excel
  =IF(OR(E2<=0.40, D2>=245), "Bonus C", "")
  ```

### 4. Bonus D
* **Criteria:** Requires a **wear coefficient of 0.30 or lower AND fewer than 5 days absent**, OR **at least 200 hours worked**.
* **Formula:**
  ```excel
  =IF(OR(AND(E2<=0.30, C2<5), D2>=200), "Bonus D", "")
  ```

### 5. Bonus E
* **Criteria:** Specific to **female employees** who meet either of the following conditions:
  * Machinery wear coefficient of **0.30 or lower**, OR
  * Produced **2,800 units or more**.
* **Formula:**
  ```excel
  =IF(AND(B2="F", OR(E2<=0.30, F2>=2800)), "Bonus E", "")
  ```

---

## 🎨 Conditional Formatting Rules

To make the eligibility status dynamically distinguishable, conditional formatting rules (using "Format only cells that contain" or specific text matching) were applied across `G2:K121`:

| Bonus Column | Cell Value Rule | Fill Color | Visual Purpose |
| :--- | :--- | :--- | :--- |
| **Bonus A (Col G)** | Cell Value equal to `"Bonus A"` | 🟩 **Light Green** | Highlights top production output |
| **Bonus B (Col H)** | Cell Value equal to `"Bonus B"` | 🟦 **Blue** | Highlights high output with low absenteeism |
| **Bonus C (Col I)** | Cell Value equal to `"Bonus C"` | 🟨 **Yellow** | Highlights machine efficiency or long hours |
| **Bonus D (Col J)** | Cell Value equal to `"Bonus D"` | 🟥 **Red** | Highlights low-wear attendance or 200+ hours |
| **Bonus E (Col K)** | Cell Value equal to `"Bonus E"` | 🟪 **Purple** | Highlights qualified female workforce members |

---

## 📈 Ineligibility Rate Analysis

The total ineligibility rate for each bonus category is calculated at the summary row (`Row 122`) by counting blank cells in each bonus column divided by the total number of employee entries:

$$\text{Ineligibility Percentage} = \frac{\text{COUNTBLANK}(X2:X121)}{\text{COUNTA}(\$A\$2:\$A\$121)}$$

### Summary Table

| Metric | Bonus A | Bonus B | Bonus C | Bonus D | Bonus E |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Excel Formula (Row 122)** | `=COUNTBLANK(G2:G121)/COUNTA($A$2:$A$121)` | `=COUNTBLANK(H2:H121)/COUNTA($A$2:$A$121)` | `=COUNTBLANK(I2:I121)/COUNTA($A$2:$A$121)` | `=COUNTBLANK(J2:J121)/COUNTA($A$2:$A$121)` | `=COUNTBLANK(K2:K121)/COUNTA($A$2:$A$121)` |
| **Ineligibility Rate (%)** | **74.17%** | **68.33%** | **53.33%** | **38.33%** | **78.33%** |
| **Eligibility Rate (%)** | **25.83%** | **31.67%** | **46.67%** | **61.67%** | **21.67%** |

### Key Takeaways:
1. **Most Accessible Incentive:** **Bonus D** has the lowest ineligibility rate (**38.33%**), meaning over $61\%$ of the workforce qualified due to the accessible 200-hour threshold.
2. **Most Stringent Production Target:** **Bonus A** disqualified **74.17%** of employees, showing that 2,800+ units is an elite production milestone achieved by only ~26% of workers.
3. **Gender-Specific Bonus E:** Ineligible rate stands at **78.33%**, representing both non-female employees as well as female employees who did not satisfy the secondary criteria.

---

## 🖼️ Process Screenshots

### Step 1: Bonus Eligibility Formulas Applied
Formulas entered across columns G through K returning designated bonus strings or blank values.

<img width="1908" height="941" alt="1 1" src="https://github.com/user-attachments/assets/cffda5d6-5ad6-48e1-b26e-529a6b264265" />


---

### Step 2: Dynamic Conditional Formatting Applied
Color-coded rules active across each bonus column for rapid visual filtering.

<img width="1917" height="933" alt="1" src="https://github.com/user-attachments/assets/d125a259-29bb-4791-bee5-ffe87cd2a68f" />

---

### Step 3: Ineligibility Percentage Calculation
Summary row calculations computing non-qualification rates using `COUNTBLANK` and `COUNTA`.

<img width="1918" height="932" alt="2" src="https://github.com/user-attachments/assets/bff19fd3-4e09-494f-8fee-c87aa266e70c" />

---

## 💡 Skills & Functions Demonstrated
- **Logical Functions:** `IF`, `AND`, `OR`, `IFS`
- **Statistical Aggregation:** `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNTIF`, `COUNTBLANK`, `COUNTA`
- **Cell Addressing:** Strict absolute referencing (`$A$1`) vs. relative referencing (`A1`)
- **Data Hygiene:** Power Query text delimiter transformation and header promotion
- **Data Visualization:** Rule-driven conditional formatting and contribution formatting (`0.0%`)

---

## 👤 Author
- **LinkedIn:** [LinkedIn Profile URL](https://www.linkedin.com/in/ritik-panwar-01a67a24b/?lipi=urn%3Ali%3Apage%3Ad_flagship3_feed%3BzPr0zrYGSguyR5pNTkBafQ%3D%3D)
- **GitHub:** [GitHub Profile URL](https://github.com/ritikpanwar10/)
- **Portfolio:** [Portfolio / Project Link](https://ritikpanwar10.github.io/ritikpanwar.github.io/)
