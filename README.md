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

### 📐 Applied Formulas

#### Helpdesk Column (`F2`):
```excel
=IF(OR(Country="South Korea", Country="Japan"), "China Helpdesk",
 IF(OR(Country="Sri Lanka", Country="New Zealand", Country="Wales"), "India Helpdesk",
 "USA Helpdesk"))
```

*Alternative (Excel 2019 / 365 `IFS` syntax):*
```excel
=IFS(
    OR(B2="South Korea", B2="Japan"), "China Helpdesk",
    OR(B2="Sri Lanka", B2="New Zealand", B2="Wales"), "India Helpdesk",
    TRUE, "USA Helpdesk"
)
```

#### Priority Column (`G2`):
```excel
=IF(AND(Reason="Damaged Product", Purchase_Value > AVERAGE($E$2:$E$500)), "High", "")
```

### 📸 Visual Documentation

#### 1. Ingested Customer Dataset
![Customer Care Data Import](screenshots/task1/01_data_import_preview.png)
*Figure 1.1: Raw customer support records imported and normalized into a structured table.*

#### 2. Helpdesk Routing Logic
![Helpdesk Assignment](screenshots/task1/02_helpdesk_routing_formula.png)
*Figure 1.2: Dynamic routing formula segmenting records into China, India, and USA helpdesks.*

#### 3. Priority Escalation Output
![Priority Escalation Output](screenshots/task1/03_priority_escalation_output.png)
*Figure 1.3: High-priority flags populated exclusively for above-average damaged item tickets.*

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

### 📐 Applied Formulas

| Metric | Formula | Description |
| :--- | :--- | :--- |
| **City Revenue** | `=C2 * $C$18` | Multiplies city ticket volume by fixed unit ticket price |
| **Total Tickets** | `=SUM(C2:C15)` | Nationwide ticket sales sum |
| **Total Revenue** | `=SUM(D2:D15)` | Total commercial gross across all cinema hubs |
| **Max Tickets** | `=MAX(C2:C15)` | Identifies top-volume cinema market |
| **Min Revenue** | `=MIN(D2:D15)` | Identifies lowest-grossing market |
| **Average Revenue** | `=AVERAGE(D2:D15)` | Mean performance benchmark across cities |
| **% Revenue Share**| `=D2 / $D$16` | City revenue divided by absolute total revenue |

### 📸 Visual Documentation

#### 1. City Sales & Revenue Performance Table
![City Sales Table](screenshots/task2/01_city_sales_table.png)
*Figure 2.1: City-by-city sales volumes, price models, and percentage revenue contributions.*

#### 2. Descriptive Statistical Summary
![Statistical Summary](screenshots/task2/02_summary_statistics.png)
*Figure 2.2: Overall market statistics displaying ticket and revenue boundaries.*

#### 3. Regional Takings Distribution
![Revenue Share Chart](screenshots/task2/03_revenue_share_chart.png)
*Figure 2.3: Visual breakdown of market share contributions across operational cities.*

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
