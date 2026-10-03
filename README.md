<!--
AUTHOR NOTES (invisible on GitHub, delete before publishing):
1. High Performers: the dashboard shows 74, which matches the Rating 5 count. Under the stated rule (Performance_Rating >= 4), Ratings 4 and 5 would give 206 + 74 = 280. Confirm which definition the KPI uses and update the formula or value.
2. "Average Salary by Job Level/Department" values (e.g., L1 = 31,232.9) look like monthly salary, while the KPI Average Annual Salary is 788,845. Confirm the field used and the currency.
3. The "Required vs Current Employees by Department" chart shows very large required values (e.g., 25,312 for Engineering), which suggests Required_Employees is summed across employee-level rows. Verify before quoting department-level required numbers.
4. Dashboard titles in the screenshots contain typos ("Workfoce", "Anlytics", "Compensation" spelled "Componsation"). Fix them in the workbook before capturing final screenshots.
5. The Salary Distribution chart in the screenshot is a pie chart of total salary by job level. It is described below as it appears.
-->

# 👥 Workforce Planning & Employee Performance Analytics

### HR Analytics | Workforce Planning | People Analytics | Microsoft Excel 2021

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel%202021-217346?style=flat&logo=microsoft-excel&logoColor=white)
![Domain](https://img.shields.io/badge/Domain-HR%20Analytics-blue?style=flat)
![Dashboards](https://img.shields.io/badge/Dashboards-3-orange?style=flat)

---

## 📌 Project Overview

This project simulates a real-world HR analytics environment for a **fictional multinational company**. Using **Microsoft Excel 2021**, it analyzes workforce capacity, staffing requirements, employee performance, productivity, attendance, overtime, compensation, and development needs.

The analysis is delivered through **three interactive Excel dashboards** that help HR teams understand whether the organization has the right workforce capacity, where staffing gaps exist, how employees are performing, and how attendance, workload, and compensation patterns can support workforce planning decisions.

---

## ❓ Problem Statement

HR leaders and workforce planners need reliable answers to questions such as:

- Do we have enough employees to meet workforce requirements?
- Which departments have staffing shortages?
- Which job roles require additional workforce capacity?
- How effectively is the current workforce being utilized?
- Which departments or job levels show stronger or weaker performance?
- How productive is the workforce, and how well are employees achieving their goals?
- Where are the training and development gaps?
- Are attendance and absenteeism patterns creating workforce capacity issues?
- Which departments have higher overtime requirements?
- How does compensation vary across departments and job levels?
- Where should HR focus workforce planning and employee development efforts?

This project addresses these questions through three dashboards covering **workforce planning**, **performance and productivity**, and **attendance, overtime and compensation**.

---

## 🎯 Business Objectives

- Analyze workforce headcount and distribution.
- Compare required workforce with current workforce.
- Identify staffing gaps and open positions.
- Measure workforce utilization.
- Analyze employee performance and productivity.
- Evaluate goal achievement and identify high-performing workforce segments.
- Analyze performance across departments and job levels.
- Evaluate training completion.
- Analyze attendance, absence, overtime, and workload patterns.
- Compare compensation across departments and job levels and understand salary distribution.
- Support workforce planning and resource allocation.
- Identify potential staffing and development priorities.

---

## 🔄 Project Workflow

| Step | Phase | Description |
|:---:|---|---|
| 1 | **Data Collection** | Assembled five HR tables: Employees, Performance, Attendance, Workforce_Planning, and Compensation. |
| 2 | **Data Cleaning & Preparation** | Validated and standardized records to prepare them for analysis (details below). |
| 3 | **Data Organization / Table Structure** | Structured the data as related tables linked through `Employee_ID` and a `Workforce_Key`. |
| 4 | **Excel Analysis & Calculations** | Built KPI and metric calculations using Excel formulas. |
| 5 | **PivotTables / PivotCharts / Dashboard Development** | Built three dashboards with KPI cards, filters/slicers, and business-focused charts. |
| 6 | **Insight Generation** | Interpreted the dashboard results to highlight workforce, performance, attendance, and compensation patterns. |

---

## 🗂️ Dataset & Data Model

The project contains **five main tables**.

### 1. Employees
`Employee_ID` · `Gender` · `Age` · `Department` · `Job_Role` · `Job_Level` · `Location` · `Hire_Date` · `Employment_Status`

### 2. Performance
`Employee_ID` · `Performance_Rating` · `Productivity_Score` · `Goal_Achievement_%` · `Manager_Rating` · `Training_Hours` · `Training_Completed` · `Promotion_Last_2_Years`

### 3. Attendance
`Employee_ID` · `Working_Days` · `Days_Present` · `Absence_Days` · `Late_Days` · `Overtime_Hours`

### 4. Workforce_Planning
`Department` · `Job_Role` · `Required_Employees` · `Current_Employees` · `Open_Positions` · `Workload_Level` · `Staffing_Gap`

### 5. Compensation
`Employee_ID` · `Monthly_Salary` · `Annual_Salary` · `Bonus` · `Salary_Hike_%` · `Benefits_Value`

### 🔗 Table Relationships

- **Employee-level tables** (`Employees`, `Performance`, `Attendance`, `Compensation`) are connected through **`Employee_ID`**.
- **`Workforce_Planning`** operates at the **Department + Job_Role level**, not at the employee level. It is connected to employee records through a **`Workforce_Key`** constructed from `Department + Job_Role`. It is not an `Employee_ID` relationship.

---

## 🧹 Data Cleaning & Preparation

Preparation focused on making the data reliable for formulas, PivotTables, charts, and dashboards:

- Checked employee records for duplicates
- Checked for missing or inconsistent values
- Standardized department names, job role names, and job levels
- Validated numeric fields
- Checked date fields such as `Hire_Date`
- Checked attendance, performance, and compensation values for consistency
- Created the `Workforce_Key` (`Department + Job_Role`) to link planning data to employee records
- Prepared the final tables for PivotTable, formula, chart, and dashboard analysis

---

## 🧮 Key Calculations / Metrics

| Metric | Calculation |
|---|---|
| **Total Employees** | Count of employees in the Employees table |
| **Required Workforce** | `SUM(Workforce_Planning[Required_Employees])` |
| **Staffing Gap** | `Required Workforce - Total Employees` |
| **Open Positions** | `SUM(Workforce_Planning[Open_Positions])` |
| **Workforce Utilization %** | `Total Employees ÷ Required Workforce × 100` |
| **Average Performance Rating** | `AVERAGE(Performance[Performance_Rating])` |
| **Average Productivity Score** | `AVERAGE(Performance[Productivity_Score])` |
| **Average Goal Achievement %** | `AVERAGE(Performance[Goal_Achievement_%])` |
| **High Performers** | Count of employees where `Performance_Rating >= 4` *(see author note above; confirm definition)* |
| **Training Completion %** | `Employees with Training_Completed = 1 ÷ Total Employees × 100` |
| **Attendance %** | `SUM(Days_Present) ÷ SUM(Working_Days) × 100` |
| **Average Absence Days** | `AVERAGE(Attendance[Absence_Days])` |
| **Average Overtime Hours** | `AVERAGE(Attendance[Overtime_Hours])` |
| **Average Annual Salary** | `AVERAGE(Compensation[Annual_Salary])` |
| **Average Salary Hike %** | `AVERAGE(Compensation[Salary_Hike_%])` |

> **Note:** Attendance % uses the aggregate `SUM(Days_Present) ÷ SUM(Working_Days)` rather than averaging individual employee attendance percentages. This weights each employee by the number of working days.

---

## 📊 Business KPIs

| Area | KPI |
|---|---|
| **Workforce Planning** | Total Employees · Required Workforce · Staffing Gap · Open Positions · Workforce Utilization % |
| **Performance** | Average Performance Rating · Average Productivity Score · Average Goal Achievement % · High Performers · Training Completion % |
| **Attendance & Compensation** | Attendance % · Average Absence Days · Average Overtime Hours · Average Annual Salary · Average Salary Hike % |

---

## 🖥️ Dashboard Overview

The project contains **three Excel dashboards**, each focused on a different HR business area:

1. **Workforce Overview & Planning**
2. **Employees Performance & Productivity**
3. **Attendance, Overtime & Compensation Analytics**

All three dashboards use Excel-based analysis, KPI cards, slicers/filters (Department, Job_Role, Location, Job_Level), and business-focused charts.

---

## 📈 Dashboard 1: Workforce Overview & Planning

**Business question:** *"Do we have the right number of employees in the right areas?"*

### KPI Cards

| KPI | Value |
|---|---:|
| Total Employees | 736 |
| Required Workforce | 815 |
| Staffing Gap | 79 |
| Open Positions | 79 |
| Workforce Utilization | 90.31% |

- **Total Employees:** current headcount in the Employees table.
- **Required Workforce:** the total required headcount across the planning data. It is the sum of `Required_Employees` across all Department + Job_Role records. It is **not** a per-department requirement.
- **Staffing Gap / Open Positions:** the shortfall between required and current headcount, and the positions currently open.
- **Workforce Utilization:** current headcount as a percentage of required workforce.

### Main Visuals

- Employees by Job Level
- Headcount by Location
- Headcount by Department
- Required vs Current Employees by Department
- Departments by Workload
- Workload Level by Department

### Supporting Data

| Job Level | Employees |
|---|---:|
| L1 - Entry | 146 |
| L2 - Associate | 234 |
| L3 - Senior | 213 |
| L4 - Lead | 101 |
| L5 - Manager | 42 |

| Location | Employees |
|---|---:|
| Bengaluru | 149 |
| Pune | 139 |
| Hyderabad | 111 |
| Delhi NCR | 102 |
| Mumbai | 86 |
| Kolkata | 81 |
| Chennai | 68 |

| Department | Employees |
|---|---:|
| Engineering | 226 |
| Customer Support | 131 |
| Sales | 102 |
| Operations | 88 |
| Data & Analytics | 60 |
| Marketing | 50 |
| Finance | 48 |
| Human Resources | 31 |

| Workload Level | Departments |
|---|---:|
| Medium | 17 |
| High | 10 |
| Low | 3 |
| Critical | 2 |

> Department-level required workforce values come from the individual `Required_Employees` records in the planning table, so they are not listed here.

### Key Insights

- The organization has **736 employees against a required workforce of 815**, a staffing gap of **79** and utilization of **90.31%**.
- **L2 - Associate** is the largest job-level group (234 employees), and **L5 - Manager** is the smallest (42).
- **Engineering** has the largest headcount (226).
- **Bengaluru** has the highest location headcount (149).
- The workload view shows a concentration of departments in the **Medium (17)** and **High (10)** categories, with 2 in the Critical category.

---

## 📈 Dashboard 2: Employees Performance & Productivity

**Business question:** *"How well is the workforce performing, and where are the performance gaps?"*

### KPI Cards

| KPI | Value |
|---|---:|
| Average Performance Rating | 3.3 |
| Average Productivity Score | 70.8 |
| Average Goal Achievement % | 87.40% |
| High Performers | 74 |
| Training Completion % | 77.58% |

### Main Visuals

1. Performance Rating Distribution
2. Average Productivity by Department
3. Goal Achievement % by Department
4. Performance vs Productivity
5. Performance by Job Level
6. Training Completion by Department
7. Average Manager Rating by Department

### Highlights

- **Performance Rating Distribution:** Rating 1: 29 · Rating 2: 88 · Rating 3: 339 · Rating 4: 206 · Rating 5: 74.
- **Productivity by Department:** average productivity ranges from approximately **69.2 to 72.6** across departments.
- **Goal Achievement:** Finance (89.72%) is among the higher departments, and Data & Analytics (83.51%) is lower.
- **Performance vs Productivity:** examines the relationship between performance ratings and productivity scores. Average productivity rises from approximately **50 to 85** across the rating scale, showing an association between the two measures.
- **Performance by Job Level:** L5 - Manager has an average rating of 3.5 and L1 - Entry has 3.2.
- **Training Completion:** Customer Support is highest (84.73%) and Human Resources is lowest (61.29%).
- **Manager Rating:** ranges from 3.1 (Data & Analytics) to 3.6 (Finance), a relatively narrow range across departments.

### Key Insights

- The average rating of **3.3** and average goal achievement of **87.40%** indicate a mid-range overall performance profile.
- **Rating 3** is the most common rating (339 employees).
- Productivity varies only slightly between departments (about 69.2–72.6).
- Higher performance ratings are **associated with** higher productivity scores.
- Average performance rating increases from L1 - Entry (3.2) to L5 - Manager (3.5).
- Training completion differs by about 23 percentage points between Customer Support and Human Resources, which helps identify where training follow-up may be needed.

---

## 📈 Dashboard 3: Attendance, Overtime & Compensation Analytics

**Business question:** *"How do attendance, workload and compensation patterns relate to workforce productivity and costs?"*

### KPI Cards

| KPI | Value |
|---|---:|
| Attendance % | 94.68% |
| Average Absence Days | 3.4 |
| Average Overtime Hours | 15.9 |
| Average Annual Salary | 788,845 |
| Average Salary Hike % | 8.42% |

### Main Visuals

1. Attendance % by Department
2. Absence Days by Department
3. Overtime Hours by Department
4. Overtime vs Productivity
5. Average Salary by Job Level
6. Average Salary by Department
7. Salary Distribution by Job Level

### Highlights

- **Attendance:** overall attendance is **94.68%**. Finance (97.14%) and Human Resources (96.93%) are among the highest, and Operations (93.64%) is among the lowest.
- **Absence:** Engineering records **891** total absence days and Customer Support **510**.
- **Overtime:** Engineering records **4,107** overtime hours and Customer Support **3,428.5**.
- **Overtime vs Productivity:** examines the relationship between overtime hours and productivity scores. It is an exploratory view and does not establish causation.
- **Salary Distribution by Job Level:** a pie chart showing how total salary is shared across job levels. It reflects the salary mix by level rather than a statistical spread.

| Job Level | Average Salary |
|---|---:|
| L1 - Entry | 31,232.9 |
| L2 - Associate | 47,647.4 |
| L3 - Senior | 71,589.2 |
| L4 - Lead | 98,247.5 |
| L5 - Manager | 178,607.1 |

| Department | Average Salary |
|---|---:|
| Data & Analytics | 90,533.3 |
| Engineering | 81,566.4 |
| Customer Support | 39,087.8 |

### Key Insights

- Overall attendance is **94.68%**, with an average of **3.4** absence days per employee.
- **Engineering** has the highest absence and overtime totals among the departments shown, in line with it being the largest department. Totals should be read alongside headcount.
- Average salary **increases at every step from L1 to L5**, with L5 - Manager the highest at 178,607.1.
- **Data & Analytics** has the highest department-level average salary, and Customer Support has the lowest of the three departments listed.

---

## 💡 Key Project Insights

### Workforce Capacity
- 736 current employees versus 815 required, a staffing gap of **79** and utilization of **90.31%**.

### Workforce Distribution
- Engineering is the largest department, L2 - Associate is the largest job-level group, and Bengaluru has the highest location headcount.

### Performance
- Average performance rating is **3.3**, average productivity score is **70.8**, and average goal achievement is **87.40%**.
- **74** employees are counted as high performers (see the calculation note above).
- Training completion is **77.58%**.

### Attendance & Workload
- Overall attendance is **94.68%**, average absence is **3.4 days**, and average overtime is **15.9 hours**.
- Engineering has the highest supplied absence and overtime totals.

### Compensation
- Average annual salary is **788,845**, and average salary hike is **8.42%**.
- Average salary rises from L1 to L5, and Data & Analytics has the highest supplied department-level average salary.

---

## 💼 Business Value

This project illustrates how HR teams could use this type of analysis to support:

- **Workforce capacity planning** and **staffing gap identification**
- **Resource allocation** and **department-level workforce planning**
- **Performance monitoring** and **productivity analysis**
- **Employee development** and **training planning**
- **Attendance monitoring**, **workload management**, and **overtime monitoring**
- **Compensation analysis** and **workforce cost planning**
- **HR decision support** more broadly

> These are potential business uses of the analysis. The project is a simulated portfolio exercise and does not claim to have changed real company decisions.

---

## 🙋 Business Questions Answered

1. How many employees are currently in the workforce?
2. How many employees are required?
3. What is the current staffing gap?
4. Where are open positions concentrated?
5. How is the workforce distributed across departments?
6. How is the workforce distributed across locations and job levels?
7. Which departments have higher workload levels?
8. What is the average employee performance rating?
9. What is the average productivity score?
10. How well are employees achieving their goals?
11. How many employees are high performers?
12. How does performance vary by job level?
13. How does productivity vary by department?
14. How complete is employee training?
15. What is the overall attendance rate?
16. Which departments have higher absence levels?
17. Which departments record higher overtime?
18. How does overtime relate to productivity?
19. How does salary vary by job level?
20. How does average salary vary across departments?

---

## 🛠️ Tech Stack

| Tool | Purpose |
|---|---|
| **Microsoft Excel 2021** | Data analysis, calculations, PivotTables, charts, dashboard development |
| **Excel Formulas** | KPI and metric calculations |
| **PivotTables / PivotCharts** | Workforce and employee analysis `[Confirm if used]` |
| **Excel Slicers / Filters** | Interactive dashboard filtering (Department, Job_Role, Location, Job_Level) |

---

## 📁 Project Structure

> The filenames below are **examples/placeholders**. Update them to match your actual repository.

```text
workforce-planning-performance-analytics/
│
├── Data/
│   └── HR_Workforce_Data.xlsx                              # Example filename
│
├── Excel/
│   └── Workforce_Planning_Performance_Analytics.xlsx       # Example filename
│
├── Screenshots/
│   ├── workforce-overview-planning.png                     # Example filename
│   ├── employee-performance-productivity.png               # Example filename
│   └── attendance-overtime-compensation.png                # Example filename
│
└── README.md
```

---

## 🖼️ Dashboard Preview

### Workforce Overview & Planning

![Workforce Overview & Planning](https://github.com/Ankar-G/Workforce-Planning-Employee-Performance-Analytics/blob/main/Dashboards%20Screenshot/Screenshot%202026-10-03%20124756.png)

### Employees Performance & Productivity

![Employees Performance & Productivity](https://github.com/Ankar-G/Workforce-Planning-Employee-Performance-Analytics/blob/main/Dashboards%20Screenshot/Screenshot%202026-10-03%20124813.png)

### Attendance, Overtime & Compensation Analytics

![Attendance, Overtime & Compensation Analytics](https://github.com/Ankar-G/Workforce-Planning-Employee-Performance-Analytics/blob/main/Dashboards%20Screenshot/Screenshot%202026-10-03%20124826.png)


---

## 🤝 Connect With Me

- 💼 **LinkedIn:** `https://www.linkedin.com/in/ankar-goswami-23a196245/`
- 📧 **Email:** `goswamijit99@gmail.com`

---

⭐ If you found this project useful, consider giving the repository a star.
