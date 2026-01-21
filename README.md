# Activity KPI Dashboard (SQL Server + Power BI)

This repository contains the SQL scripts, Power BI assets, and documentation for an **Activity KPI Dashboard** that supports:

- **A) YTD vs YTD LY** (+ % variance, up/down arrow, and trend)
- **B) Last 365 / Last 180 days** (dropdown) vs LY same period (+ % variance, arrow, and trend)
- **C) Custom Period 1 vs Period 2** (two date slicers) (+ % variance, arrow, and trend)

<img width="1137" height="737" alt="YTD" src="https://github.com/user-attachments/assets/7c66b306-fc4d-48dd-82f0-2e688ab9d912" />
<img width="1151" height="735" alt="defined" src="https://github.com/user-attachments/assets/a67cab3c-1e7d-49fb-9581-c8218ea4769b" />
<img width="1152" height="737" alt="custome" src="https://github.com/user-attachments/assets/701e6920-8a36-4476-a98f-293ec56ac04b" />

The model is indicator-driven: totals come from `Fact_Activities` (by `ind_id`) and use `Dim_Indicator` to slice by indicator.

---

## Repo structure

```
activity-kpi-dashboard/
├─ sql/
│  ├─ 01_create_shifted_view.sql
│  └─ 02_rebuild_dim_time_day.sql
├─ powerbi/
│  └─ PLACEHOLDER.md
├─ docs/
│  ├─ Technical_Implementation.pptx
│  └─ Business_KPIs.pptx
├─ images/
│  ├─ YTD.png
│  ├─ defined.png
│  └─ custome.png
├─ .gitignore
└─ README.md
```

---

## 1) Prerequisites

- SQL Server database restored from the provided `.bak` (Cloud SQL for SQL Server or any SQL Server instance)
- Power BI Desktop

If you are using **Cloud SQL for SQL Server**:
- Upload the `.bak` to Cloud Storage, then import into Cloud SQL.

---

## 2) Database setup

### 2.1 Restore the database
Restore the provided backup into a database (example name: `PBI_POC`).

### 2.2 Create the shifted fact view (date update to 2025/2026)
The assignment requires “current/previous year” data. The sample data is historical, so we shift activity dates forward by **+4 years**:
- 2021 → 2025
- 2022 → 2026

Run:

- `sql/01_create_shifted_view.sql`

This creates `dbo.v_Fact_Activities_Shifted` and casts `act_value` safely to INT.

### 2.3 Rebuild the date dimension (2025–2026)
To simplify time intelligence in Power BI, rebuild `dbo.Dim_Time_day` to cover a full daily calendar for **2025-01-01 → 2026-12-31**.

Run:

- `sql/02_rebuild_dim_time_day.sql`

---

## 3) Power BI setup

### 3.1 Connect to SQL Server
Power BI Desktop → **Get Data** → **SQL Server database**.

Load these objects:
- `dbo.v_Fact_Activities_Shifted`
- `dbo.Dim_Indicator`
- `dbo.Dim_Time_day`

### 3.2 Transform (Power Query) — filter to today (optional)
If you want the report to show data only up to “today” at refresh time:

```powerquery
TodayUTC = Date.From(DateTimeZone.UtcNow()),
FilteredToToday = Table.SelectRows(#"Changed Type", each [activity_date] <= TodayUTC)
```

### 3.3 Model relationships
Create a star schema:

- Fact → Indicator: `v_Fact_Activities_Shifted[ind_id]` → `Dim_Indicator[ind_id]`
- Fact → Date: `v_Fact_Activities_Shifted[activity_date]` → `Dim_Time_day[day_date]`

### 3.4 Mark the Date table
Right-click `Dim_Time_day` → **Mark as date table** → select `day_date`.


---

## 4) Measures (high level)

Key patterns used:
- `DATESBETWEEN()` to define dynamic date windows
- `AsOfDate` anchored to the **latest fact date** (max activity date in the fact)

---

## 5) Report pages

Implementation is:
- One page with **bookmark navigation**.

---

## 6) Deliverables

- **PBIX**: place the final Power BI file under `powerbi/`.
- **PPTX** decks:
  - `docs/Technical_Implementation.pptx`
  - `docs/Business_KPIs.pptx`

---

## Troubleshooting

- KPI measures returning blank: verify relationships and that `Dim_Time_day` is marked as Date table.
- Variance shows no % sign: set measure format to **Percentage** in Measure tools.
- Date logic ends at 2026-12-31: use `AsOfDate` based on fact max date so KPIs reflect “latest loaded data”, not the end of the calendar.
