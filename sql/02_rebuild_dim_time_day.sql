/*
Purpose
- Rebuild dbo.Dim_Time_day to include a complete calendar for 2025 and 2026.
- Populate common calendar attributes used by the report: weekday, month start, ISO week_id, LY/LW dates.

Assumptions
- Weekend definition matches the provided sample (Fri/Sat treated as weekend).
- week_id is stored as ISO week-year + ISO week number in the format YYYYWW.

SQL Server version
- Uses DATETRUNC(iso_week, ...) which is available in SQL Server 2022.
*/

BEGIN TRAN;

-- Build into a staging table first (safer swap)
IF OBJECT_ID('dbo.Dim_Time_day_new', 'U') IS NOT NULL
  DROP TABLE dbo.Dim_Time_day_new;

CREATE TABLE dbo.Dim_Time_day_new (
  day_date            date NOT NULL PRIMARY KEY,
  ly_day_date         date NOT NULL,
  lw_day_day_date     date NOT NULL,
  weekday_id          int  NOT NULL,
  weekday_name        varchar(20) NOT NULL,
  weekday_shortname   varchar(3)  NOT NULL,
  week_id             char(6) NOT NULL,
  weekday_weekend     varchar(10) NOT NULL,
  last_week_same_day  date NOT NULL,
  month_start_dt      date NOT NULL,
  day_of_month        int  NOT NULL,
  day_desc            varchar(11) NOT NULL
);

SET DATEFIRST 7; -- Sunday = 1 (aligns with typical SQL Server defaults and the sample output)

DECLARE @start date = '2025-01-01';
DECLARE @end   date = '2026-12-31';

;WITH D AS (
    SELECT @start AS d
    UNION ALL
    SELECT DATEADD(day, 1, d)
    FROM D
    WHERE d < @end
)
INSERT dbo.Dim_Time_day_new
SELECT
  d AS day_date,
  DATEADD(year, -1, d) AS ly_day_date,
  DATEADD(day, -7, d)  AS lw_day_day_date,
  DATEPART(weekday, d) AS weekday_id,
  DATENAME(weekday, d) AS weekday_name,
  LEFT(DATENAME(weekday, d), 3) AS weekday_shortname,
  CONCAT(
    YEAR(DATEADD(day, 3, DATETRUNC(iso_week, d))),
    RIGHT('00' + CAST(DATEPART(iso_week, d) AS varchar(2)), 2)
  ) AS week_id,
  CASE WHEN DATENAME(weekday, d) IN ('Friday','Saturday') THEN 'Weekend' ELSE 'Weekday' END AS weekday_weekend,
  DATEADD(day, -7, d) AS last_week_same_day,
  DATEFROMPARTS(YEAR(d), MONTH(d), 1) AS month_start_dt,
  DAY(d) AS day_of_month,
  REPLACE(CONVERT(varchar(11), d, 106), ' ', '-') AS day_desc
FROM D
OPTION (MAXRECURSION 0);

-- Swap
IF OBJECT_ID('dbo.Dim_Time_day', 'U') IS NOT NULL
BEGIN
  DROP TABLE dbo.Dim_Time_day;
END

EXEC sp_rename 'dbo.Dim_Time_day_new', 'Dim_Time_day';

COMMIT TRAN;

-- Validate
SELECT MIN(day_date) AS min_dt, MAX(day_date) AS max_dt FROM dbo.Dim_Time_day;
