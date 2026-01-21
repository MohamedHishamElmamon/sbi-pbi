/*
Purpose
- Create a shifted Fact view that moves historical activity_date values forward by +4 years
  (e.g., 2021->2025, 2022->2026) while keeping month/day intact.
- Cast act_value to INT and default invalid values to 0.

Notes
- activity_date in the restored backup is stored as an ISO-8601 string like '2022-02-11T00:00:00Z'.
- We convert via datetimeoffset to safely parse the trailing 'Z'.
*/

EXEC sys.sp_executesql N'
CREATE OR ALTER VIEW dbo.v_Fact_Activities_Shifted
AS
SELECT
    fa.emp_id,
    fa.ind_id,
    DATEADD(
        year, 4,
        CONVERT(date, TRY_CONVERT(datetimeoffset(0), fa.activity_date))
    ) AS activity_date,
    COALESCE(TRY_CONVERT(int, fa.act_value), 0) AS act_value,
    fa.Flex_Att1,
    fa.Flex_Att2
FROM dbo.Fact_Activities fa
WHERE TRY_CONVERT(datetimeoffset(0), fa.activity_date) IS NOT NULL;
';

-- Validate
SELECT MIN(activity_date) AS min_dt, MAX(activity_date) AS max_dt
FROM dbo.v_Fact_Activities_Shifted;
