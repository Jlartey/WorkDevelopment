CREATE OR ALTER PROC dbo.sp_quarterly_appropriateness2
'2025-01-01','2025-12-31'
    @from DATE,
    @to   DATE
AS
BEGIN
    SET NOCOUNT ON;

    /* ==========================
       ROUTE
       ========================== */
    WITH RouteCTE AS (
        SELECT
            req.VisitationID,
            req.EMRDate,
            YEAR(req.EMRDate) AS VisitYear,
            DATEPART(QUARTER, req.EMRDate) AS VisitQuarter,
            v.EMRVar3BName AS AppropriatenessOfRoute
        FROM dbo.EMRResults er
        JOIN dbo.EMRRequest req
            ON req.EMRRequestID = er.EMRRequestID
        JOIN dbo.EMRVar3B v
            ON TRIM(CONVERT(VARCHAR(150), er.Column5)) = v.EMRVar3BID
           AND v.EMRVar3AID = 'E059'
        WHERE er.EMRDataID = 'MED001'
          AND er.EMRComponentID = 'MED001.4'
          AND req.EMRDate BETWEEN @from AND @to
    ),

    /* ==========================
       FREQUENCY
       ========================== */
    FrequencyCTE AS (
        SELECT
            req.VisitationID,
            req.EMRDate,
            YEAR(req.EMRDate) AS VisitYear,
            DATEPART(QUARTER, req.EMRDate) AS VisitQuarter,
            v.EMRVar3BName AS AppropriatenessOfRoute
        FROM dbo.EMRResults er
        JOIN dbo.EMRRequest req
            ON req.EMRRequestID = er.EMRRequestID
        JOIN dbo.EMRVar3B v
            ON TRIM(CONVERT(VARCHAR(150), er.Column2)) = v.EMRVar3BID
           AND v.EMRVar3AID = 'E059'
        WHERE er.EMRDataID = 'MED001'
          AND er.EMRComponentID = 'MED001.4'
          AND req.EMRDate BETWEEN @from AND @to
    ),

    /* ==========================
       DOSE
       ========================== */
    DoseCTE AS (
        SELECT
            req.VisitationID,
            req.EMRDate,
            YEAR(req.EMRDate) AS VisitYear,
            DATEPART(QUARTER, req.EMRDate) AS VisitQuarter,
            v.EMRVar3BName AS AppropriatenessOfRoute
        FROM dbo.EMRResults er
        JOIN dbo.EMRRequest req
            ON req.EMRRequestID = er.EMRRequestID
        JOIN dbo.EMRVar3B v
            ON CONVERT(VARCHAR(20), er.Column5) = v.EMRVar3BID
           AND v.EMRVar3AID = 'E059'
        WHERE er.EMRDataID = 'MED001'
          AND er.EMRComponentID = 'MED001.3'
          AND req.EMRDate BETWEEN @from AND @to
    ),

    /* ==========================
       DRUG
       ========================== */
    DrugCTE AS (
        SELECT
            req.VisitationID,
            req.EMRDate,
            YEAR(req.EMRDate) AS VisitYear,
            DATEPART(QUARTER, req.EMRDate) AS VisitQuarter,
            v.EMRVar3BName AS AppropriatenessOfRoute
        FROM dbo.EMRResults er
        JOIN dbo.EMRRequest req
            ON req.EMRRequestID = er.EMRRequestID
        JOIN dbo.EMRVar3B v
            ON TRIM(CONVERT(VARCHAR(150), er.Column2)) = v.EMRVar3BID
           AND v.EMRVar3AID = 'E059'
        WHERE er.EMRDataID = 'MED001'
          AND er.EMRComponentID = 'MED001.3'
          AND req.EMRDate BETWEEN @from AND @to
    ),

    /* ==========================
       UNIONED AUDIT COUNTS
       ========================== */
    AuditCounts AS (

        SELECT 'Route' AS Category,
               AppropriatenessOfRoute,
               VisitYear,
               VisitQuarter,
               COUNT(*) AS Count
        FROM RouteCTE
        GROUP BY AppropriatenessOfRoute, VisitYear, VisitQuarter

        UNION ALL

        SELECT 'Frequency',
               AppropriatenessOfRoute,
               VisitYear,
               VisitQuarter,
               COUNT(*)
        FROM FrequencyCTE
        GROUP BY AppropriatenessOfRoute, VisitYear, VisitQuarter

        UNION ALL

        SELECT 'Dose',
               AppropriatenessOfRoute,
               VisitYear,
               VisitQuarter,
               COUNT(*)
        FROM DoseCTE
        GROUP BY AppropriatenessOfRoute, VisitYear, VisitQuarter

        UNION ALL

        SELECT 'Drug',
               AppropriatenessOfRoute,
               VisitYear,
               VisitQuarter,
               COUNT(*)
        FROM DrugCTE
        GROUP BY AppropriatenessOfRoute, VisitYear, VisitQuarter
    ),

    /* ==========================
       QUARTER LABEL
       ========================== */
    AuditWithQuarter AS (
        SELECT *,
               CAST(VisitYear AS VARCHAR(4)) + ' Q' + CAST(VisitQuarter AS VARCHAR(1)) AS QuarterLabel
        FROM AuditCounts
    ),

    /* ==========================
       QOQ CHANGE
       ========================== */
    QoQChange AS (
        SELECT
            Category,
            AppropriatenessOfRoute,
            VisitYear,
            VisitQuarter,
            QuarterLabel,
            Count AS CurrentQuarterCount,
            LAG(Count) OVER (
                PARTITION BY Category, AppropriatenessOfRoute
                ORDER BY VisitYear, VisitQuarter
            ) AS PreviousQuarterCount
        FROM AuditWithQuarter
    )

    /* ==========================
       FINAL OUTPUT (WITH NARRATIVE)
       ========================== */
    SELECT
        Category,
        AppropriatenessOfRoute,
        QuarterLabel,
        CurrentQuarterCount,
        PreviousQuarterCount,
        CurrentQuarterCount - ISNULL(PreviousQuarterCount, 0) AS AbsoluteChange,

        CASE
            WHEN PreviousQuarterCount IS NULL THEN NULL
            WHEN PreviousQuarterCount = 0 THEN NULL
            ELSE ROUND(
                ((1.0 * CurrentQuarterCount - PreviousQuarterCount)
                 / PreviousQuarterCount) * 100, 2
            )
        END AS QoQ_PercentChange,

        /* ==========================
           QI NARRATIVE
           ========================== */
        CASE
            WHEN AppropriatenessOfRoute IN ('Inappropriate', 'Uncertain')
                 AND CurrentQuarterCount >= 5
            THEN
                CONCAT(
                    QuarterLabel,
                    ': ',
                    CurrentQuarterCount,
                    ' ',
                    Category,
                    ' prescriptions were assessed as ',
                    LOWER(AppropriatenessOfRoute),
                    '. Review prescribing practices.'
                )

            WHEN AppropriatenessOfRoute = 'Appropriate'
                 AND PreviousQuarterCount IS NOT NULL
                 AND ((1.0 * CurrentQuarterCount - PreviousQuarterCount)
                      / PreviousQuarterCount) * 100 < -10
            THEN
                CONCAT(
                    QuarterLabel,
                    ': Appropriate prescribing in ',
                    Category,
                    ' declined by ',
                    FORMAT(
                        ((1.0 * CurrentQuarterCount - PreviousQuarterCount)
                         / PreviousQuarterCount) * 100,
                        '0.0'
                    ),
                    '%.'
                )
            ELSE NULL
        END AS QI_Narrative

    FROM QoQChange
    ORDER BY Category, AppropriatenessOfRoute, VisitYear, VisitQuarter;

END;