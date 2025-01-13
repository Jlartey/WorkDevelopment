WITH Prescription1_CTE AS 
(
    SELECT 
        SUM(FinalAmt) AS FinalAmount, 
        DATEPART(YEAR, PrescriptionDate) AS YEAR,
        DATENAME(QUARTER, PrescriptionDate) AS QUARTER,
        DATENAME(MONTH, PrescriptionDate) AS MONTH,
        DATEPART(MONTH, PrescriptionDate) AS MONTH_INDEX
    FROM Prescription
    WHERE PrescriptionDate BETWEEN '2018-01-01' AND '2022-12-31'
    GROUP BY 
        DATEPART(YEAR, PrescriptionDate),
        DATENAME(QUARTER, PrescriptionDate),
        DATENAME(MONTH, PrescriptionDate),
        DATEPART(MONTH, PrescriptionDate)
),

Prescription2_CTE AS 
(
    SELECT 
        SUM(FinalAmt) AS FinalAmount, 
        DATEPART(YEAR, ConsultReviewDate) AS YEAR,
        DATEPART(QUARTER, ConsultReviewDate) AS QUARTER,
        DATENAME(MONTH, ConsultReviewDate) AS MONTH,
        DATEPART(MONTH, ConsultReviewDate) AS MONTH_INDEX
    FROM Prescription2
    WHERE ConsultReviewDate BETWEEN '2018-01-01' AND '2022-12-31'
    GROUP BY 
        DATEPART(YEAR, ConsultReviewDate),
        DATEPART(QUARTER, ConsultReviewDate),
        DATENAME(MONTH, ConsultReviewDate),
        DATEPART(MONTH, ConsultReviewDate)
),

SUM_FINAL_AMOUNT_CTE AS 
(
    SELECT * FROM Prescription1_CTE
    UNION ALL 
    SELECT * FROM Prescription2_CTE
)

SELECT 
    CONVERT(DECIMAL(18, 2), SUM(FinalAmount)) AS SUM_OF_FINAL_AMOUNT, 
    [YEAR], 
    [QUARTER],
    [MONTH],
    [MONTH_INDEX],
    CONVERT(DECIMAL(18, 2), LAG(SUM(FinalAmount), 1, 0) OVER (PARTITION BY YEAR ORDER BY MONTH_INDEX)) AS PREVIOUS_AMOUNT,
    CONVERT(DECIMAL(18, 2), SUM(FinalAmount) - LAG(SUM(FinalAmount), 1, 0) OVER (PARTITION BY YEAR ORDER BY MONTH_INDEX)) AS DIFFERENCE,
    CONVERT(DECIMAL(18, 2), 
            CASE 
                WHEN LAG(SUM(FinalAmount), 1, 0) OVER (PARTITION BY YEAR ORDER BY MONTH_INDEX) = 0 THEN 0
                ELSE ((SUM(FinalAmount) - LAG(SUM(FinalAmount), 1, 0) OVER (PARTITION BY YEAR ORDER BY MONTH_INDEX)) * 100.0 / LAG(SUM(FinalAmount), 1, 0) OVER (PARTITION BY YEAR ORDER BY MONTH_INDEX))
            END) AS PERCENTAGE_CHANGE
FROM 
    SUM_FINAL_AMOUNT_CTE
GROUP BY 
    YEAR, QUARTER, MONTH, MONTH_INDEX
ORDER BY 
    YEAR, QUARTER, MONTH_INDEX;
