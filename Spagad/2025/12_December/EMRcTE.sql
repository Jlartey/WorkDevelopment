--DECLARE @StartDate DATE = '2025-01-01';
--DECLARE @EndDate DATE = '2025-01-31';
--DECLARE @branchID VARCHAR(20) = 'b001';

--WITH EMRResultsCTE AS (
    SELECT
        CONVERT(VARCHAR(20), er.Column2) AS DATE,  -- Improved: trim spaces for reliable matching
        req.EMRDate,
        req.VisitationID
		
    FROM dbo.EMRResults er
    JOIN dbo.EMRRequest req
        ON req.EMRRequestID = er.EMRRequestID
    WHERE er.EMRDataID = 'MED001'
      AND er.EMRComponentID IN ('MED001.1', 'MED001.3')
      AND req.EMRDate BETWEEN '2025-01-01' AND DATEADD(DAY, 1, '2025-12-31')
       AND req.BranchID = 'B001' 
--),
