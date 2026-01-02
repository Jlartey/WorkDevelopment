--DECLARE @StartDate DATE = '2022-01-01';
--DECLARE @EndDate DATE = '2022-01-31';
--DECLARE @branchID VARCHAR(20) = 'b001';

WITH EMRResultsCTE AS (
    SELECT
       CONVERT(VARCHAR(20), er.Column2) AS EMRVar3BID, 
        req.EMRDate,
        req.VisitationID
		
    FROM dbo.EMRResults er
    JOIN dbo.EMRRequest req
        ON req.EMRRequestID = er.EMRRequestID
    WHERE er.EMRDataID = 'MED001'
      AND er.EMRComponentID = 'MED001.5'
      AND req.EMRDate BETWEEN '2025-01-01' AND DATEADD(DAY, 1, '2025-12-31')
       AND req.BranchID = 'B001' 
),
EMRVar2BCTE AS (
    SELECT *
    FROM dbo.EMRVar3B
    WHERE CONVERT(VARCHAR(20), EMRVar3AID) = 'E060'
    )
SELECT 'Therapeutic Duplication' EMRVar3BID, EMRDate, VisitationID, EMRVar3BName,EMRVar3AID 
FROM EMRResultsCTE 
JOIN EMRVar2BCTE on EMRResultsCTE.EMRVar3BID = EMRVar2BCTE.EMRVar3BID   
    
    
    
    
    
    
    
    
    
    
    