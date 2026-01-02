SELECT 
    Staff.StaffName Doctor, 
    COUNT(DISTINCT Diagnosis.VisitationId) Consults, 
    VisitTotals.VisitAmount 
    From diagnosis 
    Join SystemUser 
    ON SystemUser.SystemUserID = Diagnosis.SystemUserID 
    Join Staff 
    ON Staff.StaffID = SystemUser.StaffID 
    JOIN ( 

    SELECT 
    v.VisitationID, 
    MIN(v.visitCost) As visitCost 
    FROM Visitation v 
    WHERE EXISTS ( 
    SELECT 1 
    FROM Diagnosis d 
    Where d.visitationID = v.visitationID 
    ) 
    GROUP BY v.VisitationID 
    ) UniqueVisits 
    ON Diagnosis.VisitationId = UniqueVisits.VisitationID 
    JOIN ( 
    SELECT StaffName, SUM(VisitCost) AS VisitAmount 
    FROM ( 
    SELECT DISTINCT s.StaffName, v.VisitationID, v.VisitCost 
    FROM Diagnosis d 
    JOIN SystemUser su ON su.SystemUserID = d.SystemUserID 
    JOIN Staff s ON s.StaffID = su.StaffID 
    JOIN Visitation v ON d.VisitationId = v.VisitationID 
    WHERE ConsultReviewDate BETWEEN '2025-08-01' AND '2025-08-31' 
    ) DistinctVisits 
    GROUP BY StaffName 
    ) VisitTotals 
    ON Staff.StaffName = VisitTotals.StaffName 
    
        WHERE ConsultReviewDate BETWEEN '2025-08-01' AND '2025-08-31' 
    
    GROUP BY Staff.StaffName, VisitTotals.VisitAmount 
    ORDER BY Consults DESC 


    -- GOOD TO GOO
    SELECT SystemUserID, COUNT(DISTINCT VisitationID)Count FROM
EMRRequestItems 
WHERE EMRDataID IN ('TH060', 'IM051')
AND EMRDate BETWEEN '2025-08-01' AND '2025-08-31'
GROUP BY SystemUserID 
ORDER BY COUNT desc

--2 GOOD AS WELL
WITH DoctorConsults AS (SELECT SystemUserID, COUNT(DISTINCT VisitationID)Consults FROM
EMRRequestItems 
WHERE EMRDataID IN ('TH060', 'IM051')
AND EMRDate BETWEEN '2025-09-16' AND '2025-09-17'
GROUP BY SystemUserID 

)
SELECT Staff.StaffName, dc.Consults
FROM DoctorConsults dc
JOIN SystemUser
	ON SystemUser.SystemUserID = dc.SystemUserID
JOIN Staff
	ON Staff.StaffID = SystemUser.StaffID
ORDER BY dc.Consults DESC


--3 GOOD AS WELL
SELECT distinct VisitationID, p.PatientName FROM
EMRRequestItems emr
JOIN 
Patient p
	ON p.PatientID = emr.PatientID
WHERE EMRDataID IN ('TH060', 'IM051')
AND EMRDate BETWEEN '2025-08-01' AND '2025-08-31'
AND emr.SystemUserID = 'mh1810002'

