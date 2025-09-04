SELECT DISTINCT v1.PatientID, convert(VARCHAR(20), v1.VisitDate, 106) VisitDay, v1.VisitationID, v2.VisitationID
FROM Visitation v1
INNER JOIN Visitation v2 ON v1.PatientID = v2.PatientID 
    AND CAST(v1.VisitDate AS DATE) = CAST(v2.VisitDate AS DATE)
    AND v1.SponsorID = 'NHIS'
    AND v2.SponsorID != 'NHIS'
    AND v1.VisitationID != v2.VisitationID
    AND convert(VARCHAR(20), v1.VisitDate, 106) BETWEEN '01 Jul 2025' AND '31 Jul 2025'
    
--WHERE convert(VARCHAR(20), v1.VisitDate, 106) BETWEEN 01 Jul 2025' AND '31 Jul 2025'
SELECT DISTINCT v1.PatientID, convert(VARCHAR(20), v1.VisitDate, 106)VisitDay, v1.VisitationID, v2.VisitationID
FROM Visitation v1
INNER JOIN Visitation v2 ON v1.PatientID = v2.PatientID 
    AND CAST(v1.VisitDate AS DATE) = CAST(v2.VisitDate AS DATE)
    AND v1.SponsorID = 'NHIS'
    AND v2.SponsorID != 'NHIS'
    AND v1.VisitationID != v2.VisitationID
WHERE v1.VisitDate BETWEEN '2025-01-01' AND '2025-01-31'