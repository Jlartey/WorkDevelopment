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
    ) DistinctVisits 
  GROUP BY StaffName 
  ) VisitTotals 
  ON Staff.StaffName = VisitTotals.StaffName 
  WHERE ConsultReviewDate BETWEEN '01 May 2025 00:00:00' AND '31 May 2025 23:59:59' 
  GROUP BY Staff.StaffName, VisitTotals.VisitAmount 
  ORDER BY Consults DESC