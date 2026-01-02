WITH DoctorConsults AS (
SELECT DISTINCT SystemUserID, VisitationID
FROM EMRRequestItems
WHERE EMRDataID IN ('TH060', 'IM051')),

PatientGroups AS (
  SELECT vst.insuranceNo, 
  vst.patientID, 
  vst.InsSchemeModeID, 
  vst.VisitTypeID, 
  vst.GenderID, 
  vst.SpecialistTypeID, ip.InitialDependantID, 
  pat.ResidencePhone, 
  MAX(vst.VisitDate) AS MaxVisitDate
FROM Visitation vst

INNER JOIN Patient pat ON vst.patientID = pat.PatientID
JOIN InsuredPatient ip ON vst.PatientID = ip.PatientID
WHERE vst.VisitDate BETWEEN '01 Oct 2025 00:00:00' AND '28 Oct 2025 23:59:59' AND vst.sponsorID = '005' AND vst.patientID NOT IN ('p1', 'p2', 'E01') AND vst.AgeGroupID = 'A002'
GROUP BY vst.insuranceNo, vst.patientID, pat.ResidencePhone, vst.InsSchemeModeID, vst.VisitTypeID, vst.GenderID, ip.InitialDependantID, vst.SpecialistTypeID),

LatestVisitation AS (
SELECT pg.*, vst.VisitationID, ROW_NUMBER() OVER (PARTITION BY pg.insuranceNo, pg.patientID, pg.InsSchemeModeID, pg.VisitTypeID, pg.SpecialistTypeID, pg.InitialDependantID
ORDER BY vst.VisitDate DESC) AS rn
FROM PatientGroups pg
INNER JOIN Visitation vst ON vst.patientID = pg.patientID AND vst.VisitDate = pg.MaxVisitDate AND vst.InsSchemeModeID = pg.InsSchemeModeID AND vst.VisitTypeID = pg.VisitTypeID AND vst.SpecialistTypeID = pg.SpecialistTypeID AND vst.sponsorID = '005')
SELECT lv.insuranceNo AS MembershipID, lv.MaxVisitDate AS VisitDate, lv.patientID, lv.InsSchemeModeID, lv.VisitTypeID, lv.GenderID, lv.ResidencePhone, lv.SpecialistTypeID, lv.InitialDependantID, (
SELECT TOP 1 Staff.StaffName
FROM DoctorConsults dc
JOIN SystemUser ON SystemUser.SystemUserID = dc.SystemUserID
JOIN Staff ON Staff.StaffID = SystemUser.StaffID
WHERE dc.VisitationID = lv.VisitationID) AS [Attending Doctor]
FROM LatestVisitation lv
WHERE lv.rn = 1;