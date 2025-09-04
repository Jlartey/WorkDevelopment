SELECT DISTINCT sp.SpecialistTypeName AS [Consultation Type],
    CAST(CONVERT(DATE, vst.VisitDate) AS VARCHAR) AS [Date of Consultation],
    Patient.PatientID AS [Patient ID],
    Patient.PatientName AS [Patient Name],
    vst.PatientAge AS [Age], gen.GenderName AS [Sex],
    VisitType.VisitTypeName AS [First/Follow Up],
    vst.VisitationID AS [Diagnosis],
    dg.DiagnosisType AS [Diagnosis Type],
    temp.temperature AS [Temperature],
    wg.weight12 AS [Weight],
    ht.height AS [Height],
    
    vst.VisitationID AS [Investigations Requested],
    vst.VisitationID  AS [Medications], mo.MedicalOutcomeName AS [Outcome],
    spn.SponsorName AS [Insurance Status],
    s.SpecialistName AS [Attending Doctor],
    Patient.ResidencePhone as [Patient Tel.] 
    FROM Visitation AS vst 
    LEFT JOIN Patient ON Patient.PatientID=vst.PatientID 
    LEFT JOIN Gender AS gen ON gen.GenderID=vst.GenderID 
    LEFT JOIN VisitType ON vst.VisitTypeID = VisitType.VisitTypeID
    LEFT JOIN SpecialistType AS sp ON sp.SpecialistTypeID=vst.SpecialistTypeID 
    LEFT JOIN MedicalOutcome AS mo ON mo.MedicalOutcomeID=vst.MedicalOutcomeID 
    LEFT JOIN SystemUser AS su ON su.SystemUserID=vst.SpecialistID 
    LEFT JOIN Specialist AS s ON s.SpecialistID=vst.SpecialistID 
    LEFT JOIN Sponsor AS spn ON spn.SponsorID=vst.SponsorID 
    LEFT JOIN Diagnosis AS dg ON dg.VisitationID=vst.VisitationID 
    
    
     LEFT JOIN (
     SELECT DISTINCT visitationid, CONVERT(NVARCHAR(MAX), Column3) AS Temperature FROM EMRResults AS emrres
     JOIN EMRRequest emrreq ON emrres.EMRRequestID = emrreq.EMRRequestID
     WHERE emrdataid ='EMR034' AND EMRComponentID='EMR03410') AS temp ON temp.visitationid=vst.VisitationID
     LEFT JOIN (
     SELECT DISTINCT visitationid, CONVERT(NVARCHAR(MAX), Column2) AS Weight12 FROM EMRResults AS emrres
     JOIN EMRRequest emrreq ON emrres.EMRRequestID = emrreq.EMRRequestID
     WHERE emrdataid ='EMR034' AND EMRComponentID='EMR05003') AS wg ON wg.visitationid=vst.VisitationID
     LEFT JOIN (
     SELECT DISTINCT visitationid, CONVERT(NVARCHAR(MAX), Column4) AS height FROM EMRResults AS emrres
     JOIN EMRRequest emrreq ON emrres.EMRRequestID = emrreq.EMRRequestID
     WHERE emrdataid ='EMR034' AND EMRComponentID='EMR05004') AS ht ON ht.visitationid=vst.VisitationID
     WHERE vst.VisitDate BETWEEN '2025-05-01 00:00:00.000' AND '2025-05-31 23:59:59.000'
     AND vst.PatientID <> 'P3'
     AND vst.AgeGroupID = 'A002'
     ORDER BY CAST(CONVERT(DATE, vst.VisitDate) AS VARCHAR) ASC, Patient.PatientName ASC
