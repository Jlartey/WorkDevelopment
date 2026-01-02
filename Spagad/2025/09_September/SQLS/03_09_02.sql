WITH Measurements AS (
    SELECT 
        visitationid,
        MAX(CASE WHEN EMRComponentID = 'EMR03410' THEN TRY_CONVERT(FLOAT, CAST(Column3 AS NVARCHAR(MAX))) END) AS Temperature,
        MAX(CASE WHEN EMRComponentID = 'EMR05003' THEN TRY_CONVERT(FLOAT, CAST(Column2 AS NVARCHAR(MAX))) END) AS Weight12,
        MAX(CASE WHEN EMRComponentID = 'EMR05004' THEN TRY_CONVERT(FLOAT, CAST(Column4 AS NVARCHAR(MAX))) END) AS Height
    FROM EMRResults AS emrres
    JOIN EMRRequest emrreq ON emrres.EMRRequestID = emrreq.EMRRequestID
    WHERE emrdataid = 'EMR034'
    GROUP BY visitationid
),
NHIS_Visits AS (
    SELECT DISTINCT
        v1.PatientID,
        CAST(v1.VisitDate AS DATE) AS VisitDay,
        v1.VisitationID AS NHIS_VisitationID,
        v2.VisitationID AS Other_VisitationID
    FROM Visitation v1
    INNER JOIN Visitation v2
        ON v1.PatientID = v2.PatientID
        AND CAST(v1.VisitDate AS DATE) = CAST(v2.VisitDate AS DATE)
        AND v1.SponsorID = 'NHIS'
        AND v2.SponsorID != 'NHIS'
        AND v1.VisitationID != v2.VisitationID
    WHERE v1.VisitDate BETWEEN '2025-01-01' AND '2025-03-31'
)
SELECT
    sp.SpecialistTypeName AS [Consultation Type],
    CONVERT(VARCHAR(20), vst.VisitDate, 106) AS [Date of Consultation],
    Patient.PatientID AS [Patient ID],
    Patient.PatientName AS [Patient Name],
    vst.PatientAge AS [Age],
    gen.GenderName AS [Sex],
    VisitType.VisitTypeName AS [Visit Type],
    COALESCE(dg.DiseaseID, vst.VisitationID) AS [Diagnosis],
    dg.DiagnosisType AS [Diagnosis Type],
    COALESCE(CAST(temp.Temperature AS VARCHAR(20)), 'N/A') AS [Temperature],
    COALESCE(CAST(temp.Weight12 AS VARCHAR(20)), 'N/A') AS [Weight],
    COALESCE(CAST(temp.Height AS VARCHAR(20)), 'N/A') AS [Height],
    COALESCE(emrreq.EMRRequestName, vst.VisitationID) AS [Investigations Requested],
    COALESCE(pr.PrescriptionName, vst.VisitationID) AS [Medications],
    mo.MedicalOutcomeName AS [Outcome],
    spn.SponsorName AS [Insurance Status],
    s.SpecialistName AS [Attending Doctor],
    Patient.ResidencePhone AS [Patient Tel.],
    nhis.NHIS_VisitationID AS [NHIS Visitation ID],
    nhis.Other_VisitationID AS [Other Visitation ID]
FROM Visitation AS vst
INNER JOIN NHIS_Visits AS nhis
    ON vst.PatientID = nhis.PatientID
    AND CAST(vst.VisitDate AS DATE) = nhis.VisitDay
    AND vst.VisitationID IN (nhis.NHIS_VisitationID, nhis.Other_VisitationID)
LEFT JOIN Patient ON Patient.PatientID = vst.PatientID
LEFT JOIN Gender AS gen ON gen.GenderID = vst.GenderID
LEFT JOIN VisitType ON vst.VisitTypeID = VisitType.VisitTypeID
LEFT JOIN SpecialistType AS sp ON sp.SpecialistTypeID = vst.SpecialistTypeID
LEFT JOIN MedicalOutcome AS mo ON mo.MedicalOutcomeID = vst.MedicalOutcomeID
LEFT JOIN Specialist AS s ON s.SpecialistID = vst.SpecialistID
LEFT JOIN Sponsor AS spn ON spn.SponsorID = vst.SponsorID
LEFT JOIN Diagnosis AS dg ON dg.VisitationID = vst.VisitationID
LEFT JOIN EMRRequest AS emrreq ON emrreq.VisitationID = vst.VisitationID
LEFT JOIN Prescription AS pr ON pr.VisitationID = vst.VisitationID
LEFT JOIN Measurements AS temp ON temp.visitationid = vst.VisitationID
WHERE vst.VisitDate BETWEEN '2025-01-01' AND '2025-03-31'
    AND vst.PatientID <> 'P3'
ORDER BY vst.VisitDate ASC, Patient.PatientName ASC;

WITH Measurements AS (
    SELECT 
        visitationid,
        MAX(CASE WHEN EMRComponentID = 'EMR03410' THEN TRY_CONVERT(FLOAT, CAST(Column3 AS NVARCHAR(MAX))) END) AS Temperature,
        MAX(CASE WHEN EMRComponentID = 'EMR05003' THEN TRY_CONVERT(FLOAT, CAST(Column2 AS NVARCHAR(MAX))) END) AS Weight12,
        MAX(CASE WHEN EMRComponentID = 'EMR05004' THEN TRY_CONVERT(FLOAT, CAST(Column4 AS NVARCHAR(MAX))) END) AS Height
    FROM EMRResults AS emrres
    JOIN EMRRequest emrreq ON emrres.EMRRequestID = emrreq.EMRRequestID
    WHERE emrdataid = 'EMR034'
    GROUP BY visitationid
),
NHIS_Visits AS (
    SELECT DISTINCT
        v1.PatientID,
        CAST(v1.VisitDate AS DATE) AS VisitDay,
        v1.VisitationID AS NHIS_VisitationID,
        v2.VisitationID AS Other_VisitationID
    FROM Visitation v1
    INNER JOIN Visitation v2
        ON v1.PatientID = v2.PatientID
        AND CAST(v1.VisitDate AS DATE) = CAST(v2.VisitDate AS DATE)
        AND v1.SponsorID = 'NHIS'
        AND v2.SponsorID != 'NHIS'
        AND v1.VisitationID != v2.VisitationID
    WHERE v1.VisitDate BETWEEN '2025-01-01' AND '2025-03-31'
),
AggregatedData AS (
    SELECT 
        vst.VisitationID,
        STRING_AGG(COALESCE(dg.DiagnosisDescription, vst.VisitationID), ', ') AS Diagnosis,
        STRING_AGG(dg.DiagnosisType, ', ') AS DiagnosisType,
        STRING_AGG(COALESCE(emrreq.RequestDetails, vst.VisitationID), ', ') AS InvestigationsRequested,
        STRING_AGG(COALESCE(pr.MedicationName, vst.VisitationID), ', ') AS Medications
    FROM Visitation AS vst
    LEFT JOIN Diagnosis AS dg ON dg.VisitationID = vst.VisitationID
    LEFT JOIN EMRRequest AS emrreq ON emrreq.VisitationID = vst.VisitationID
    LEFT JOIN Prescription AS pr ON pr.VisitationID = vst.VisitationID
    GROUP BY vst.VisitationID
)
SELECT
    sp.SpecialistTypeName AS [Consultation Type],
    CONVERT(VARCHAR(20), vst.VisitDate, 106) AS [Date of Consultation],
    Patient.PatientID AS [Patient ID],
    Patient.PatientName AS [Patient Name],
    vst.PatientAge AS [Age],
    gen.GenderName AS [Sex],
    VisitType.VisitTypeName AS [Visit Type],
    agg.Diagnosis AS [Diagnosis],
    agg.DiagnosisType AS [Diagnosis Type],
    COALESCE(CAST(temp.Temperature AS VARCHAR(20)), 'N/A') AS [Temperature],
    COALESCE(CAST(temp.Weight12 AS VARCHAR(20)), 'N/A') AS [Weight],
    COALESCE(CAST(temp.Height AS VARCHAR(20)), 'N/A') AS [Height],
    agg.InvestigationsRequested AS [Investigations Requested],
    agg.Medications AS [Medications],
    mo.MedicalOutcomeName AS [Outcome],
    spn.SponsorName AS [Insurance Status],
    s.SpecialistName AS [Attending Doctor],
    Patient.ResidencePhone AS [Patient Tel.],
    nhis.NHIS_VisitationID AS [NHIS Visitation ID],
    nhis.Other_VisitationID AS [Other Visitation ID]
FROM Visitation AS vst
INNER JOIN NHIS_Visits AS nhis
    ON vst.PatientID = nhis.PatientID
    AND CAST(vst.VisitDate AS DATE) = nhis.VisitDay
    AND vst.VisitationID IN (nhis.NHIS_VisitationID, nhis.Other_VisitationID)
LEFT JOIN Patient ON Patient.PatientID = vst.PatientID
LEFT JOIN Gender AS gen ON gen.GenderID = vst.GenderID
LEFT JOIN VisitType ON vst.VisitTypeID = VisitType.VisitTypeID
LEFT JOIN SpecialistType AS sp ON sp.SpecialistTypeID = vst.SpecialistTypeID
LEFT JOIN MedicalOutcome AS mo ON mo.MedicalOutcomeID = vst.MedicalOutcomeID
LEFT JOIN Specialist AS s ON s.SpecialistID = vst.SpecialistID
LEFT JOIN Sponsor AS spn ON spn.SponsorID = vst.SponsorID
LEFT JOIN Measurements AS temp ON temp.visitationid = vst.VisitationID
LEFT JOIN AggregatedData AS agg ON agg.VisitationID = vst.VisitationID
WHERE vst.VisitDate BETWEEN '2025-01-01' AND '2025-03-31'
    AND vst.PatientID <> 'P3'
GROUP BY
    sp.SpecialistTypeName,
    CONVERT(VARCHAR(20), vst.VisitDate, 106),
    Patient.PatientID,
    Patient.PatientName,
    vst.PatientAge,
    gen.GenderName,
    VisitType.VisitTypeName,
    agg.Diagnosis,
    agg.DiagnosisType,
    temp.Temperature,
    temp.Weight12,
    temp.Height,
    agg.InvestigationsRequested,
    agg.Medications,
    mo.MedicalOutcomeName,
    spn.SponsorName,
    s.SpecialistName,
    Patient.ResidencePhone,
    nhis.NHIS_VisitationID,
    nhis.Other_VisitationID
ORDER BY vst.VisitDate ASC, Patient.PatientName ASC;