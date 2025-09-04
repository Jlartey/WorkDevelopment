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
    SELECT 
        v1.PatientID,
        CAST(v1.VisitDate AS DATE) AS VisitDay,
        CAST(STRING_AGG(CAST(v1.VisitationID AS NVARCHAR(MAX)), ', ') AS NVARCHAR(MAX)) AS NHIS_VisitationIDs,
        CAST(STRING_AGG(CAST(v2.VisitationID AS NVARCHAR(MAX)), ', ') AS NVARCHAR(MAX)) AS Other_VisitationIDs
    FROM Visitation v1
    INNER JOIN Visitation v2
        ON v1.PatientID = v2.PatientID
        AND CAST(v1.VisitDate AS DATE) = CAST(v2.VisitDate AS DATE)
        AND v1.SponsorID = 'NHIS'
        AND v2.SponsorID != 'NHIS'
    WHERE v1.VisitDate BETWEEN '2025-01-01' AND '2025-03-31'
    GROUP BY v1.PatientID, CAST(v1.VisitDate AS DATE)
),
AggregatedData AS (
    SELECT 
        vst.VisitationID,
        CAST(STRING_AGG(CAST(dg.DiseaseID AS NVARCHAR(MAX)), ', ') AS NVARCHAR(MAX)) AS Diagnosis,
        CAST(STRING_AGG(CAST(dg.DiagnosisType AS NVARCHAR(MAX)), ', ') AS NVARCHAR(MAX)) AS DiagnosisType,
        CAST(STRING_AGG(CAST(emrreq.EMRRequestName AS NVARCHAR(MAX)), ', ') AS NVARCHAR(MAX)) AS InvestigationsRequested,
        CAST(STRING_AGG(CAST(pr.PrescriptionName AS NVARCHAR(MAX)), ', ') AS NVARCHAR(MAX)) AS Medications
    FROM Visitation AS vst
    LEFT JOIN (
        SELECT DISTINCT VisitationID, DiseaseID, DiagnosisType
        FROM Diagnosis
    ) AS dg ON dg.VisitationID = vst.VisitationID
    LEFT JOIN (
        SELECT DISTINCT VisitationID, EMRRequestName
        FROM EMRRequest
    ) AS emrreq ON emrreq.VisitationID = vst.VisitationID
    LEFT JOIN (
        SELECT DISTINCT VisitationID, PrescriptionName
        FROM Prescription
    ) AS pr ON pr.VisitationID = vst.VisitationID
    GROUP BY vst.VisitationID
)
SELECT
    MAX(sp.SpecialistTypeName) AS [Consultation Type],
    CONVERT(VARCHAR(20), vst.VisitDate, 106) AS [Date of Consultation],
    Patient.PatientID AS [Patient ID],
    MAX(Patient.PatientName) AS [Patient Name],
    MAX(vst.PatientAge) AS [Age],
    MAX(gen.GenderName) AS [Sex],
    MAX(VisitType.VisitTypeName) AS [Visit Type],
    CAST(STRING_AGG(agg.Diagnosis, ', ') AS NVARCHAR(MAX)) AS [Diagnosis],
    CAST(STRING_AGG(agg.DiagnosisType, ', ') AS NVARCHAR(MAX)) AS [Diagnosis Type],
    COALESCE(CAST(MAX(temp.Temperature) AS VARCHAR(20)), 'N/A') AS [Temperature],
    COALESCE(CAST(MAX(temp.Weight12) AS VARCHAR(20)), 'N/A') AS [Weight],
    COALESCE(CAST(MAX(temp.Height) AS VARCHAR(20)), 'N/A') AS [Height],
    CAST(STRING_AGG(agg.InvestigationsRequested, ', ') AS NVARCHAR(MAX)) AS [Investigations Requested],
    CAST(STRING_AGG(agg.Medications, ', ') AS NVARCHAR(MAX)) AS [Medications],
    MAX(mo.MedicalOutcomeName) AS [Outcome],
    MAX(spn.SponsorName) AS [Insurance Status],
    MAX(s.SpecialistName) AS [Attending Doctor],
    MAX(Patient.ResidencePhone) AS [Patient Tel.],
    nhis.NHIS_VisitationIDs AS [NHIS Visitation ID],
    nhis.Other_VisitationIDs AS [Other Visitation ID]
FROM Visitation AS vst
INNER JOIN NHIS_Visits AS nhis
    ON vst.PatientID = nhis.PatientID
    AND CAST(vst.VisitDate AS DATE) = nhis.VisitDay
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
    Patient.PatientID,
    CAST(vst.VisitDate AS DATE),
    CONVERT(VARCHAR(20), vst.VisitDate, 106),
    nhis.NHIS_VisitationIDs,
    nhis.Other_VisitationIDs
ORDER BY CAST(vst.VisitDate AS DATE) ASC, MAX(Patient.PatientName) ASC;