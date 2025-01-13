WITH EMRResultsCTE AS (
    SELECT TOP 100 
        LEFT(CONVERT(VARCHAR(20), column2), 8) AS EMRVar2BID, 
        emrdate, 
        VisitationID 
    FROM EmrResults 
    JOIN emrrequest ON emrrequest.emrrequestid = emrresults.emrrequestid 
    WHERE EMRDataID = 'TH060' 
      AND EMRComponentID = 'TH06008' 
      AND EMRdate BETWEEN '2018-01-01' AND '2018-02-28'
),

EMRVar2BCTE AS (
    SELECT * 
    FROM emrvar2b 
    WHERE emrvar2AID = 'TH065'
),

DiagnosisCTE AS (
    SELECT 
        EMRResultsCTE.EMRVar2BID, 
        EMRDate, 
        VisitationID, 
        EMRVar2BName, 
        EMRVar2AID 
    FROM EMRResultsCTE 
    JOIN EMRVar2BCTE ON EMRResultsCTE.EMRVar2BID = EMRVar2BCTE.EMRVar2BID
),

DiagnosisDiseaseCTE AS (
    SELECT 
        DiagnosisCTE.VisitationID, 
        DiseaseName, 
        EMRVar2BName AS [Diagnosis_Status], 
        EMRDate 
    FROM DiagnosisCTE 
    JOIN Diagnosis ON DiagnosisCTE.VisitationID = Diagnosis.VisitationID 
    JOIN Disease ON Disease.DiseaseID = Diagnosis.DiseaseID
)

SELECT 
    [Diagnosis_Status], 
    COUNT(*) AS [Total_Diagnosis] 
FROM DiagnosisDiseaseCTE 
GROUP BY [Diagnosis_Status]
