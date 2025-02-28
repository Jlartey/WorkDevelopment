WITH CombinedLabVisits AS (
    SELECT PatientID, VisitationID
    FROM Investigation
    WHERE RequestDate BETWEEN 'dateArr(0)' AND 'dateArr(1)'
    UNION
    SELECT PatientID, VisitationID
    FROM Investigation2
    WHERE RequestDate BETWEEN 'dateArr(0)' AND 'dateArr(1)'
),
LabVisitCounts AS (
    SELECT PatientID, COUNT(DISTINCT VisitationID) AS TotalLabVisits
    FROM CombinedLabVisits
    GROUP BY PatientID
)
SELECT c.PatientID,
       p.PatientName,
       c.TotalLabVisits
FROM LabVisitCounts c
JOIN Patient p ON p.PatientID = c.PatientID
ORDER BY c.TotalLabVisits DESC;


