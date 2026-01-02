SELECT p.GenderID, CASE WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4 THEN '00-04' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9 THEN '05-09' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14 THEN '10-14' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19 THEN '15-19' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24 THEN '20-24' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29 THEN '25-29' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34 THEN '30-34' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39 THEN '35-39' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44 THEN '40-44' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49 THEN '45-49' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54 THEN '50-54' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59 THEN '55-59' ELSE '60+' END AS AgeRange, COUNT(*) AS TotalRecords, SUM(CASE WHEN Combined.Column1 = '2' THEN 1 ELSE 0 END) AS Positive, SUM(CASE WHEN Combined.Column1 = '1' THEN 1 ELSE 0 END) AS Negative
FROM (
SELECT DISTINCT i.LabRequestID, i.patientID, CAST(lr.Column1 AS NVARCHAR(MAX)) AS Column1
FROM Investigation i
JOIN LabResults lr ON i.LabRequestID = lr.LabRequestID
WHERE i.LabTestID = '86750' AND i.RequestStatusID = 'RRD002' AND lr.LabTestID = '86750' AND lr.testcomponentid = 'L0698' AND i.requestdate BETWEEN '2025-10-01' AND '2025-10-31' UNION
SELECT DISTINCT i.LabRequestID, i.patientID, CAST(lr.Column1 AS NVARCHAR(MAX)) AS Column1
FROM Investigation2 i
JOIN LabResults lr ON i.LabRequestID = lr.LabRequestID
WHERE i.LabTestID = '86750' AND i.RequestStatusID = 'RRD002' AND lr.LabTestID = '86750' AND lr.testcomponentid = 'L0698' AND i.requestdate BETWEEN '2025-10-01' AND '2025-10-31') AS Combined
JOIN Patient p ON Combined.patientID = p.patientID
GROUP BY p.GenderID, CASE WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4 THEN '00-04' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9 THEN '05-09' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14 THEN '10-14' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19 THEN '15-19' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24 THEN '20-24' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29 THEN '25-29' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34 THEN '30-34' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39 THEN '35-39' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44 THEN '40-44' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49 THEN '45-49' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54 THEN '50-54' WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59 THEN '55-59' ELSE '60+' END
ORDER BY AgeRange ASC

-- Formatted

SELECT 
    p.GenderID,
    CASE 
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4 THEN '00-04'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9 THEN '05-09'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14 THEN '10-14'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19 THEN '15-19'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24 THEN '20-24'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29 THEN '25-29'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34 THEN '30-34'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39 THEN '35-39'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44 THEN '40-44'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49 THEN '45-49'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54 THEN '50-54'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59 THEN '55-59'
        ELSE '60+'
    END AS AgeRange,
    COUNT(*) AS TotalRecords,
    SUM(CASE WHEN Combined.Column1 = '2' THEN 1 ELSE 0 END) AS Positive,
    SUM(CASE WHEN Combined.Column1 = '1' THEN 1 ELSE 0 END) AS Negative
FROM (
    -- First part: From Investigation table
    SELECT DISTINCT 
        i.LabRequestID, 
        i.patientID, 
        CAST(lr.Column1 AS NVARCHAR(MAX)) AS Column1
    FROM Investigation i
    JOIN LabResults lr ON i.LabRequestID = lr.LabRequestID
    WHERE i.LabTestID = '86750' 
        AND i.RequestStatusID = 'RRD002' 
        AND lr.LabTestID = '86750' 
        AND lr.testcomponentid = 'L0698' 
        AND i.requestdate BETWEEN '2025-10-01' AND '2025-10-31'

    UNION

    -- Second part: From Investigation2 table
    SELECT DISTINCT 
        i.LabRequestID, 
        i.patientID, 
        CAST(lr.Column1 AS NVARCHAR(MAX)) AS Column1
    FROM Investigation2 i
    JOIN LabResults lr ON i.LabRequestID = lr.LabRequestID
    WHERE i.LabTestID = '86750' 
        AND i.RequestStatusID = 'RRD002' 
        AND lr.LabTestID = '86750' 
        AND lr.testcomponentid = 'L0698' 
        AND i.requestdate BETWEEN '2025-10-01' AND '2025-10-31'
) AS Combined
JOIN Patient p ON Combined.patientID = p.patientID
GROUP BY 
    p.GenderID, 
    CASE 
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4 THEN '00-04'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9 THEN '05-09'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14 THEN '10-14'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19 THEN '15-19'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24 THEN '20-24'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29 THEN '25-29'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34 THEN '30-34'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39 THEN '35-39'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44 THEN '40-44'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49 THEN '45-49'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54 THEN '50-54'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59 THEN '55-59'
        ELSE '60+'
    END
ORDER BY AgeRange ASC;







SELECT 
    p.PatientName,
    p.GenderID,
    CASE 
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4 THEN '00-04'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9 THEN '05-09'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14 THEN '10-14'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19 THEN '15-19'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24 THEN '20-24'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29 THEN '25-29'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34 THEN '30-34'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39 THEN '35-39'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44 THEN '40-44'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49 THEN '45-49'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54 THEN '50-54'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59 THEN '55-59'
        ELSE '60+'
    END AS AgeRange,
    COUNT(*) AS TotalRecords,
    SUM(CASE WHEN Combined.Column1 = '2' THEN 1 ELSE 0 END) AS Positive,
    SUM(CASE WHEN Combined.Column1 = '1' THEN 1 ELSE 0 END) AS Negative
FROM (
    -- First part: From Investigation table
    SELECT DISTINCT 
        i.LabRequestID, 
        i.patientID, 
        CAST(lr.Column1 AS NVARCHAR(MAX)) AS Column1
    FROM Investigation i
    JOIN LabResults lr ON i.LabRequestID = lr.LabRequestID
    WHERE i.LabTestID = '86750' 
        AND i.RequestStatusID = 'RRD002' 
        AND lr.LabTestID = '86750' 
        AND lr.testcomponentid = 'L0698' 
        AND i.requestdate BETWEEN '2025-10-01' AND '2025-10-31'

    UNION

    -- Second part: From Investigation2 table
    SELECT DISTINCT 
        i.LabRequestID, 
        i.patientID, 
        CAST(lr.Column1 AS NVARCHAR(MAX)) AS Column1
    FROM Investigation2 i
    JOIN LabResults lr ON i.LabRequestID = lr.LabRequestID
    WHERE i.LabTestID = '86750' 
        AND i.RequestStatusID = 'RRD002' 
        AND lr.LabTestID = '86750' 
        AND lr.testcomponentid = 'L0698' 
        AND i.requestdate BETWEEN '2025-10-01' AND '2025-10-31'
) AS Combined
JOIN Patient p ON Combined.patientID = p.patientID
GROUP BY 
    p.PatientName,
    p.GenderID, 
    CASE 
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4 THEN '00-04'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9 THEN '05-09'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14 THEN '10-14'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19 THEN '15-19'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24 THEN '20-24'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29 THEN '25-29'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34 THEN '30-34'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39 THEN '35-39'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44 THEN '40-44'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49 THEN '45-49'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54 THEN '50-54'
        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59 THEN '55-59'
        ELSE '60+'
    END
ORDER BY p.PatientName ASC;