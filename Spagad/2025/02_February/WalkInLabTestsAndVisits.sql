-- Lab Tests
SELECT 
    LabRequestName, 
    COUNT(LabTestID) AS TotalLabTests
FROM LabRequest 
JOIN  Investigation  ON LabRequest.LabRequestID = Investigation.LabRequestID
WHERE LabRequest.VisitationId = 'E01'
GROUP BY LabRequest.LabRequestName
ORDER BY TotalLabTests DESC 

-- Lab Visits
SELECT  
    LabRequestName, 
    CONVERT(VARCHAR(20), CONVERT(DATE, RequestDate), 106) AS LabRequestDate, 
    COUNT(DISTINCT CONVERT(DATE, RequestDate)) AS NumberOfVisits
FROM LabRequest 
WHERE VisitationID = 'E01'
GROUP BY LabRequestName, CONVERT(DATE, RequestDate)

SELECT  
    LabRequestName, 
    CONVERT(VARCHAR(20), CONVERT(DATE, Investigation.RequestDate), 106) AS LabRequestDate, 
    COUNT(DISTINCT CONVERT(DATE, Investigation.RequestDate)) AS NumberOfVisits
FROM LabRequest 
JOIN Investigation 
ON LabRequest.LabRequestID = Investigation.LabRequestID
WHERE Investigation.VisitationID = 'E01'
GROUP BY LabRequestName, CONVERT(DATE, Investigation.RequestDate)

-- Correct Solution - 25/02/2025
SELECT LabRequestName, COUNT(DISTINCT CONVERT(date, RequestDate)) AS LabVisit
FROM LabRequest
WHERE VisitationID = 'E01'
AND RequestDate BETWEEN '2022-02-25 10:32:32.000' AND '2025-02-25 10:32:32.000'
GROUP BY LabRequestName
ORDER BY LabVisit DESC

--In VbScript
sql = "SELECT LabRequestName, COUNT(DISTINCT CONVERT(date, RequestDate)) AS LabVisit "
sql = sql & "FROM LabRequest "
sql = sql & "WHERE VisitationID = 'E01' "
sql = sql & "AND RequestDate BETWEEN '2022-02-25 10:32:32.000' AND '2025-02-25 10:32:32.000' "
sql = sql & "GROUP BY LabRequestName "
sql = sql & "ORDER BY LabVisit DESC "