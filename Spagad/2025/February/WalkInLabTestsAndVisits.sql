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
