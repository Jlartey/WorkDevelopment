SELECT Sponsor.SponsorName, CONVERT(VARCHAR(20), MAX(InsuredPatient.EntryDate), 103) AS EntryDate
FROM Sponsor
JOIN InsuredPatient ON Sponsor.SponsorID = InsuredPatient.SponsorID
JOIN SponsorType ON SponsorType.SponsorTypeID = InsuredPatient.SponsorTypeID
--WHERE InsuredPatient.SponsorTypeID = 'S004' AND InsuredPatient.EntryDate BETWEEN '" & periodStart &' AND '" & midYearDate &'
GROUP BY Sponsor.SponsorName
ORDER BY MAX(InsuredPatient.EntryDate) DESC