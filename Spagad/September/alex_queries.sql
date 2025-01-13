--Registerred patient details
SELECT PatientID, PatientName, Age, GenderName, CountryName
FROM Patient 
JOIN Country ON Country.CountryID = Patient.CountryID
JOIN Gender ON Gender.GenderID = Patient.genderID
WHERE 1=1
--AND firstvisitdate BETWEEN '' AND ''
--AND CountryID = ''
AND firstYearID = 'YRS2023'
ORDER BY CountryName ASC

--vbs
sql = "SELECT PatientID, PatientName, Age, GenderName, CountryName "
sql = sql & "FROM Patient "
sql = sql & "JOIN Country ON Country.CountryID = Patient.CountryID "
sql = sql & "JOIN Gender ON Gender.GenderID = Patient.genderID "
sql = sql & "WHERE 1=1 "
sql = sql & "AND firstYearID = 'YRS2023' "
sql = sql & "ORDER BY CountryName ASC "


--provide details of patient who died
SELECT visitation.PatientID, Patient.Age, Patient.GenderID, SpecialistTypeID, SpecialistID
FROM Visitation 
JOIN Patient ON Patient.PatientID = Visitation.PatientID
WHERE medicaloutcomeID = 'M002' --AND visitdate BETWEEN '' AND ''

sql = "SELECT visitation.PatientID, Patient.Age, Patient.GenderID, SpecialistTypeID, SpecialistID "
sql = sql & "FROM Visitation "                                          
sql = sql & "JOIN Patient ON Patient.PatientID = Visitation.PatientID "
sql = sql & "WHERE medicaloutcomeID = 'M002' "