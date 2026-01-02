SELECT DISTINCT STRING_AGG(ResidencePhone, ', ') PhoneNumbers
FROM Patient
JOIN Visitation ON Patient.PatientID = Visitation.PatientID
WHERE SpecialistGroupID = 'CD011'

SELECT DISTINCT STRING_AGG(ResidencePhone, ', ') PhoneNumbers
FROM Patient
JOIN Visitation ON Patient.PatientID = Visitation.PatientID
WHERE SpecialistGroupID = 'CD011'

SELECT DISTINCT STRING_AGG(ResidencePhone, ', ') PhoneNumbers
FROM Patient
JOIN Visitation ON Patient.PatientID = Visitation.PatientID
WHERE SpecialistGroupID = 'CD011'

string_agg aggreagation result exceeded the limit of 8000 bytes. use lob types to avoid result truncation

