'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
tableStyles
MortalityReport

Sub MortalityReport()
    Dim count, sql, rst, visitationID, emrRequestID
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT DISTINCT "
    sql = sql & "EMRRequestID, "
    sql = sql & "EMRRequestItems.visitationID, "
    sql = sql & "PatientName AS Patient, "
    sql = sql & "GenderName AS Gender, "
    sql = sql & "DATEDIFF(YEAR, Patient.Birthdate, GETDATE()) AS Age, "
    sql = sql & "Patient.Occupation, "
    sql = sql & "MaritalStatus.MaritalStatusName AS [Marital Status], "
    sql = sql & "CAST(Patient.ResidenceAddress AS NVARCHAR(50)) AS [Residence Address], "
    sql = sql & "Country.CountryName AS Country, "
    sql = sql & "Sponsor.SponsorName AS SponsorName, "
    sql = sql & "CONVERT(VARCHAR(20), Admission.AdmissionDate, 106) AS [Admission Date], "
    sql = sql & "Ward.WardName AS Ward, "
    sql = sql & "CONVERT(VARCHAR(20), Admission.DischargeDate, 106) AS [Discharge Date], "
    sql = sql & "MedicalOutcome.MedicalOutcomeName AS Outcome, "
    sql = sql & "MedicalStaff.MedicalStaffName AS Doctor "
    sql = sql & "FROM EMRRequestItems "
    sql = sql & "LEFT JOIN Patient ON Patient.PatientID = EMRRequestItems.PatientID "
    sql = sql & "LEFT JOIN Gender ON Gender.GenderID = EMRRequestItems.GenderID "
    sql = sql & "LEFT JOIN MaritalStatus ON MaritalStatus.MaritalStatusID = Patient.MaritalStatusID "
    sql = sql & "LEFT JOIN Country ON Country.CountryID = Patient.CountryID "
    sql = sql & "LEFT JOIN InsuredPatient ON EMRRequestItems.InsuredPatientID = InsuredPatient.InsuredPatientID "
    sql = sql & "LEFT JOIN Sponsor ON InsuredPatient.SponsorID = Sponsor.SponsorID "
    sql = sql & "LEFT JOIN Admission ON Admission.VisitationID = EMRRequestItems.VisitationID "
    sql = sql & "LEFT JOIN Visitation ON Visitation.VisitationID = EMRRequestItems.VisitationID "
    sql = sql & "LEFT JOIN MedicalOutcome ON MedicalOutcome.MedicalOutcomeID = Visitation.MedicalOutcomeID "
    sql = sql & "LEFT JOIN MedicalStaff ON MedicalStaff.MedicalStaffID = Admission.MedicalStaffID "
    sql = sql & "LEFT JOIN Ward ON Ward.WardID = Admission.WardID "
    sql = sql & "WHERE EMRDataID = 'TH080';"

    'response.write sql

    With rst
        .Open qryPro.FltQry(sql), conn, 3, 4
        
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Visitation ID</th>"
            response.write "<th class='myth'>Patient</th>"
            response.write "<th class='myth'>Gender</th>"
            response.write "<th class='myth'>Outcome</th>"
            response.write "<th class='myth'>Age</th>"
            response.write "<th class='myth'>Occupation</th>"
            response.write "<th class='myth'>Marital Status</th>"
            response.write "<th class='myth'>Residential Address</th>"
            response.write "<th class='myth'>Nationality</th>"
            response.write "<th class='myth'>Type Of Insurance</th>"
            response.write "<th class='myth'>Admission Date</th>"
            response.write "<th class='myth'>Ward Admitted To</th>"
            response.write "<th class='myth'>Date Of Discharge</th>"
            response.write "<th class='myth'>Diagnosis</th>"
            response.write "<th class='myth'>Investigation Names</th>"
            response.write "<th class='myth'>Drugs</th>"
            
            response.write "<th class='myth'>Doctor</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                visitationID = .fields("VisitationID")
                emrRequestID = .fields("EMRRequestID")
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & visitationID & "</td>"
                response.write "<td class='mytd' style='min-width: 200px;'>" & .fields("Patient") & "</td>"
                response.write "<td class='mytd'>" & .fields("Gender") & "</td>"
                response.write "<td class='mytd'>" & .fields("Outcome") & "</td>"
                response.write "<td class='mytd'>" & .fields("Age") & "</td>"
                response.write "<td class='mytd'>" & .fields("Occupation") & "</td>"
                response.write "<td class='mytd'>" & .fields("Marital Status") & "</td>"
                response.write "<td class='mytd'>" & .fields("Residence Address") & "</td>"
                response.write "<td class='mytd'>" & .fields("Country") & "</td>"
                response.write "<td class='mytd'>" & .fields("SponsorName") & "</td>"
                response.write "<td class='mytd'>" & .fields("Admission Date") & "</td>"
                response.write "<td class='mytd'>" & .fields("Ward") & "</td>"
                response.write "<td class='mytd'>" & .fields("Discharge Date") & "</td>"
                response.write "<td class='mytd' style='width: 600px;'>" & diagnosis(visitationID) & "</td>"
                response.write "<td class='mytd'>" & LabTests(visitationID) & "</td>"
                response.write "<td class='mytd'>" & drugs(visitationID) & "</td>"
                
                response.write "<td class='mytd' style='min-width: 150px;'>" & .fields("Doctor") & "</td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop

            response.write "</table>"
        Else
            response.write "<h1>No records found</h1>"
        End If
        
        .Close
    End With
    
    Set rst = Nothing
End Sub


Sub tableStyles()
    response.write "<style>"
        response.write ".mytable {"
        response.write "    width: fit-content;"
        response.write "    border-collapse: collapse;"
        response.write "    margin: 50px 50px;"
        response.write "    font-size: 16px;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "}"
        response.write ".mytable, .myth, .mytd {"
        response.write "    border: 1px solid #dddddd;"
        response.write "}"
        response.write ".myth, .mytd {"
        response.write "    padding: 12px;"
        response.write "    text-align: left;"
        response.write "}"
        response.write ".myth {"
        response.write "    background-color: #f2f2f2;"
        response.write "    color: #333;"
        response.write "    font-weight: bold;"
        response.write "}"
        response.write ".mytr:nth-child(even) {"
        response.write "    background-color: #f9f9f9;"
        response.write "}"
        response.write ".mytr:hover {"
        response.write "    background-color: #f1f1f1;"
        response.write "}"
        response.write ".myth {"
        response.write "    text-transform: uppercase;"
        response.write "}"
        response.write "h1 {"
        response.write "    font-size: 18px;"
        response.write "    color: #555;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "    margin: 20px 0;"
        response.write "}"
response.write "</style>"

End Sub

Function diagnosis(visitationID)
    Dim sql, rst

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT STRING_AGG(Disease.DiseaseName, ', ') Diagnosis "
    sql = sql & "From Disease Join Diagnosis "
    sql = sql & "ON Diagnosis.DiseaseID = Disease.DiseaseID "
    sql = sql & "WHERE Diagnosis.VisitationID = '" & visitationID & "'"
    
    With rst
        .Open qryPro.FltQry(sql), conn, 3, 4
        
        If .RecordCount > 0 Then
            diagnosis = .fields("Diagnosis")
        End If
              
    End With
End Function

Function LabTests(visitationID)
    Dim sql, rst

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT ISNULL(STRING_AGG(LabTestName, ', '), 'No Investigations Done') AS Labtests "
    sql = sql & "FROM ( "
    sql = sql & "    SELECT LabTest.LabTestName "
    sql = sql & "    FROM LabTest "
    sql = sql & "    JOIN Investigation ON LabTest.LabTestID = Investigation.LabTestID "
    sql = sql & "    WHERE Investigation.VisitationID = '" & visitationID & "'"
    sql = sql & "    UNION ALL "
    sql = sql & "    SELECT LabTest.LabTestName "
    sql = sql & "    FROM LabTest "
    sql = sql & "    JOIN Investigation2 ON LabTest.LabTestID = Investigation2.LabTestID "
    sql = sql & "    WHERE Investigation2.VisitationID = '" & visitationID & "'"
    sql = sql & ") AS CombinedResults;"

    With rst
        .Open qryPro.FltQry(sql), conn, 3, 4
        
        If .RecordCount > 0 Then
            LabTests = .fields("Labtests")
        Else
            LabTests = "No Investigations Done"
        End If
              
    End With
End Function



Function drugs(visitationID)
    Dim sql, rst

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT ISNULL(STRING_AGG(Drug.DrugName, ', '), 'No drugs prescribed') AS drugs "
    sql = sql & "FROM Drug "
    sql = sql & "JOIN Prescription ON Drug.DrugID = Prescription.DrugID "
    sql = sql & "WHERE Prescription.VisitationID = '" & visitationID & "'"

    With rst
        .Open qryPro.FltQry(sql), conn, 3, 4
        
        If Not .EOF Then
            drugs = .fields("drugs")
        Else
            drugs = "No drugs prescribed"
        End If
        .Close
    End With

    Set rst = Nothing
End Function


'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
