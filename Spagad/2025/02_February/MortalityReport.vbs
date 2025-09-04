'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
tableStyles
MortalityReport

Sub MortalityReport()
    Dim count, sql, rst, visitationID
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT "
    sql = sql & "DISTINCT EMRRequestID, "
    sql = sql & "EMRRequestItems.VisitationID, "
    sql = sql & "PatientName Patient, "
    sql = sql & "GenderName Gender,"
    sql = sql & "DATEDIFF(YEAR, Patient.Birthdate, GETDATE()) Age, "
    sql = sql & "Patient.Occupation, "
    sql = sql & "MaritalStatus.MaritalStatusName [Marital Status], "
    sql = sql & "CAST(Patient.ResidenceAddress AS NVARCHAR(50)) [Residence Address], "
    sql = sql & "Country.CountryName Country, "
    sql = sql & "Sponsor.SponsorName, "
    sql = sql & "CONVERT(VARCHAR(20), Admission.AdmissionDate, 106) [Admission Date], "
    sql = sql & "Ward.WardName Ward, "
    sql = sql & "CONVERT(VARCHAR(20), Admission.DischargeDate, 106) [Discharge Date], "
    sql = sql & "AdmissionStatus.AdmissionStatusName [Admission Status], "
    sql = sql & "MedicalStaff.MedicalStaffName Doctor "
    sql = sql & "From EmrRequestitems "
    sql = sql & "Join Patient "
    sql = sql & "ON Patient.PatientID = EMRRequestItems.PatientID "
    sql = sql & "Join Gender "
    sql = sql & "ON Gender.GenderID = EMRRequestItems.GenderID "
    sql = sql & "Join MaritalStatus "
    sql = sql & "ON MaritalStatus.MaritalStatusID = Patient.MaritalStatusID "
    sql = sql & "Join Country "
    sql = sql & "ON Country.CountryID = Patient.CountryID "
    sql = sql & "Join InsuredPatient "
    sql = sql & "ON EMRRequestItems.InsuredPatientID = InsuredPatient.InsuredPatientID "
    sql = sql & "Join sponsor "
    sql = sql & "ON InsuredPatient.SponsorID = Sponsor.SponsorID "
    sql = sql & "Join Admission "
    sql = sql & "ON Admission.VisitationID = EmrRequestitems.VisitationID "
    sql = sql & "Join AdmissionStatus "
    sql = sql & "ON Admission.AdmissionStatusID = AdmissionStatus.AdmissionStatusID "
    sql = sql & "Join MedicalStaff "
    sql = sql & "ON MedicalStaff.MedicalStaffID = Admission.MedicalStaffID "
    sql = sql & "Join Ward "
    sql = sql & "ON Ward.WardID = Admission.WardID "
    sql = sql & "WHERE EMRDataID = 'TH080'"

    'response.write sql

    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Visitation ID</th>"
            response.write "<th class='myth'>Patient</th>"
            response.write "<th class='myth'>Gender</th>"
            response.write "<th class='myth'>Age</th>"
            response.write "<th class='myth'>Occupation</th>"
            response.write "<th class='myth'>Marital Status</th>"
            response.write "<th class='myth'>Residencial Address</th>"
            response.write "<th class='myth'>Nationality</th>"
            response.write "<th class='myth'>Type Of Insurance</th>"
            response.write "<th class='myth'>Admission Date</th>"
            response.write "<th class='myth'>Ward Admitted To</th>"
            response.write "<th class='myth'>Date Of Discharge</th>"
            response.write "<th class='myth'>Diagnosis</th>"
            response.write "<th class='myth'>Investigation Names</th>"
            response.write "<th class='myth'>Drugs</th>"
            response.write "<th class='myth'>Outcome</th>"
            response.write "<th class='myth'>Admitting Doctor</th>"
            'response.write "<th class='myth'>Discharging Physician</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                visitationID = .fields("VisitationID")
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & visitationID & "</td>"
                response.write "<td class='mytd'>" & .fields("Patient") & "</td>"
                response.write "<td class='mytd'>" & .fields("Gender") & "</td>"
                response.write "<td class='mytd'>" & .fields("Age") & "</td>"
                response.write "<td class='mytd'>" & .fields("Occupation") & "</td>"
                response.write "<td class='mytd'>" & .fields("Marital Status") & "</td>"
                response.write "<td class='mytd'>" & .fields("Residence Address") & "</td>"
                response.write "<td class='mytd'>" & .fields("Country") & "</td>"
                response.write "<td class='mytd'>" & .fields("SponsorName") & "</td>"
                response.write "<td class='mytd'>" & .fields("Admission Date") & "</td>"
                response.write "<td class='mytd'>" & .fields("Ward") & "</td>"
                response.write "<td class='mytd'>" & .fields("Discharge Date") & "</td>"
                response.write "<td class='mytd' style='width: 600px;'>" & Diagnosis(visitationID) & "</td>"
                response.write "<td class='mytd'>" & LabTests(visitationID) & "</td>"
               'response.write "<td class='mytd'>" & Drugs(visitationID) & "</td>"
                response.write "<td class='mytd' style='width: 200px; height: 75px; overflow-y: auto; display: block;'>" & Server.HTMLEncode(Drugs(visitationID)) & "</td>"
                response.write "<td class='mytd'>" & .fields("Admission Status") & "</td>"
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
        response.write "    margin: 20px 0;"
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

Function Diagnosis(visitationID)
    Dim sql, rst

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT STRING_AGG(Disease.DiseaseName, ', ') Diagnosis "
    sql = sql & "From Disease Join Diagnosis "
    sql = sql & "ON Diagnosis.DiseaseID = Disease.DiseaseID "
    sql = sql & "WHERE Diagnosis.VisitationID = '" & visitationID & "'"
    
    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            Diagnosis = .fields("Diagnosis")
        End If
              
    End With
End Function

Function LabTests(visitationID)
    Dim sql, rst

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT STRING_AGG(LabTest.LabTestName, ', ') Labtests "
    sql = sql & "From LabTest Join Investigation2 "
    sql = sql & "ON LabTest.LabTestID = Investigation2.LabTestID "
    sql = sql & "WHERE Investigation2.VisitationID = '" & visitationID & "'"
    
    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            LabTests = .fields("Labtests")
        End If
              
    End With
End Function

Function Drugs(visitationID)
    Dim sql, rst

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT STRING_AGG(Drug.DrugName, ', ') drugs "
    sql = sql & "From drug Join Prescription "
    sql = sql & "ON Drug.DrugID = Prescription.DrugID "
    sql = sql & "WHERE Prescription.VisitationID = '" & visitationID & "'"
    
    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            Drugs = .fields("drugs")
        End If
              
    End With
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>


