'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim medicalOutcome
medicalOutcome = Trim(Request.QueryString("MedicalOutcomeID"))
If IsNull(medicalOutcome) Or medicalOutcome = "" Then
    medicalOutcome = ""
End If

tableStyles
header
MortalityReport

Sub header()
    Dim dropdownOptions

    sql = "SELECT MedicalOutcomeID, MedicalOutcomeName FROM MedicalOutcome"
    
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.open sql, conn, 3, 4

    dropdownOptions = "<option value=''>" & "All" & "</option>"

    With rstDropdown
        If .recordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                optionHTML = "<option value='" & .fields("MedicalOutcomeID") & "'>" & .fields("MedicalOutcomeName") & "</option>"
                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With

    rstDropdown.Close
    Set rstDropdown = Nothing

response.write "<div class='filters'>"
    response.write "        <label for='medicalOutcome' class='font-style'>Select Medical Outcome:</label><br>"
    response.write "        <select id='medicalOutcome' name='medicalOutcome'>"
    response.write dropdownOptions
    response.write "        </select>"
    response.write "        <button type='button' onclick='updateUrl()'>Show Data</button> <br />"
response.write "</div>"

response.write "<script>"
    response.write "    function updateUrl() {"
    
    response.write "        const medicalOutcomes = Array.from(document.getElementById('medicalOutcome').selectedOptions).map(option => option.value).join(',');"
    response.write "        const baseUrl = 'http://172.19.0.36/hms/wpgPrtPrintLayoutAll.asp';"
    response.write "        const params = new URLSearchParams({"
    response.write "            PrintLayoutName: 'MortalityReport',"
    response.write "            PositionForTableName: 'WorkingDay',"
    response.write "            WorkingDayID: '',"
    response.write "            MedicalOutcomeID: medicalOutcomes"
    response.write "        });"
    response.write "        const newUrl = baseUrl + '?' + params.toString();"
    response.write "        window.location.href = newUrl;"
    response.write "        console.log(newUrl);"
    response.write "    }"
    response.write "</script>"

End Sub

Sub MortalityReport()
    Dim count, sql, rst, VisitationID, emrRequestID, recordCount
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
    sql = sql & "WHERE EMRDataID = 'TH080' "
    If medicalOutcome <> "" Then
        sql = sql & "AND MedicalOutcome.MedicalOutcomeID = '" & medicalOutcome & "'"
    End If
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
        If .recordCount > 0 Then
            recordCount = .recordCount
            If medicalOutcome = "" Then
                response.write "<h1>Showing All " & recordCount & " Records</h1>"
            Else
                response.write "<h1>Showing Data Of " & recordCount & " Patients Whose Medical Outcome = " & .fields("Outcome") & "</h1>"
            End If
            
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
            
            response.Flush
            
            Do While Not .EOF
                VisitationID = .fields("VisitationID")
                emrRequestID = .fields("EMRRequestID")
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & VisitationID & "</td>"
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
                response.write "<td class='mytd'>" & diagnosis(VisitationID) & "</td>"
                response.write "<td class='mytd'>" & LabTests(VisitationID) & "</td>"
                response.write "<td class='mytd'>" & drugs(VisitationID) & "</td>"
                
                response.write "<td class='mytd' style='min-width: 150px;'>" & .fields("Doctor") & "</td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
                
                response.Flush
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
        response.write "    font-size: 22px;"
        response.write "    color: #555;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "    margin: 20px 0;"
        response.write "    text-transform: uppercase;"
        response.write "}"
        response.write "    .filters {"
        response.write "        padding: 10px;"
        response.write "        border: 1px solid #ccc;"
        response.write "        background-color: #f9f9f9;"
        response.write "        border-radius: 8px;"
        response.write "        width: 300px;"
        response.write "        margin-bottom: 15px;"
        response.write "        margin-top: 30px; /* Added 30px padding on top */"
        response.write "    }"
        
        response.write "    .font-style {"
        response.write "        font-family: Arial, sans-serif;"
        response.write "        font-size: 14px;"
        response.write "        font-weight: bold;"
        response.write "        color: #333;"
        response.write "    }"
        
        response.write "    select {"
        response.write "        width: 100%;"
        response.write "        padding: 8px;"
        response.write "        border-radius: 4px;"
        response.write "        border: 1px solid #aaa;"
        response.write "        font-size: 14px;"
        response.write "        background-color: #fff;"
        response.write "        cursor: pointer;"
        response.write "    }"
        
        response.write "    button {"
        response.write "        margin-top: 10px;"
        response.write "        padding: 8px 15px;"
        response.write "        border: none;"
        response.write "        background-color: #007BFF;"
        response.write "        color: white;"
        response.write "        font-size: 14px;"
        response.write "        border-radius: 4px;"
        response.write "        cursor: pointer;"
        response.write "        transition: background 0.3s ease-in-out;"
        response.write "    }"
        
        response.write "    button:hover {"
        response.write "        background-color: #0056b3;"
        response.write "    }"

        response.write "</style>"

End Sub

Function diagnosis(VisitationID)
    Dim sql, rst, output, n

    Set rst = CreateObject("ADODB.Recordset")
    sql = "SELECT Disease.DiseaseName " & _
          "FROM Disease " & _
          "JOIN Diagnosis ON Diagnosis.DiseaseID = Disease.DiseaseID " & _
          "WHERE Diagnosis.VisitationID = '" & VisitationID & "'"
    
    output = "<ol>"
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If Not .EOF Then
            Do While Not .EOF
                output = output & "<li>" & .fields("DiseaseName") & "</li>"
                .MoveNext
            Loop
        Else
            output = output & "<li>No Diagnosis</li>"
        End If
        .Close
    End With
    output = output & "</ol>"
    Set rst = Nothing
    
    diagnosis = output
End Function

Function LabTests(VisitationID)
    Dim sql, rst, output

    Set rst = CreateObject("ADODB.Recordset")
    sql = "SELECT LabTest.LabTestName " & _
          "FROM LabTest " & _
          "JOIN Investigation ON LabTest.LabTestID = Investigation.LabTestID " & _
          "WHERE Investigation.VisitationID = '" & VisitationID & "'" & _
          " UNION ALL " & _
          "SELECT LabTest.LabTestName " & _
          "FROM LabTest " & _
          "JOIN Investigation2 ON LabTest.LabTestID = Investigation2.LabTestID " & _
          "WHERE Investigation2.VisitationID = '" & VisitationID & "'"

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .EOF Then
            LabTests = "No Investigations Done"
            .Close
            Set rst = Nothing
            Exit Function
        End If

        output = "<ol>"
        Do While Not .EOF
            output = output & "<li>" & .fields("LabTestName") & "</li>"
            .MoveNext
        Loop
        output = output & "</ol>"
        .Close
    End With
    Set rst = Nothing

    LabTests = output
End Function


Function drugs(VisitationID)
    Dim sql, rst, output

    Set rst = CreateObject("ADODB.Recordset")
    sql = "SELECT Drug.DrugName " & _
          "FROM Drug " & _
          "JOIN Prescription ON Drug.DrugID = Prescription.DrugID " & _
          "WHERE Prescription.VisitationID = '" & VisitationID & "'"
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .EOF Then
            drugs = "No Drugs Prescribed"
            .Close
            Set rst = Nothing
            Exit Function
        End If

        output = "<ol>"
        Do While Not .EOF
            output = output & "<li>" & .fields("DrugName") & "</li>"
            .MoveNext
        Loop
        output = output & "</ol>"
        .Close
    End With
    
    Set rst = Nothing
    drugs = output
End Function



'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
