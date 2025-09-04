'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim medicalOutcome
medicalOutcome = Trim(Request.querystring("MedicalOutcomeID"))
If IsNull(medicalOutcome) Or medicalOutcome = "" Then
    medicalOutcome = ""
End If

tableStyles
Header
MortalityReport

Sub Header()
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

Response.write "<div class='filters'>"
    Response.write "        <label for='medicalOutcome' class='font-style'>Select Medical Outcome:</label><br>"
    Response.write "        <select id='medicalOutcome' name='medicalOutcome'>"
    Response.write dropdownOptions
    Response.write "        </select>"
    Response.write "        <button type='button' onclick='updateUrl()'>Show Data</button> <br />"
Response.write "</div>"

Response.write "<script>"
    Response.write "    function updateUrl() {"
    
    Response.write "        const medicalOutcomes = Array.from(document.getElementById('medicalOutcome').selectedOptions).map(option => option.value).join(',');"
    Response.write "        const baseUrl = 'http://172.19.0.36/hms/wpgPrtPrintLayoutAll.asp';"
    Response.write "        const params = new URLSearchParams({"
    Response.write "            PrintLayoutName: 'MortalityReport',"
    Response.write "            PositionForTableName: 'WorkingDay',"
    Response.write "            WorkingDayID: '',"
    Response.write "            MedicalOutcomeID: medicalOutcomes"
    Response.write "        });"
    Response.write "        const newUrl = baseUrl + '?' + params.toString();"
    Response.write "        window.location.href = newUrl;"
    Response.write "        console.log(newUrl);"
    Response.write "    }"
    Response.write "</script>"

End Sub

Sub MortalityReport()
    Dim count, sql, rst, visitationID, emrRequestID, recordCount
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
                Response.write "<h1>Showing All " & recordCount & " Records</h1>"
            Else
                Response.write "<h1>Showing Data Of " & recordCount & " Patients Whose Medical Outcome = " & .fields("Outcome") & "</h1>"
            End If
            
            Response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            Response.write "<tr class='mytr'>"
            Response.write "<th class='myth'>No.</th>"
            Response.write "<th class='myth'>Visitation ID</th>"
            Response.write "<th class='myth'>Patient</th>"
            Response.write "<th class='myth'>Gender</th>"
            Response.write "<th class='myth'>Outcome</th>"
            Response.write "<th class='myth'>Age</th>"
            Response.write "<th class='myth'>Occupation</th>"
            Response.write "<th class='myth'>Marital Status</th>"
            Response.write "<th class='myth'>Residential Address</th>"
            Response.write "<th class='myth'>Nationality</th>"
            Response.write "<th class='myth'>Type Of Insurance</th>"
            Response.write "<th class='myth'>Admission Date</th>"
            Response.write "<th class='myth'>Ward Admitted To</th>"
            Response.write "<th class='myth'>Date Of Discharge</th>"
            Response.write "<th class='myth'>Diagnosis</th>"
            Response.write "<th class='myth'>Investigation Names</th>"
            Response.write "<th class='myth'>Drugs</th>"
            
            Response.write "<th class='myth'>Doctor</th>"
            Response.write "</tr class='mytr'>"
            
            Response.Flush
            
            Do While Not .EOF
                visitationID = .fields("VisitationID")
                emrRequestID = .fields("EMRRequestID")
                Response.write "<tr class='mytr'>"
                Response.write "<td class='mytd'>" & count & "</td>"
                Response.write "<td class='mytd'>" & visitationID & "</td>"
                Response.write "<td class='mytd' style='min-width: 200px;'>" & .fields("Patient") & "</td>"
                Response.write "<td class='mytd'>" & .fields("Gender") & "</td>"
                Response.write "<td class='mytd'>" & .fields("Outcome") & "</td>"
                Response.write "<td class='mytd'>" & .fields("Age") & "</td>"
                Response.write "<td class='mytd'>" & .fields("Occupation") & "</td>"
                Response.write "<td class='mytd'>" & .fields("Marital Status") & "</td>"
                Response.write "<td class='mytd'>" & .fields("Residence Address") & "</td>"
                Response.write "<td class='mytd'>" & .fields("Country") & "</td>"
                Response.write "<td class='mytd'>" & .fields("SponsorName") & "</td>"
                Response.write "<td class='mytd'>" & .fields("Admission Date") & "</td>"
                Response.write "<td class='mytd'>" & .fields("Ward") & "</td>"
                Response.write "<td class='mytd'>" & .fields("Discharge Date") & "</td>"
                Response.write "<td class='mytd'>" & diagnosis(visitationID) & "</td>"
                Response.write "<td class='mytd'>" & LabTests(visitationID) & "</td>"
                Response.write "<td class='mytd'>" & drugs(visitationID) & "</td>"
                
                Response.write "<td class='mytd' style='min-width: 150px;'>" & .fields("Doctor") & "</td>"
                Response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
                
                Response.Flush
            Loop

            Response.write "</table>"
        Else
            Response.write "<h1>No records found</h1>"
        End If
        
        .Close
    End With
    
    Set rst = Nothing
End Sub


Sub tableStyles()
    Response.write "<style>"
        Response.write ".mytable {"
        Response.write "    width: fit-content;"
        Response.write "    border-collapse: collapse;"
        Response.write "    margin: 50px 50px;"
        Response.write "    font-size: 16px;"
        Response.write "    font-family: Arial, sans-serif;"
        Response.write "}"
        Response.write ".mytable, .myth, .mytd {"
        Response.write "    border: 1px solid #dddddd;"
        Response.write "}"
        Response.write ".myth, .mytd {"
        Response.write "    padding: 12px;"
        Response.write "    text-align: left;"
        Response.write "}"
        Response.write ".myth {"
        Response.write "    background-color: #f2f2f2;"
        Response.write "    color: #333;"
        Response.write "    font-weight: bold;"
        Response.write "}"
        Response.write ".mytr:nth-child(even) {"
        Response.write "    background-color: #f9f9f9;"
        Response.write "}"
        Response.write ".mytr:hover {"
        Response.write "    background-color: #f1f1f1;"
        Response.write "}"
        Response.write ".myth {"
        Response.write "    text-transform: uppercase;"
        Response.write "}"
        Response.write "h1 {"
        Response.write "    font-size: 22px;"
        Response.write "    color: #555;"
        Response.write "    font-family: Arial, sans-serif;"
        Response.write "    margin: 20px 0;"
        Response.write "    text-transform: uppercase;"
        Response.write "}"
        Response.write "    .filters {"
        Response.write "        padding: 10px;"
        Response.write "        border: 1px solid #ccc;"
        Response.write "        background-color: #f9f9f9;"
        Response.write "        border-radius: 8px;"
        Response.write "        width: 300px;"
        Response.write "        margin-bottom: 15px;"
        Response.write "        margin-top: 30px; /* Added 30px padding on top */"
        Response.write "    }"
        
        Response.write "    .font-style {"
        Response.write "        font-family: Arial, sans-serif;"
        Response.write "        font-size: 14px;"
        Response.write "        font-weight: bold;"
        Response.write "        color: #333;"
        Response.write "    }"
        
        Response.write "    select {"
        Response.write "        width: 100%;"
        Response.write "        padding: 8px;"
        Response.write "        border-radius: 4px;"
        Response.write "        border: 1px solid #aaa;"
        Response.write "        font-size: 14px;"
        Response.write "        background-color: #fff;"
        Response.write "        cursor: pointer;"
        Response.write "    }"
        
        Response.write "    button {"
        Response.write "        margin-top: 10px;"
        Response.write "        padding: 8px 15px;"
        Response.write "        border: none;"
        Response.write "        background-color: #007BFF;"
        Response.write "        color: white;"
        Response.write "        font-size: 14px;"
        Response.write "        border-radius: 4px;"
        Response.write "        cursor: pointer;"
        Response.write "        transition: background 0.3s ease-in-out;"
        Response.write "    }"
        
        Response.write "    button:hover {"
        Response.write "        background-color: #0056b3;"
        Response.write "    }"

        Response.write "</style>"

End Sub

Function diagnosis(visitationID)
    Dim sql, rst, output, n

    Set rst = CreateObject("ADODB.Recordset")
    sql = "SELECT Disease.DiseaseName " & _
          "FROM Disease " & _
          "JOIN Diagnosis ON Diagnosis.DiseaseID = Disease.DiseaseID " & _
          "WHERE Diagnosis.VisitationID = '" & visitationID & "'"
    
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

Function LabTests(visitationID)
    Dim sql, rst, output

    Set rst = CreateObject("ADODB.Recordset")
    sql = "SELECT LabTest.LabTestName " & _
          "FROM LabTest " & _
          "JOIN Investigation ON LabTest.LabTestID = Investigation.LabTestID " & _
          "WHERE Investigation.VisitationID = '" & visitationID & "'" & _
          " UNION ALL " & _
          "SELECT LabTest.LabTestName " & _
          "FROM LabTest " & _
          "JOIN Investigation2 ON LabTest.LabTestID = Investigation2.LabTestID " & _
          "WHERE Investigation2.VisitationID = '" & visitationID & "'"

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


Function drugs(visitationID)
    Dim sql, rst, output

    Set rst = CreateObject("ADODB.Recordset")
    sql = "SELECT Drug.DrugName " & _
          "FROM Drug " & _
          "JOIN Prescription ON Drug.DrugID = Prescription.DrugID " & _
          "WHERE Prescription.VisitationID = '" & visitationID & "'"
    
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


