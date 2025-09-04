'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
response.buffer = True
addCSS
generateReport

Sub generateReport()
    Dim rst1, rst, sql, sql1, cnt, sponsorID
    sponsorID = Request.querystring("PrintFilter")
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst1 = CreateObject("ADODB.RecordSet")

    cnt = 0
    
    sql1 = "SELECT DISTINCT workingyearID FROM visitation"
    sql1 = sql1 & " WHERE sponsorID = '" & sponsorID & "'"

    With rst1
    .open qryPro.FltQry(sql1), conn, 3, 4
    If .recordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            sql = "SELECT DISTINCT Visitation.insuranceNo AS MembershipID,MAX(Visitation.VisitDate) AS VisitDate,"
            sql = sql & " visitation.patientID, Patient.ResidencePhone"
            sql = sql & " FROM Visitation"
            sql = sql & " INNER JOIN Patient ON visitation.patientID = Patient.PatientID"
            sql = sql & " WHERE Visitation.WorkingYearID = '" & .fields("WorkingYearID") & "' AND Visitation.sponsorID = '" & sponsorID & "'"
            sql = sql & " GROUP BY Visitation.insuranceNo, visitation.patientID, Patient.ResidencePhone"
        
            response.Write "<div class='working-year'>" & GetComboName("WorkingYear", .fields("workingyearID")) & "</div>"
            response.Write "<table id='myTable'> "
            response.Write "    <thead> "
            response.Write "        <tr class='h_title'> "
            response.Write "            <td colspan='6'><b>ATTENDANCE REPORT FOR " & GetComboName("Sponsor", sponsorID) & "</b></td>"
            response.Write "        </tr> "
            response.Write "        <tr class='h_names'> "
            response.Write "            <th>#</th> "
            response.Write "            <th>DATE</th>"
            response.Write "            <th>NAME</th> "
            response.Write "            <th>PATIENT ID</th> "
            response.Write "            <th>INSURANCE ID</th> "
            response.Write "            <th>CONTACT</th> "
            response.Write "        </tr> "
            response.Write "    </thead><tbody>"

            With rst
            .open qryPro.FltQry(sql), conn, 3, 4
            If .recordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cnt = cnt + 1

                response.Write "        <tr> "
                response.Write "            <td>" & cnt & "</td> "
                response.Write "            <td>" & FormatDateDetail(.fields("VisitDate")) & "</td> "
                response.Write "            <td>" & GetComboName("Patient", .fields("patientID")) & "</td> "
                response.Write "            <td>" & .fields("patientID") & "</td> "
                response.Write "            <td><center>" & .fields("MembershipID") & "</center></td>"
                response.Write "            <td><center>" & .fields("ResidencePhone") & "</center></td>"
                response.Write "        </tr>"
                .MoveNext
                response.flush
            Loop
            response.Write "</tbody></table><br><br>"
            .Close
            End If
            End With
            .MoveNext
            response.flush
        Loop
    End If
    .Close
    End With
    Set rst = Nothing
End Sub

Sub addCSS()
response.Write "<style> "
    response.Write "    table#myTable, table#myTable th, table#myTable td { "
    response.Write "        border: 1px solid silver; "
    response.Write "        border-collapse: collapse; "
    response.Write "        padding: 5px; "
    response.Write "    } "
    response.Write "    .working-year {"
    response.Write "        text-align: center; "
    response.Write "    }"
    response.Write "    table#myTable { "
    response.Write "        width: 80vw; "
    response.Write "        margin: 0 auto; "
    response.Write "        font-size: 13px; "
    response.Write "        font-family: sans-serif; "
    response.Write "        box-sizing: border-box; "
    response.Write "    } "
    response.Write "    table#myTable thead { "
    response.Write "        text-align: center; "
    response.Write "        font-size: 16px; "
    response.Write "    } "
    response.Write "    table#myTable thead th { "
    response.Write "        padding: 4px; "
    response.Write "    } "
    response.Write "    table#myTable thead .h_res { "
    response.Write "        background-color: #FC046A; "
    response.Write "        color:#fff; "
    response.Write "    } "
    response.Write "    table#myTable thead .h_title { "
    response.Write "        background-color: blanchedalmond; "
    response.Write "    } "
    response.Write "    table#myTable thead .h_names { "
    response.Write "        position: sticky;"
    response.Write "        top: 0;"
    response.Write "        background-color: blanchedalmond;"
    response.Write "    }"
    response.Write "        font-size: 14px; "
    response.Write "    } "
    response.Write "    table#myTable tbody td { "
    response.Write "        text-align:center; "
    response.Write "    } "
    response.Write "    table#myTable .last { "
    response.Write "        background-color: #3C8F6D; "
    response.Write "        color:#fff; "
    response.Write "        font-weight: 700; "
    response.Write "        text-align:center; "
    response.Write "    } "
    response.Write "</style>"
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
