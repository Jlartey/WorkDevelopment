'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
'SponsorAttendanceRpt
response.buffer = True
addCSS
generateReport

Sub generateReport()
    Dim rst, sql, cnt, SponsorID, datePeriod, specialistTypeID, ageGroupID
    SponsorID = Request.QueryString("PrintFilter")
    datePeriod = Request.QueryString("PrintFilter1")
    specialistTypeID = Request.QueryString("PrintFilter2")
    ageGroupID = Request.QueryString("PrintFilter3")
    Set rst = CreateObject("ADODB.RecordSet")

    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If

    cnt = 0

    ' sql = "SELECT DISTINCT Visitation.insuranceNo AS MembershipID,MAX(Visitation.VisitDate) AS VisitDate,"
    ' sql = sql & " visitation.patientID, visitation.InsSchemeModeID, Patient.ResidencePhone"
    ' sql = sql & " FROM Visitation"
    ' sql = sql & " INNER JOIN Patient ON visitation.patientID = Patient.PatientID"
    ' sql = sql & " WHERE Visitation.WorkingYearID = '" & workingyearID & "' AND Visitation.sponsorID = '" & SponsorID & "'"
    ' sql = sql & " AND visitation.patientID NOT IN ('p1', 'p2', 'E01')"
    ' sql = sql & " GROUP BY Visitation.insuranceNo, visitation.patientID, Patient.ResidencePhone, visitation.InsSchemeModeID"

     sql = "SELECT DISTINCT Visitation.insuranceNo AS MembershipID,MAX(Visitation.VisitDate) AS VisitDate,"
    sql = sql & " visitation.patientID, visitation.InsSchemeModeID, visitation.VisitTypeID, visitation.GenderID, Patient.ResidencePhone"
    sql = sql & " FROM Visitation"
    sql = sql & " INNER JOIN Patient ON visitation.patientID = Patient.PatientID"
    sql = sql & " WHERE Visitation.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' AND Visitation.sponsorID = '" & SponsorID & "'"
    sql = sql & " AND visitation.patientID NOT IN ('p1', 'p2', 'E01')"

    If ageGroupID <> "" Then
        sql = sql & "AND Visitation.AgeGroupID = '" & ageGroupID & "' "
    Else
        sql = sql & " "
    End If

    If specialistTypeID <> "" Then
        sql = sql & "AND Visitation.specialistTypeID = '" & specialistTypeID & "' "
    Else
        sql = sql & " "
    End If

    sql = sql & " GROUP BY Visitation.insuranceNo, visitation.patientID, Patient.ResidencePhone, visitation.InsSchemeModeID, visitation.VisitTypeID, visitation.GenderID"


    With rst
    .open qryPro.FltQry(sql), conn, 3, 4

    If .recordCount > 0 Then

        response.write "<div class='working-year'></div>"
        response.write "<table id='myTable'> "
        response.write "    <thead> "
        response.write "        <tr class='h_title'> "
        response.write "            <td colspan='9'><b>ATTENDANCE REPORT FOR " & GetComboName("Sponsor", SponsorID) & " IN " & GetComboName("WorkingYear", workingYearID) & "</b></td>"
        response.write "        </tr> "
        response.write "        <tr class='h_names'> "
        response.write "            <th>#</th> "
        response.write "            <th>DATE</th>"
        response.write "            <th>NAME</th> "
        response.write "            <th>PATIENT ID</th> "
        response.write "            <th>GENDER</th> "
        response.write "            <th>INSURANCE ID</th> "
        response.write "            <th>DEPARTMENT</th> "
        response.write "            <th>CONTACT</th> "
        response.write "            <th>FIRST/FOLLOW UP</th>"
        response.write "        </tr> "
        response.write "    </thead><tbody>"
        .MoveFirst

        Do While Not .EOF
            cnt = cnt + 1

            response.write "        <tr> "
            response.write "            <td>" & cnt & "</td> "
            response.write "            <td>" & FormatDate(.fields("VisitDate")) & "</td> "
            response.write "            <td>" & GetComboName("Patient", .fields("patientID")) & "</td> "
            response.write "            <td>" & .fields("patientID") & "</td> "
            response.write "            <td>" & GetComboName("Gender", .fields("GenderID")) & "</td> "
            response.write "            <td>" & .fields("MembershipID") & "</td>"
            response.write "            <td>" & GetComboName("InsSchemeMode", .fields("InsSchemeModeID")) & "</td>"
            response.write "            <td>" & .fields("ResidencePhone") & "</td>"
            response.write "            <td>" & GetComboName("VisitType", .fields("VisitTypeID")) & "</td>"
            response.write "        </tr>"
            .MoveNext
        Loop
        response.write "</tbody></table><br><br>"
    End If
    .Close
    End With
    Set rst = Nothing
End Sub

Sub addCSS()
response.write "<style> "
    response.write "    table#myTable, table#myTable th, table#myTable td { "
    response.write "        border: 1px solid silver; "
    response.write "        border-collapse: collapse; "
    response.write "        padding: 5px; "
    response.write "    } "
    response.write "    .working-year {"
    response.write "        text-align: center; "
    response.write "    }"
    response.write "    table#myTable { "
    response.write "        width: 80vw; "
    response.write "        margin: 0 auto; "
    response.write "        font-size: 13px; "
    response.write "        font-family: sans-serif; "
    response.write "        box-sizing: border-box; "
    response.write "    } "
    response.write "    table#myTable thead { "
    response.write "        text-align: center; "
    response.write "        font-size: 16px; "
    response.write "    } "
    response.write "    table#myTable thead th { "
    response.write "        padding: 4px; "
    response.write "    } "
    response.write "    table#myTable thead .h_res { "
    response.write "        background-color: #FC046A; "
    response.write "        color:#fff; "
    response.write "    } "
    response.write "    table#myTable thead .h_title { "
    response.write "        background-color: blanchedalmond; "
    response.write "    } "
    response.write "    table#myTable thead .h_names { "
    response.write "        position: sticky;"
    response.write "        top: 0;"
    response.write "        background-color: blanchedalmond;"
    response.write "    }"
    response.write "        font-size: 14px; "
    response.write "    } "
    response.write "    table#myTable tbody td { "
    response.write "        text-align:center; "
    response.write "    } "
    response.write "    table#myTable .last { "
    response.write "        background-color: #3C8F6D; "
    response.write "        color:#fff; "
    response.write "        font-weight: 700; "
    response.write "        text-align:center; "
    response.write "    } "
    response.write "</style>"
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
