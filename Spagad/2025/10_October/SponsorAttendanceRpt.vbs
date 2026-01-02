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

'    sql = "SELECT DISTINCT Visitation.insuranceNo AS MembershipID,MAX(Visitation.VisitDate) AS VisitDate,"
'    sql = sql & " visitation.patientID, visitation.InsSchemeModeID, visitation.VisitTypeID, visitation.GenderID, Patient.ResidencePhone"
'    sql = sql & " FROM Visitation"
'    sql = sql & " INNER JOIN Patient ON visitation.patientID = Patient.PatientID"
'    sql = sql & " WHERE Visitation.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' AND Visitation.sponsorID = '" & SponsorID & "'"
'    sql = sql & " AND visitation.patientID NOT IN ('p1', 'p2', 'E01')"
'
'    If ageGroupID <> "" Then
'        sql = sql & "AND Visitation.AgeGroupID = '" & ageGroupID & "' "
'    Else
'        sql = sql & " "
'    End If
'
'    If specialistTypeID <> "" Then
'        sql = sql & "AND Visitation.specialistTypeID = '" & specialistTypeID & "' "
'    Else
'        sql = sql & " "
'    End If
'
'    sql = sql & " GROUP BY Visitation.insuranceNo, visitation.patientID, Patient.ResidencePhone, visitation.InsSchemeModeID, visitation.VisitTypeID, visitation.GenderID"

    sql = "WITH DoctorConsults AS ( "
    sql = sql & "SELECT DISTINCT SystemUserID, VisitationID "
    sql = sql & "FROM EMRRequestItems "
    sql = sql & "WHERE EMRDataID IN ('TH060', 'IM051') "
    sql = sql & "),"
    sql = sql & "PatientGroups AS ("
    sql = sql & "    SELECT "
    sql = sql & "        vst.insuranceNo,"
    sql = sql & "        vst.patientID, "
    sql = sql & "        vst.InsSchemeModeID, "
    sql = sql & "        vst.VisitTypeID, "
    sql = sql & "        vst.GenderID,"
    sql = sql & "        vst.SpecialistTypeID,"
    sql = sql & "        ip.InitialDependantID,"
    sql = sql & "        pat.ResidencePhone,"
    sql = sql & "        MAX(vst.VisitDate) AS MaxVisitDate"
    sql = sql & "    FROM Visitation vst"
    sql = sql & "    INNER JOIN Patient pat ON vst.patientID = pat.PatientID"
    sql = sql & "    JOIN InsuredPatient ip ON vst.PatientID = ip.PatientID"
    sql = sql & "    WHERE vst.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "      AND vst.sponsorID = '" & SponsorID & "'"
    sql = sql & "      AND vst.patientID NOT IN ('p1', 'p2', 'E01') "
    If ageGroupID <> "" Then
        sql = sql & "      AND vst.AgeGroupID = '" & ageGroupID & "' "
    Else
        sql = sql & " "
    End If
    If specialistTypeID <> "" Then
        sql = sql & "      AND vst.specialistTypeID = '" & specialistTypeID & "' "
    Else
        sql = sql & " "
    End If
    sql = sql & "    GROUP BY vst.insuranceNo, "
    sql = sql & "             vst.patientID, "
    sql = sql & "             pat.ResidencePhone, "
    sql = sql & "             vst.InsSchemeModeID, "
    sql = sql & "             vst.VisitTypeID, "
    sql = sql & "             vst.GenderID,"
    sql = sql & "             ip.InitialDependantID, "
    sql = sql & "             vst.SpecialistTypeID"
    sql = sql & "),"
    sql = sql & "LatestVisitation AS ("
    sql = sql & "    SELECT pg.*,"
    sql = sql & "           vst.VisitationID,"
    sql = sql & "           ROW_NUMBER() OVER ("
    sql = sql & "               PARTITION BY pg.insuranceNo, pg.patientID, pg.InsSchemeModeID, pg.VisitTypeID, pg.SpecialistTypeID, pg.InitialDependantID"
    sql = sql & "               ORDER BY vst.VisitDate DESC"
    sql = sql & "           ) AS rn"
    sql = sql & "    FROM PatientGroups pg"
    sql = sql & "    INNER JOIN Visitation vst ON vst.patientID = pg.patientID"
    sql = sql & "        AND vst.VisitDate = pg.MaxVisitDate"
    sql = sql & "        AND vst.InsSchemeModeID = pg.InsSchemeModeID"
    sql = sql & "        AND vst.VisitTypeID = pg.VisitTypeID"
    sql = sql & "        AND vst.SpecialistTypeID = pg.SpecialistTypeID"
    sql = sql & "        AND vst.sponsorID = '" & SponsorID & "' "
    sql = sql & ")"
    sql = sql & "SELECT "
    sql = sql & "    lv.insuranceNo AS MembershipID, "
    sql = sql & "    lv.MaxVisitDate AS VisitDate, "
    sql = sql & "    lv.patientID, "
    sql = sql & "    lv.InsSchemeModeID, "
    sql = sql & "    lv.VisitTypeID, "
    sql = sql & "    lv.GenderID, "
    sql = sql & "    lv.ResidencePhone,"
    sql = sql & "    lv.SpecialistTypeID,"
    sql = sql & "    lv.InitialDependantID,"
    sql = sql & "    (SELECT TOP 1 Staff.StaffName"
    sql = sql & "     FROM DoctorConsults dc"
    sql = sql & "     JOIN SystemUser ON SystemUser.SystemUserID = dc.SystemUserID"
    sql = sql & "     JOIN Staff ON Staff.StaffID = SystemUser.StaffID"
    sql = sql & "     WHERE dc.VisitationID = lv.VisitationID"
    sql = sql & "    ) AS [Attending Doctor]"
    sql = sql & "FROM LatestVisitation lv "
    sql = sql & "WHERE lv.rn = 1;"
    
    With rst
    .open qryPro.FltQry(sql), conn, 3, 4

    If .recordCount > 0 Then

        response.write "<div class='working-year'></div>"
        response.write "<table id='myTable'> "
        response.write "    <thead> "
        response.write "        <tr class='h_title'> "
        response.write "            <td colspan='12'><b>ATTENDANCE REPORT FOR " & GetComboName("Sponsor", SponsorID) & " IN " & GetComboName("WorkingYear", workingYearID) & "</b></td>"
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
        response.write "            <th>CONSULTATION TYPE</th>"
        response.write "            <th>DEPENDANT</th>"
        response.write "            <th>CONSULTING DOCTOR</th>"
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
            response.write "            <td>" & GetComboName("SpecialistType", .fields("SpecialistTypeID")) & "</td>"
            response.write "            <td>" & GetComboName("InitialDependant", .fields("InitialDependantID")) & "</td>"
            response.write "            <td>" & .fields("Attending Doctor") & "</td>"
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
