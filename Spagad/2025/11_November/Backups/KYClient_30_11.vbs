'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim sql, cnt
Dim arPeriod, Gender, MaritalStatus, country, dateRange
Gender = Trim(Request.QueryString("PrintFilter1"))
MaritalStatus = Trim(Request.QueryString("PrintFilter2"))
country = Trim(Request.QueryString("PrintFilter3"))
dateRange = Split(Trim(Request.QueryString("PrintFilter0")), "||")

response.write "<style> "
response.write "table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse;padding: 5px;}"
response.write " table#myTable{width: 80vw;; margin: 0 auto; font-size: 13px; font-family: sans-serif;box-sizing: border-box; } "
response.write " table#myTable thead{ text-align: center;font-size:16px; } table#myTable thead th{padding: 4px;} "
response.write " table#myTable thead .h_res{ background-color: #FC046A; color:#fff } "
response.write " table#myTable thead .h_title{ background-color: blanchedalmond; } "
response.write " table#myTable thead .h_names{ font-size: 14px;} table#myTable tbody td{text-align:center;} "
response.write " table#myTable .last{background-color: #3C8F6D;color:#fff;font-weight:700;text-align:center;} "
response.write " table#myTable tr:hover{background-color: #E1BDAB;color:#000; cursor:pointer;} "
response.write " </style>"
kyc Gender, MaritalStatus, country

Sub kyc(Gender, MaritalStatus, country)
    sql = "SELECT Patient.Patientid, SurName, patientname, gendername, maritalstatusname, residencephone, countryname"
    sql = sql & " , occupation, residenceAddress, patientinfo1 "
    sql = sql & " FROM Patient "
    sql = sql & " LEFT JOIN Visitation"
    sql = sql & " ON Patient.PatientID = Visitation.PatientID"
    sql = sql & " JOIN Gender ON Gender.GenderID=Patient.genderid "
    sql = sql & " JOIN MaritalStatus ON MaritalStatus.MaritalStatusID=Patient.MaritalStatusID "
    sql = sql & " JOIN Country ON Country.countryid=Patient.CountryID "
    sql = sql & " WHERE Patient.Patientid NOT LIKE '%P%' AND surname NOT LIKE '%.%' AND surname NOT LIKE '%-%' "
    sql = sql & " AND LEN(LEFT(SurName, CHARINDEX(' ', SurName + ' ') - 1)) > 3 "
    If Len(Gender) > 0 Then
        sql = sql & " AND Patient.GenderID='" & Gender & "' "
    Else
        sql = sql & " "
    End If
    If Len(MaritalStatus) > 0 Then
        sql = sql & " AND Patient.MaritalStatusID='" & MaritalStatus & "' "
    Else
        sql = sql & " "
    End If
    If Len(country) > 0 Then
        sql = sql & " AND Patient.CountryID='" & country & "' "
    Else
        sql = sql & " "
    End If
   
    If IsArray(dateRange) And UBound(dateRange) >= 1 And Len(Trim(dateRange(0))) > 0 And Len(Trim(dateRange(1))) > 0 Then
        Dim formattedDates
        formattedDates = getDatePeriodFromDelim(Trim(Request.QueryString("PrintFilter0")))
        sql = sql & " AND Visitation.VisitDate BETWEEN '" & formattedDates(0) & "' AND '" & formattedDates(1) & "' "
    End If
    sql = sql & " ORDER BY patientname"
    
   
    
    Set rst0 = CreateObject("ADODB.Recordset")
    On Error Resume Next
    With rst0
        .open qryPro.FltQry(sql), conn, 0, 1
        If Err.number <> 0 Then
            response.write "SQL Error: " & Err.Description & "<br>Query: " & sql
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
        
        If .recordCount > 0 Then
            
            .MoveFirst
            Dim seenIDs
            Set seenIDs = CreateObject("Scripting.Dictionary")
            seenIDs.CompareMode = 1
            Do While Not .EOF
                Dim tempID
                tempID = Trim(.fields("Patientid").value)
                If Len(tempID) > 0 And Not seenIDs.Exists(tempID) Then
                    seenIDs.Add tempID, True
                End If
                .MoveNext
            Loop
            Dim uniqueCount
            uniqueCount = seenIDs.count
            
            
            .MoveFirst
            response.write "<table id='myTable'><thead><tr class='h_title'><td colspan='15'>"
            response.write "<b>" & uniqueCount & " unique patients</b></td></tr>"
            response.write "<tr class='h_names'><th>#</th><th>ID</th><th>FULL NAME</th>"
            response.write "<th>GENDER</th><th>CONTACT</th><th>EMAIL</th><th>ADDRESS</th>"
            response.write "<th>MARITAL STATUS</th><th>COUNTRY</th><th>OCCUPATION</th>"
            response.write "</tr></thead><tbody>"
            
            cnt = 0
            seenIDs.RemoveAll
            Do While Not .EOF
                patID = Trim(.fields("Patientid").value)
                If Len(patID) > 0 And Not seenIDs.Exists(patID) Then
                    seenIDs.Add patID, True
                    cnt = cnt + 1
                    patName = .fields("patientname").value
                    Gender = .fields("gendername").value
                    MaritalStatus = .fields("maritalstatusname").value
                    residencephone = .fields("residencephone").value
                    email = .fields("patientinfo1").value
                    address = .fields("residenceAddress").value
                    country = .fields("countryname").value
                    occupation = .fields("occupation").value
                    response.write "<tr><td>" & cnt & "</td> <td>" & patID & "</td>"
                    response.write "<td style='text-align:left;'>" & patName & "</td>"
                    response.write "<td>" & Gender & "</td>"
                    response.write "<td>" & residencephone & "</td>"
                    response.write "<td>" & email & "</td>"
                    response.write "<td>" & address & "</td>"
                    response.write "<td>" & MaritalStatus & "</td>"
                    response.write "<td>" & country & "</td>"
                    response.write "<td>" & occupation & "</td>"
                    response.write "</tr>"
                    response.flush
                End If
                .MoveNext
            Loop
        Else
            response.write "<p>No records found.</p>"
        End If
        response.write "</tbody></table>"
        .Close
    End With
    Set rst0 = Nothing
    Set seenIDs = Nothing
End Sub
Sub kyc01(Gender, MaritalStatus, country)
    sql = "SELECT DISTINCT Patient.Patientid, SurName, patientname, gendername, maritalstatusname, residencephone, countryname"
    sql = sql & " , occupation, residenceAddress, patientinfo1 "
    sql = sql & " FROM Patient "
    sql = sql & " LEFT JOIN Visitation"
    sql = sql & " ON Patient.PatientID = Visitation.PatientID"
    sql = sql & " JOIN Gender ON Gender.GenderID=Patient.genderid "
    sql = sql & " JOIN MaritalStatus ON MaritalStatus.MaritalStatusID=Patient.MaritalStatusID "
    sql = sql & " JOIN Country ON Country.countryid=Patient.CountryID "
    sql = sql & " WHERE Patient.Patientid NOT LIKE '%P%' AND surname NOT LIKE '%.%' AND surname NOT LIKE '%-%' "
    sql = sql & " AND LEN(LEFT(SurName, CHARINDEX(' ', SurName + ' ') - 1)) > 3 "

If Len(Gender) > 0 Then
    sql = sql & " AND Patient.GenderID='" & Gender & "' "
  Else
    sql = sql & " "
End If
If Len(MaritalStatus) > 0 Then
    sql = sql & " AND Patient.MaritalStatusID='" & MaritalStatus & "' "
  Else
    sql = sql & " "
End If
If Len(country) > 0 Then
    sql = sql & " AND Patient.CountryID='" & country & "' "
  Else
    sql = sql & " "
End If

'If Len(dateRange) > 0 Then
    sql = sql & " AND Visitation.VisitDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
'End If

    sql = sql & " ORDER BY patientname"

    Set rst0 = CreateObject("ADODB.Recordset")
    With rst0
        .open qryPro.FltQry(sql), conn, 0, 1
        If .recordCount > 0 Then
            response.write "<table id='myTable'><thead><tr class='h_title'><td colspan='15'>"
            response.write "<b>" & .recordCount & "</b></td></tr>"
            response.write "<tr class='h_names'><th>#</th><th>ID</th><th>FULL NAME</th>"
            response.write "<th>GENDER</th><th>CONTACT</th><th>EMAIL</th><th>ADDRESS</th>"
            response.write "<th>MARITAL STATUS</th><th>COUNTRY</th><th>OCCUPATION</th>"
            response.write "</tr></thead><tbody>"
        .MoveFirst
        Do While Not .EOF
            cnt = cnt + 1
            patID = .fields("Patientid")
            patName = .fields("patientname")
            Gender = .fields("gendername")
            MaritalStatus = .fields("maritalstatusname")
            residencephone = .fields("residencephone")
            email = .fields("patientinfo1")
            address = .fields("residenceAddress")
            country = .fields("countryname")
            occupation = .fields("occupation")
            response.write "<tr><td>" & cnt & "</td> <td>" & patID & "</td>"
            response.write "<td style='text-align:left;'>" & patName & "</td>"
            response.write "<td>" & Gender & "</td>"
            response.write "<td>" & residencephone & "</td>"
            response.write "<td>" & email & "</td>"
            response.write "<td>" & address & "</td>"
            response.write "<td>" & MaritalStatus & "</td>"
            response.write "<td>" & country & "</td>"
            response.write "<td>" & occupation & "</td>"
            response.write "</tr>"
            
            response.flush
        .MoveNext
        Loop
        End If
        response.write "</tbody></table>"
        rst0.Close
        Set rst0 = Nothing
    End With
End Sub

Function getDatePeriodFromDelim(strDelimPeriod)
        
    Dim arPeriod, periodStart, periodEnd

    Dim arOut(1)

    arPeriod = Split(strDelimPeriod, "||")

    If UBound(arPeriod) >= 0 Then
        periodStart = arPeriod(0)
    End If

    If UBound(arPeriod) >= 1 Then
        periodEnd = arPeriod(1)
    End If

    periodStart = makeDatePeriod(Trim(periodStart), periodEnd, "0:00:00")
    periodEnd = makeDatePeriod(Trim(periodEnd), periodStart, "23:59:59")

    arOut(0) = periodStart
    arOut(1) = periodEnd

    getDatePeriodFromDelim = arOut

End Function

Function makeDatePeriod(strDateStart, defaultDate, strTime)

    If IsDate(strDateStart) Then
        makeDatePeriod = FormatDate(strDateStart) & " " & Trim(strTime)
    Else

        If IsDate(defaultDate) Then
            makeDatePeriod = FormatDate(defaultDate) & " " & Trim(strTime)
        Else
            makeDatePeriod = FormatDate(Now()) & " " & Trim(strTime)
        End If
    End If

End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
