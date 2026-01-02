'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim sql, cnt
Dim arPeriod, periodStart, periodEnd
arPeriod = getDatePeriodFromDelim(Trim(Request.QueryString("PrintFilter0")))
periodStart = arPeriod(0)
periodEnd = arPeriod(1)

generateReport periodStart, periodEnd
addCSS

Sub generateReport(periodStart, periodEnd)
    Dim rst, sql, cnt, workingDayID, systemUserID, subtotalRecords, totalRecords, systemUserName
    Set rst = CreateObject("ADODB.Recordset")

    cnt = 0
    totalRecords = 0
    'workingDayID = Request.QueryString("PrintFilter")
    'workingDay = GetComboName("workingDay", workingDayID)

    sql = "SELECT count(DISTINCT(patientid)) AS patients, systemuserid FROM Visitation"
    sql = sql & " WHERE visitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' AND visitationid NOT LIKE '%-C%' "
    sql = sql & " GROUP BY systemuserid"
    




    With rst
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        response.write "<table id='myTable'> "
        response.write "    <thead> "
        response.write "        <tr class='h_title'> "
        response.write "            <td colspan='15'> "
        response.write "                <b>OFFICER'S BILLING REPORT BETWEEN " & periodStart & " and " & periodEnd & "</b> "
        response.write "            </td></tr> "
        response.write "        <tr class='h_names'> "
        response.write "            <th>#</th> "
        response.write "            <th>NAME</th> "
        response.write "            <th>DAILY BILL REPORT COUNT</th> "
        response.write "            <th>VIEW MORE</th> "
        response.write "        </tr> "
        response.write "    </thead><tbody>"
        .MoveFirst

        Do While Not .EOF
            cnt = cnt + 1
            subtotalRecords = .fields("patients")
            systemUserID = .fields("systemuserid")
            systemUserName = Replace(systemUserID, ".", " ")
            totalRecords = totalRecords + subtotalRecords

            hrf = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=recordsDailyBillingDetail&PositionForTableName=WorkingDay&PrintFilter=" & systemUserID & "&printfilter1=" & periodStart & "||" & periodEnd

            response.write "        <tr> "
            response.write "            <td>" & cnt & "</td> "
            response.write "            <td>" & systemUserName & "</td> "
            response.write "            <td>" & subtotalRecords & "</td>"
            response.write "            <td><a href='" & hrf & "' target='_blank'>View More</a></td> "
            response.write "        </tr>"
            response.flush
            'totalRecords = totalRecords + getPatCount
            .MoveNext
        Loop
            response.write "        <tr> "
            response.write "            <td colspan='2'><b>TOTAL </b></td> "
            response.write "            <td><b>" & totalRecords & "</b></td> "
            response.write "            <td></td> "
            response.write "        </tr>"
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
    response.write "    table#myTable { "
    response.write "        width: 80vw; "
    response.write "        margin: 0 auto; "
    response.write "        font-size: 13px; "
    response.write "        font-family: sans-serif; "
    response.write "        box-sizing: border-box; "
    response.write "    } "
    response.write "    table#myTable thead { "
    response.write "        text-align: center; "
    response.write "        font-size:16px; "
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
