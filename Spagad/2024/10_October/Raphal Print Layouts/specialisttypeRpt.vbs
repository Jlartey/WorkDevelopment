'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim sql, cnt
Dim arPeriod, periodStart, periodEnd
arPeriod = getDatePeriodFromDelim(Trim(Request.QueryString("PrintFilter0")))
periodStart = arPeriod(0)
periodEnd = arPeriod(1)
response.write "<style> table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 80vw;; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center;font-size:16px; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{ background-color: #FC046A; color:#fff } table#myTable thead .h_title{ background-color: blanchedalmond; } table#myTable thead .h_names{ font-size: 14px;} table#myTable tbody td{text-align:center;} table#myTable .last{background-color: #3C8F6D;color:#fff;font-weight:700;text-align:center;}  </style>"

AllSpecialistType periodStart, periodEnd

Sub AllSpecialistType(periodStart, periodEnd)
    Dim fpats, fvisits, visits, patids, spid, patients, rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "SELECT specialisttypeid, COUNT(distinct(patientid)) AS visits "
    sql = sql & " FROM Visitation WHERE patientID <> 'P1' AND visitationid NOT LIKE '%-C%' "
    sql = sql & " AND visitdate between '" & periodStart & "' and '" & periodEnd & "' GROUP BY specialisttypeid"
    
    cnt = 0
    fvisits = 0
    patids = 0
    fpats = 0
    visits = 0
    With rst
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        response.write "<table id='myTable'> "
        response.write "    <thead> "
        response.write "        <tr class='h_title'> "
        response.write "            <td colspan='15'> "
        response.write "                <b>Departmental Visit Reports from [" & periodStart & "] to [" & periodEnd & "]</b> "
        response.write "            </td></tr> "
        response.write "        <tr class='h_names'> "
        response.write "            <th>#</th> "
        response.write "            <th>DEPARTMENT ID</th> "
        response.write "            <th>DEPARTMENT NAME</th>"
        ' response.write "            <th>PATIENTS[Count]</th> "
        response.write "            <th>VISIT[Count]</th>"
        response.write "            <th>VIEW MORE</th> "
        response.write "        </tr> "
        response.write "    </thead><tbody>"
      .MoveFirst

      Do While Not .EOF
        cnt = cnt + 1
        spid = .fields("specialisttypeid")
        visits = .fields("visits")
        ' patids = .fields("patients")
        
        hrf = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=specialistdetails&PositionForTableName=WorkingDay&PrintFilter=" & periodStart & "&PrintFilter1=" & periodEnd & "&PrintFilter2=" & spid

        response.write "        <tr> "
        response.write "            <td>" & cnt & "</td> "
        response.write "            <td>" & spid & "</td> "
        response.write "            <td style='text-align:left;'>" & GetComboName("specialisttype", spid) & "</td>"
        ' response.write "            <td>" & patids & "</td> "
        response.write "            <td>" & visits & "</td>"
        response.write "            <td><a href='" & hrf & "' target='_blank'>View More</a></td> "
        response.write "        </tr>"
        response.flush
        fpats = fpats + patids
        fvisits = fvisits + visits
      .MoveNext
      Loop
        response.write "        <tr> "
        response.write "            <td colspan='3'><b>TOTAL </b></td> "
        ' response.write "            <td>" & fpats & "</td>"
        response.write "            <td><b>" & fvisits & "</b></td> "
        response.write "            <td></td> "
        response.write "        </tr>"
        response.write "</tbody></table><br><br>"
      End If
    .Close
    End With
    Set rst = Nothing
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
