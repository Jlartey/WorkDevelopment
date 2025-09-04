'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim rst, sql, cnt, wkdid
Dim arPeriod, periodStart, periodEnd
arPeriod = getDatePeriodFromDelim(Trim(Request.QueryString("PrintFilter0")))
periodStart = arPeriod(0)
periodEnd = arPeriod(1)

'wkdid = Trim(Request.QueryString("PrintFilter"))
Set rst = CreateObject("ADODB.RecordSet")

response.write "<style> table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 80vw; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{ background-color: #FC046A; color:#fff } table#myTable thead .h_title{ background-color: blanchedalmond; } table#myTable thead .h_names{ font-size: 14px;} table#myTable th:last-child{background-color: #3C8F6D;color:#fff} small.blk{display:block;width:100%;text-align:center;align-items:center;}  </style>"

sql = "SELECT patientid,visitdate,visitationid,systemuserid,sponsorid,insuranceno,medicaloutcomeid,specialistid,specialisttypeid,visittypeid"
sql = sql & " FROM Visitation"
sql = sql & " WHERE patientID <> 'P1' AND visitationid NOT LIKE '%-C%'"
sql = sql & " AND visitdate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
sql = sql & " AND visitationid NOT LIKE '%-C'"
cnt = 0
With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    response.write "<table id='myTable'> <thead><tr class='h_res'><th colspan='9'>Found " & rst.RecordCount & " Results...</th></tr>"
    response.write "<tr class='h_title'><th colspan='9'>Generated Daily OPD Registration Report from " & periodStart & " to " & periodEnd & "</th></tr>"

    response.write "<tr class='h_names'><th>No.</th><th>Visit Info</th><th>Patient Detail</th><th>Sponsor Detail</th><th>Visit Type</th><th>Specialist</th><th>Med. Outcome</th><th>User</th></tr></thead><tbody>"
    .MoveFirst
    Do While Not .EOF
      cnt = cnt + 1
      vdt = .fields("visitdate")
      patid = .fields("patientid")
      vID = .fields("visitationid")
      spid = .fields("sponsorid")
      insno = .fields("insuranceno")
      medOt = .fields("medicaloutcomeid")
      spltid = .fields("specialistid")
      splttyp = .fields("specialisttypeid")
      vstyp = .fields("visittypeid")
      sid = .fields("systemuserid")
      response.write "<tr><td>" & cnt & "</td><td><small class='blk'>" & vdt & "</small> <small class='blk'>" & vID & "</small></td><td>" & GetComboName("Patient", patid) & "<small class='blk'><b>" & patid & "</b></small></td> <td><small class='blk'><b>" & GetComboName("Sponsor", spid) & "</b></small><small class='blk'>" & insno & "</small></td> <td><small class='blk'><b>" & GetComboName("SpecialistType", splttyp) & "</b></small><small class='blk'>" & GetComboName("VisitType", vstyp) & "</small></td> <td>" & GetComboName("Specialist", spltid) & "</td> <td>" & GetComboName("MedicalOutcome", medOt) & "</td><td><small>" & sid & "</small></td></tr>"
      .MoveNext
    Loop
  End If
  rst.Close
  Set rst = Nothing
End With
response.write "</tbody></table>"


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
