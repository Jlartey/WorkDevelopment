'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim rst, sql, cnt, gen
Dim arPeriod, periodStart, periodEnd
arPeriod = getDatePeriodFromDelim(Trim(Request.QueryString("PrintFilter1")))
periodStart = arPeriod(0)
periodEnd = arPeriod(1)

gen = Trim(Request.QueryString("PrintFilter0"))
Set rst = CreateObject("ADODB.Recordset")
response.write "<style> table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 80vw;; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{ background-color: #FC046A; color:#fff } table#myTable thead .h_title{ background-color: blanchedalmond; } table#myTable thead .h_names{ font-size: 14px;} table#myTable tbody td{text-align:center;} table#myTable .last{background-color: #3C8F6D;color:#fff;font-weight:700;text-align:center;}  </style>"

sql = "SELECT DISTINCT Visitation.PatientID, Patient.ResidencePhone, Visitation.VisitDate, visitation.specialisttypeid "
sql = sql & " FROM Visitation INNER JOIN patient ON Visitation.PatientID = Patient.PatientID WHERE "
If Len(gen) > 0 Then
sql = sql & " Visitation.GenderID = '" & gen & "' AND "
End If
sql = sql & " Visitation.visitdate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
sql = sql & " ORDER BY Visitation.VisitDate"
cnt = 0
With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    response.write "<table id='myTable'> <thead><tr class='h_res'><th colspan='15'>Found " & rst.RecordCount & " Results...</th></tr> "
    response.write "<tr class='h_title'><th colspan='15'>Generated Attendance Report " & GetComboName("Gender", gen) & "</th></tr>"
    response.write "<tr class='h_names'><th>No.</th><th>VisitDate</th><th>Full Name</th><th>Contact Number</th><th>Cons. Type</th></tr></thead><tbody>"
  .MoveFirst
  Do While Not .EOF
  cnt = cnt + 1
  response.write "<tr><td>" & cnt & "</td><td>" & .fields("VisitDate") & "</td><td>" & GetComboName("Patient", .fields("PatientID")) & "</td>"
  response.write "<td>" & .fields("ResidencePhone") & "</td><td>" & GetComboName("SpecialistType", .fields("specialisttypeid")) & "</td></tr>"
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
