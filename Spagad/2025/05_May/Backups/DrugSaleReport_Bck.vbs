'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim rst, sql, cnt, insTyp
Dim arPeriod, periodStart, periodEnd
arPeriod = getDatePeriodFromDelim(Trim(Request.queryString("PrintFilter")))

periodStart = arPeriod(0)
periodEnd = arPeriod(1)

'wkdid = Trim(request.QueryString("PrintFilter"))
Set rst = CreateObject("ADODB.RecordSet")

Response.write "<style>table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 90vw; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{ background-color: #FC046A; color:#fff } table#myTable thead .h_title{ background-color: blanchedalmond; } table#myTable thead .h_names{ font-size: 14px;} table#myTable th:last-child{background-color: #3C8F6D;color:#fff} table#myTable small.blk{display:block;margin:0;padding:0;}</style>"

sql = "SELECT data.drugid, SUM(data.qt) AS qt, SUM(data.amt) AS amt FROM( "
sql = sql & " SELECT drugid, SUM(qty) AS qt, SUM(finalamt) AS amt "
sql = sql & " FROM DrugsaleItems where dispensedate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' GROUP BY drugid Union All "
sql = sql & " SELECT drugid, -SUM(returnqty) AS qt, -SUM(finalamt) AS amt "
sql = sql & " FROM DrugReturnItems where returndate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' GROUP BY drugid Union All "
sql = sql & " SELECT drugid, SUM(qty) AS qt, SUM(finalamt) AS amt "
sql = sql & " FROM DrugSaleItems2 where dispensedate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' GROUP BY drugid Union All "
sql = sql & " SELECT drugid, -SUM(returnqty) AS qt, -SUM(finalamt) AS amt "
sql = sql & " FROM DrugReturnItems2 where returndate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' GROUP BY drugid "
sql = sql & " ) AS data GROUP BY drugid ORDER BY amt desc"

Response.write sql
cnt = 0
With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    Response.write "<table id='myTable'> <thead><tr class='h_res'><th colspan='7'>Found " & rst.RecordCount & " Results...</th></tr>"
    Response.write "<tr class='h_title'><th colspan='7'>Generated Drug Dispensed Report From [" & periodStart & " to " & periodEnd & "]</th></tr>"

    Response.write "<tr class='h_names'><th>No.</th><th>DrugID</th><th>Drug Name</th><th>Qty</th><th>Final Amount</th></tr></thead><tbody>"
    .movefirst
    Do While Not .EOF
      cnt = cnt + 1
      amt = .fields("amt")
      qt = .fields("qt")
      totQt = totQt + qt
      totAmt = totAmt + amt
      Response.write "<tr><td>" & cnt & "</td><td>" & .fields("drugid") & "</td><td>" & GetComboName("Drug", .fields("drugid")) & "</td><td>" & qt & "</td><td>" & (FormatNumber(CStr(amt), 2, , , -1)) & "</td></tr>"
      .MoveNext
    Loop
    Response.write "<tr><td colspan='3'><b>TOTAL</b></td><td>" & (FormatNumber(CStr(totQt), 2, , , -1)) & "</td><td><b>" & (FormatNumber(CStr(totAmt), 2, , , -1)) & "</b></td></tr>"
  End If
  rst.Close
  Set rst = Nothing
End With
Response.write "</tbody></table>"

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
