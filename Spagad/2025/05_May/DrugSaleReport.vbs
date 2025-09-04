'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Response.write "hello joe"
Dim rst, sql, cnt, insTyp
Dim arPeriod, periodStart, periodEnd
arPeriod = getDatePeriodFromDelim(Trim(Request.queryString("PrintFilter")))

periodStart = arPeriod(0)
periodEnd = arPeriod(1)

'wkdid = Trim(request.QueryString("PrintFilter"))
Set rst = CreateObject("ADODB.RecordSet")

Response.write "<style>table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 90vw; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{ background-color: #FC046A; color:#fff } table#myTable thead .h_title{ background-color: blanchedalmond; } table#myTable thead .h_names{ font-size: 14px;} table#myTable th:last-child{background-color: #3C8F6D;color:#fff} table#myTable small.blk{display:block;margin:0;padding:0;}</style>"

sql = "WITH SaleItems1 AS ("
sql = sql & " SELECT drugid, SUM(qty) AS qt, SUM(finalamt) AS amt"
sql = sql & " FROM DrugsaleItems WHERE dispensedate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' GROUP BY drugid"
sql = sql & "), ReturnItems1 AS ("
sql = sql & " SELECT drugid, -SUM(returnqty) AS qt, -SUM(finalamt) AS amt"
sql = sql & " FROM DrugReturnItems WHERE returndate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' GROUP BY drugid"
sql = sql & "), SaleItems2 AS ("
sql = sql & " SELECT drugid, SUM(qty) AS qt, SUM(finalamt) AS amt"
sql = sql & " FROM DrugSaleItems2 WHERE dispensedate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' GROUP BY drugid"
sql = sql & "), ReturnItems2 AS ("
sql = sql & " SELECT drugid, -SUM(returnqty) AS qt, -SUM(finalamt) AS amt"
sql = sql & " FROM DrugReturnItems2 WHERE returndate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' GROUP BY drugid"
sql = sql & "), CombinedData AS ("
sql = sql & " SELECT drugid, qt, amt FROM SaleItems1"
sql = sql & " UNION ALL"
sql = sql & " SELECT drugid, qt, amt FROM ReturnItems1"
sql = sql & " UNION ALL"
sql = sql & " SELECT drugid, qt, amt FROM SaleItems2"
sql = sql & " UNION ALL"
sql = sql & " SELECT drugid, qt, amt FROM ReturnItems2"
sql = sql & ") SELECT drugid, SUM(qt) AS qt, SUM(amt) AS amt FROM CombinedData GROUP BY drugid ORDER BY amt DESC"


Response.write sql
With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    Response.write "<table id='myTable'> "
    Response.write "<thead>"
    Response.write "<tr class='h_res'>"
    Response.write "<th colspan='7'>Found " & rst.RecordCount & " Results...</th>"
    Response.write "</tr>"

    Response.write "<tr class='h_title'>"
    Response.write "<th colspan='7'>Generated Drug Dispensed Report From [" & periodStart & " to " & periodEnd & "]</th>"
    Response.write "</tr>"

    Response.write "<tr class='h_names'>"
    Response.write "<th>No.</th>"
    Response.write "<th>DrugID</th>"
    Response.write "<th>Drug Name</th>"
    Response.write "<th>Qty</th>"
    Response.write "<th>Final Amount</th>"
    Response.write " </tr>"
    Response.write "</thead>"
    Response.write "<tbody>"
    .movefirst

    cnt = 0
    totQt = 0
    totalAmt = 0

    Do While Not .EOF
      cnt = cnt + 1
      amt = .fields("amt")
      qt = .fields("qt")
      totQt = totQt + qt
      totAmt = totAmt + amt
      Response.write "<tr>"
            Response.write "<td>" & cnt & "</td>"
            Response.write "<td>" & .fields("drugid") & "</td>"
            Response.write "<td>" & GetComboName("Drug", .fields("drugid")) & "</td>"
            Response.write "<td>" & qt & "</td>"
            Response.write "<td>" & (FormatNumber(CStr(amt), 2, , , -1)) & "</td>"
      Response.write "</tr>"
      .MoveNext
    Loop
    Response.write "<tr>"
      Response.write "<td colspan='3'><b>TOTAL</b></td>"
      Response.write "<td>" & (FormatNumber(CStr(totQt), 2, , , -1)) & "</td>"
      Response.write "<td><b>" & (FormatNumber(CStr(totAmt), 2, , , -1)) & "</b></td>"
    Response.write "</tr>"
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

