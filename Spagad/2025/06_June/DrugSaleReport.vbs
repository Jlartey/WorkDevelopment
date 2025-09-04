'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim rst, sql, cnt, insTyp
Dim arPeriod, periodStart, periodEnd
arPeriod = getDatePeriodFromDelim(Trim(Request.queryString("PrintFilter")))

periodStart = arPeriod(0)
periodEnd = arPeriod(1)

'wkdid = Trim(request.QueryString("PrintFilter"))
Set rst = CreateObject("ADODB.RecordSet")

response.write "<style>table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 90vw; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{ background-color: #FC046A; color:#fff } table#myTable thead .h_title{ background-color: blanchedalmond; } table#myTable thead .h_names{ font-size: 14px;} table#myTable th:last-child{background-color: #3C8F6D;color:#fff} table#myTable small.blk{display:block;margin:0;padding:0;}</style>"

sql = "WITH SaleItems1 AS ("
sql = sql & " SELECT drugid, SUM(qty) AS qt, SUM(finalamt) AS amt"
sql = sql & " FROM DrugsaleItems WHERE dispensedate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' GROUP BY drugid"
sql = sql & "), ReturnItems1 AS ("
sql = sql & " SELECT drugid, -SUM(returnqty) AS qt, -SUM(finalamt) AS amt"
sql = sql & " FROM DrugReturnItems WHERE returndate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' GROUP BY drugid"
sql = sql & "), SaleItems2 AS ("
sql = sql & " SELECT drugid, SUM(DispenseAmt1) AS qt, SUM(DispenseAmt2) AS amt"
sql = sql & " FROM DrugSaleItems2 WHERE dispensedate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' GROUP BY drugid"
sql = sql & "), ReturnItems2 AS ("
sql = sql & " SELECT drugid, -SUM(returnqty) AS qt, -SUM(MainItemValue1) AS amt"
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

With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    response.write "<table id='myTable'> "
    response.write "<thead>"
    response.write "<tr class='h_res'>"
    response.write "<th colspan='7'>Found " & rst.RecordCount & " Results...</th>"
    response.write "</tr>"

    response.write "<tr class='h_title'>"
    response.write "<th colspan='7'>Generated Drug Dispensed Report From [" & periodStart & " to " & periodEnd & "]</th>"
    response.write "</tr>"

    response.write "<tr class='h_names'>"
    response.write "<th>No.</th>"
    response.write "<th>DrugID</th>"
    response.write "<th>Drug Name</th>"
    response.write "<th>Qty</th>"
    response.write "<th>Final Amount</th>"
    response.write " </tr>"
    response.write "</thead>"
    response.write "<tbody>"
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
      response.write "<tr>"
            response.write "<td>" & cnt & "</td>"
            response.write "<td>" & .fields("drugid") & "</td>"
            response.write "<td>" & GetComboName("Drug", .fields("drugid")) & "</td>"
            response.write "<td>" & qt & "</td>"
            response.write "<td>" & (FormatNumber(CStr(amt), 2, , , -1)) & "</td>"
      response.write "</tr>"
      .MoveNext
    Loop
    response.write "<tr>"
      response.write "<td colspan='3'><b>TOTAL</b></td>"
      response.write "<td>" & (FormatNumber(CStr(totQt), 2, , , -1)) & "</td>"
      response.write "<td><b>" & (FormatNumber(CStr(totAmt), 2, , , -1)) & "</b></td>"
    response.write "</tr>"
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
