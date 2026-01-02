'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim rst, sql, cnt, arPeriod, periodStart, periodEnd, hrf
arPeriod = getDatePeriodFromDelim(Trim(Request.QueryString("PrintFilter")))
periodStart = arPeriod(0)
periodEnd = arPeriod(1)

Set rst = CreateObject("ADODB.RecordSet")


sql = "SELECT distinct visitationid, specialistid, patientid, sponsorid"
sql = sql & " FROM Prescription"
sql = sql & " WHERE prescriptionstatusid='P001' AND prescriptionDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
sql = sql & " ORDER BY visitationid desc"

cnt = 0
With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .recordCount > 0 Then
    response.write "<table id='myTable'> "
    response.write "    <thead> "
    response.write "        <tr class='h_res'> "
    response.write "            <th colspan='7'>Found " & rst.recordCount & " Results...</th> "
    response.write "        </tr> "
    response.write "        <tr class='h_title'> "
    response.write "            <th colspan='7'>Generated Unserved Prescription Report Between " & periodStart & " And " & periodEnd & "</th> "
    response.write "        </tr>"

    response.write "        <tr class='h_names'> "
    response.write "            <th>No.</th> "
    response.write "            <th>VisitID</th> "
    response.write "            <th>Client Name</th> "
    response.write "            <th>Description/Items</th> "
    response.write "            <th>Amt</th> "
    response.write "            <th>Sponsor</th> "
    response.write "            <th>Physician</th> "
    response.write "        </tr> "
    response.write "    </thead><tbody>"
    .MoveFirst

    Do While Not .EOF

      cnt = cnt + 1
      vid = .fields("visitationid")
      hrf = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation&VisitationID=" & vid & ""

      response.write "      <tr> "
      response.write "        <td>" & cnt & "</td> "
      ' response.write "        <td>" & .fields("visitationid") & "</td> "
      response.write "        <td><a target=""_Blank"" href=""" & hrf & """>" & vid & "</td> "
      response.write "        <td>" & GetComboName("Patient", .fields("patientid")) & "</td> "
      response.write "        <td>"
      DisplayProd vid
      response.write "        </td> "
      response.write "        <td>"
      DisplayAmt vid
      response.write "        </td> "
      response.write "        <td><small>" & GetComboName("Sponsor", .fields("sponsorid")) & "</small></td> "
      response.write "        <td><small>" & .fields("specialistid") & "</small></td> "
      response.write "      </tr>"
      .MoveNext
      
    Loop
  End If
  rst.Close
  Set rst = Nothing
End With
response.write "</tbody></table>"

Sub DisplayProd(vid)
 Dim rst, sql
 Set rst = CreateObject("ADODB.RecordSet")
 sql = "SELECT PrescriptionName,qty,prescribeinfo1 FROM Prescription WHERE visitationid='" & vid & "'"
 With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .recordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
      response.write "<small class='blk'><b>(" & .fields("Qty") & ")</b>X " & .fields("PrescriptionName") & " =<b>{" & .fields("prescribeinfo1") & "}</b></small><hr>"
      .MoveNext
    Loop
  End If
  .Close
 End With
 Set rst = Nothing
End Sub

Sub DisplayAmt(vid)
 Dim rst, sql, amt
 Set rst = CreateObject("ADODB.RecordSet")
 sql = "SELECT sum(FinalAmt) AS amt FROM Prescription WHERE visitationid='" & vid & "'"
 With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .recordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
        amt = .fields("amt")
        response.write amt
      .MoveNext
    Loop
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

response.write "<style>table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 90vw; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{ background-color: #FC046A; color:#fff } table#myTable thead .h_title{ background-color: blanchedalmond; } table#myTable thead .h_names{ font-size: 14px;} table#myTable th:last-child{background-color: #3C8F6D;color:#fff} table#myTable small.blk{display:block;margin:0;padding:0;}</style>"

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
