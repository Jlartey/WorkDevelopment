'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim rst, sql, cnt, sponsorType
sponsorType = Trim(request.QueryString("PrintFilter"))
Set rst = CreateObject("ADODB.Recordset")
response.write "<style> table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 80vw;; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{ background-color: #FC046A; color:#fff } table#myTable thead .h_title{ background-color: blanchedalmond; } table#myTable thead .h_names{ font-size: 14px;} table#myTable .last{background-color: #3C8F6D;color:#fff;font-weight:700;text-align:center;}  </style>"

sql = "SELECT sponsorid, sponsorname FROM Sponsor  where sponsorTypeID = '" & sponsorType & "' ORDER BY sponsorname asc"

cnt = 0
With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    response.write "<table id='myTable'> <thead><tr class='h_res'><th colspan='3'>Found " & rst.RecordCount & " Results...</th></tr> "
    response.write "<tr class='h_title'><th colspan='3'>Generated Sponsor List</th></tr>"
    response.write "<tr class='h_names'><th>No.</th><th>SponsorID</th><th>Sponsor Name</th></tr></thead><tbody>"
  .MoveFirst
  Do While Not .EOF
  cnt = cnt + 1
  response.write "<tr><td align='center'>" & cnt & "</td><td align='center'>" & .fields("sponsorid") & "</td><td>" & .fields("sponsorname") & "</td></tr>"
  .MoveNext
  Loop
  Else
    response.write "<p style='font-family: Arial, sans-serif; font-size: 1rem; color: red;'> No records found </p>"
  End If
  rst.Close
  Set rst = Nothing
End With
response.write "</tbody></table>"
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
