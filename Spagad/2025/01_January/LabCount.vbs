'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim datePeriod, dateArr

datePeriod = Trim(Request("PrintFilter"))
dateArr = Split(datePeriod, "||")
tableStyles
LabCount

Sub LabCount()
  Dim count, sql, rst
  count = 1

  Set rst = CreateObject("ADODB.Recordset")

  sql = "WITH InvestigationCount AS ("
  sql = sql & " SELECT i.PatientID, COUNT(i.LabTestID) AS LabTestCount "
  sql = sql & " FROM Investigation i "
  sql = sql & " WHERE i.RequestDate BETWEEN '" & dateArr(0) & "' AND '" & dateArr(1) & "' "
  sql = sql & " GROUP BY i.PatientID), "
  sql = sql & "Investigation2Count AS ("
  sql = sql & " SELECT i2.PatientID, COUNT(i2.LabTestID) AS LabTestCount "
  sql = sql & " FROM Investigation2 i2 "
  sql = sql & " WHERE i2.RequestDate BETWEEN '" & dateArr(0) & "' AND '" & dateArr(1) & "' "
  sql = sql & " GROUP BY i2.PatientID) "
  sql = sql & "SELECT COALESCE(i.PatientID, i2.PatientID) AS PatientID, "
  sql = sql & " p.PatientName, "
  sql = sql & " COALESCE(i.LabTestCount, 0) + COALESCE(i2.LabTestCount, 0) AS TotalLabTests "
  sql = sql & "FROM InvestigationCount i "
  sql = sql & "FULL OUTER JOIN Investigation2Count i2 ON i.PatientID = i2.PatientID "
  sql = sql & "JOIN Patient p ON p.PatientID = COALESCE(i.PatientID, i2.PatientID); "
  sql = sql & "ORDER BY TotalLabTests DESC"
  
  With rst
      .open qryPro.FltQry(sql), conn, 3, 4
      
      If .RecordCount > 0 Then
          
          response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
          response.write "<tr class='mytr'>"
          response.write "<th class='myth'>No.</th>"
          response.write "<th class='myth'>Patient ID</th>"
          response.write "<th class='myth'>Patient Name</th>"
          response.write "<th class='myth'>Total Lab Tests</th>"
          response.write "</tr class='mytr'>"

          Do While Not .EOF
              response.write "<tr class='mytr'>"
              response.write "<td class='mytd'>" & count & "</td>"
              response.write "<td class='mytd'>" & .fields("PatientID") & "</td>"
              response.write "<td class='mytd'>" & .fields("PatientName") & "</td>"
              response.write "<td class='mytd'>" & .fields("TotalLabTests") & "</td>"
              response.write "</tr class='mytr'>"

              .MoveNext
              count = count + 1
          Loop

          response.write "</table>"
      Else
          response.write "<h1>No records found</h1>"
      End If
      
      .Close
  End With
  
  Set rst = Nothing
End Sub


Sub tableStyles()
    response.write "<style>"
        response.write ".mytable {"
        response.write "    width: 65vw;"
        response.write "    border-collapse: collapse;"
        response.write "    margin: 20px 0;"
        response.write "    font-size: 16px;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "}"
        response.write ".mytable, .myth, .mytd {"
        response.write "    border: 1px solid #dddddd;"
        response.write "}"
        response.write ".myth, .mytd {"
        response.write "    padding: 12px;"
        response.write "    text-align: left;"
        response.write "}"
        response.write ".myth {"
        response.write "    background-color: #f2f2f2;"
        response.write "    color: #333;"
        response.write "    font-weight: bold;"
        response.write "}"
        response.write ".mytr:nth-child(even) {"
        response.write "    background-color: #f9f9f9;"
        response.write "}"
        response.write ".mytr:hover {"
        response.write "    background-color: #f1f1f1;"
        response.write "}"
        response.write ".myth {"
        response.write "    text-transform: uppercase;"
        response.write "}"
        response.write "h1 {"
        response.write "    font-size: 18px;"
        response.write "    color: #555;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "    margin: 20px 0;"
        response.write "}"
response.write "</style>"

End Sub



'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
