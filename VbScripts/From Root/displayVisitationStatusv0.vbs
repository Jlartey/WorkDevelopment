'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
displayVisitationStatus

Sub displayVisitationStatus()
  Dim sql, visitID, visitDate

  sql = "SELECT TOP 100 VisitationID, VisitDate FROM Visitation"

  Set rst = server.CreateObject("ADODB.Recordset")
  With rst
    .Open sql, conn, 3, 4

    If Not .EOF Then
    response.write "<table cellpadding='1' border='1' width='100%' cellspacing='0'>"
    response.write "<tr><th>Visitation Id</th><th>Visit Date</th></tr>"
    
    .MoveFirst
    Do While Not .EOF
      visitID = .fields("VisitationID")
      visitDate = .fields("VisitDate")

      response.write "<tr>"
        response.write "<td align='center'>" & visitID & "</td>"
        response.write "<td align='center'>" & visitDate & "</td>"
      response.write "</tr>"
      .MoveNext
    Loop
      response.write "</table>"
    Else
     response.write "No records found"
    End If
    .Close
  End With
End Sub



'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>

