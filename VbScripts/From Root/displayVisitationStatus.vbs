'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
displayVisitationStatus

Sub displayVisitationStatus()
  Dim sql, visitID, visitDate

  sql = "SELECT TOP 100 VisitationID, VisitDate FROM Visitation"

  Set rst = server.CreateObject("ADODB.Recordset")
  With rst
    .Open sql, conn, 3, 4

    If NOT .EOF Then
    response.write "<table cellpadding='1' border='1' width='100%' cellspacing='0'>"
    response.write "<tr><th>Visitation Id</th><th>Visit Date</th><th>Visit Waver Period</th><th>Remarks</th></tr>"
    
    .MoveFirst
    Do While Not .EOF
      visitID = .fields("VisitationID")
      visitDate = .fields("VisitDate")

      response.write "<tr>"
        response.write "<td align='center'>" & visitID & "</td>"
        response.write "<td align='center'>" & visitDate & "</td>"
        response.write "<td align='center'>" & calculateVisitPeriod(visitDate) & "</td>"
        response.write "<td align='center'>" & visitRemarks(visitDate) & "</td>"

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

Function calculateVisitPeriod(visitDate) 
  Dim todayDate, difference
  
  todayDate = Date
  difference = DateDiff("d", todayDate, visitDate)
  calculateVisitPeriod = difference
End Function

Function visitRemarks(visitDate)
  Dim remarks
  Dim visitPeriod

  ' Calculate visit period
  visitPeriod = calculateVisitPeriod(visitDate)

  ' Determine remarks based on visit period
  If visitPeriod >= 0 And visitPeriod <= 4 Then
    remarks = "Free Consultation!"
  ElseIf visitPeriod >=5 And visitPeriod <= 7 Then
    remarks = "You pay half of the consultation!"
  Else 
    remarks = "You pay the full price!"
  End If

  visitRemarks = remarks

End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>