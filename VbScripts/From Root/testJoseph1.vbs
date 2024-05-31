'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
'Almost works Code
printPatientData

Sub printPatientData()
  Dim patientName, patientId, phoneNumber, phoneNumbers

  ' sql = "SELECT PatientName, PatientID, BusinessPhone FROM Patient"
  ' sql = "SELECT PatientName, PatientID, BusinessPhone FROM Patient WHERE Patient.FirstVisitDate BETWEEN '2024-03-01' AND '2024-03-31'"

  sql = "SELECT  Patient.PatientName, Patient.PatientID, Patient.BusinessPhone, Visitation.VisitDate FROM "
  sql = sql & " Patient  LEFT OUTER JOIN Visitation ON Visitation.PatientId = Patient.PatientID"
  sql = sql & " WHERE Visitation.VisitDate BETWEEN '2024-03-01' AND '2024-03-31'"
  
  Set rst = server.CreateObject("ADODB.Recordset")
  With rst
    .Open sql, conn, 3, 4

    If Not .EOF Then
      Response.Write "<table cellpadding='1' border='1' width='100%' cellspacing='0'>"
      Response.Write "<tr><th>Name</th><th>Patient ID</th><th>Phone Number 1</th><th>Phone Number 2</th><th>Phone Number 3</th></tr>"

      .MoveFirst
      Do While Not .EOF
        patientName = .fields("PatientName")
        patientId = .fields("PatientID")
        phoneNumber = .fields("BusinessPhone")

        phoneNumbers = Split(phoneNumber, "/")
        
        Response.Write "<tr>"
          Response.Write "<td align='center'>" & patientName & "</td>"
          Response.Write "<td align='center'>" & patientId & "</td>"
          Response.Write "<td align='center'>" & phoneNumbers(0) & "</td>"
          If UBound(phoneNumbers) >= 1 Then
                Response.Write "<td align='center'>" & phoneNumbers(1) & "</td>"
            Else
                Response.Write "<td align='center'></td>"
            End If
            If UBound(phoneNumbers) >= 2 Then
                Response.Write "<td align='center'>" & phoneNumbers(2) & "</td>"
            Else
                Response.Write "<td align='center'></td>"
            End If
        Response.Write "</tr>"
        
        .MoveNext
      Loop
        Response.Write "</table>"
    Else
      Response.Write "No records found"
    End If
    .Close
  End With
End Sub
Sub printPatientData()
  Dim patientName, patientId, phoneNumber, phoneNumbers

  ' sql = "SELECT PatientName, PatientID, BusinessPhone FROM Patient"
  ' sql = "SELECT PatientName, PatientID, BusinessPhone FROM Patient WHERE Patient.FirstVisitDate BETWEEN '2024-03-01' AND '2024-03-31'"

  sql = "SELECT  Patient.PatientName, Patient.PatientID, Patient.BusinessPhone, Visitation.VisitDate FROM "
  sql = sql & " Patient  LEFT OUTER JOIN Visitation ON Visitation.PatientId = Patient.PatientID"
  sql = sql & " WHERE Visitation.VisitDate BETWEEN '2024-03-01' AND '2024-03-31'"
  
  Set rst = server.CreateObject("ADODB.Recordset")
  With rst
    .Open sql, conn, 3, 4

    If Not .EOF Then
      Response.Write "<table cellpadding='1' border='1' width='100%' cellspacing='0'>"
      Response.Write "<tr><th>Name</th><th>Patient ID</th><th>Phone Number 1</th><th>Phone Number 2</th><th>Phone Number 3</th></tr>"

      .MoveFirst
      Do While Not .EOF
        patientName = .fields("PatientName")
        patientId = .fields("PatientID")
        phoneNumber = .fields("BusinessPhone")

        phoneNumbers = Split(phoneNumber, "/")
        
        ' Output table row
        Response.Write "<tr>"
          Response.Write "<td align='center'>" & patientName & "</td>"
          Response.Write "<td align='center'>" & patientId & "</td>"
          Response.Write "<td align='center'>" & phoneNumber & "</td>"
            If UBound(phoneNumbers) >= 1 Then
                 Response.Write "<td align='center'>" & phoneNumbers(1) & "</td>"
             Else
                 Response.Write "<td align='center'></td>"
             End If
             If UBound(phoneNumbers) >= 2 Then
                 Response.Write "<td align='center'>" & phoneNumbers(2) & "</td>"
             Else
                 Response.Write "<td align='center'></td>"
             End If
        Response.Write "</tr>"
        
        .MoveNext
      Loop
        Response.Write "</table>"
    Else
      Response.Write "No records found"
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
