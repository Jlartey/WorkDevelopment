'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
'Almost works Code
printPatientData

Sub printPatientData()
  Dim patientName, patientId, phoneNumber, phoneNumbers

  ' sql = "SELECT PatientName, PatientID, BusinessPhone FROM Patient"
  ' sql = "SELECT PatientName, PatientID, BusinessPhone FROM Patient WHERE Patient.FirstVisitDate BETWEEN '2024-03-01' AND '2024-03-31'"

  sql = "SELECT Patient.PatientName, Patient.PatientID, Patient.BusinessPhone, Visitation.VisitDate FROM "
  sql = sql & " Patient  LEFT OUTER JOIN Visitation ON Visitation.PatientId = Patient.PatientID "
  sql = sql & " WHERE Visitation.VisitDate BETWEEN '2024-03-01' AND '2024-03-31' "
  response.write sql
  Set rst = server.CreateObject("ADODB.Recordset")
  With rst
    .Open sql, conn, 3, 4

    If Not .EOF Then
      response.write "<table cellpadding='1' border='1' width='100%' cellspacing='0'>"
      response.write "<tr><th>Name</th><th>Patient ID</th><th>Phone Number 1</th><th>Phone Number 2</th><th>Phone Number 3</th></tr>"

      .MoveFirst
      Do While Not .EOF
        patientName = .fields("PatientName")
        patientId = .fields("PatientID")
        phoneNumber = .fields("BusinessPhone")
        
        response.write "<tr>"
          response.write "<td align='center'>" & patientName & "</td>"
          response.write "<td align='center'>" & patientId & "</td>"
          
          If InStr(phoneNumber, "/") Then
          
            phoneNumbers = Split(phoneNumber, "/")
            response.write "<td align='center'>" & phoneNumbers(0) & "</td>"
            If UBound(phoneNumbers) >= 1 Then
                response.write "<td align='center'>" & phoneNumbers(1) & "</td>"
            Else
                response.write "<td align='center'></td>"
            End If
            If UBound(phoneNumbers) >= 2 Then
                response.write "<td align='center'>" & phoneNumbers(2) & "</td>"
            Else
                response.write "<td align='center'></td>"
            End If
        Else
        response.write "<td>" & phoneNumber & "</td>"
        response.write "<td></td>"
        response.write "<td></td>"
        End If
          
          
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

