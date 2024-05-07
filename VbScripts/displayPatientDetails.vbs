'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
' This was given to me on the 2nd of May by Edwin. 
' I managed to get it done and it worked
displayTopTenPatients

Sub displayTopTenPatients()
    Dim sql, patientID
    'sql = "SELECT TOP 10 Patient.PatientName FROM Patient INNER JOIN Visitation ON Patient.PatientID = Visitation.PatientID where Visitation.VisitDate = '2024-05-02'"
    sql = "SELECT TOP 10 PatientID FROM Visitation WHERE WorkingDayID  = 'DAY20240502'"
    
    Set rst = server.CreateObject("ADODB.RecordSet")
    With rst
      .open sql, conn, 3, 4
        If Not .EOF Then
          response.write "<table>"
          response.write "<tr><th>Patient Name</th></tr>"
          
          .MoveFirst
          Do While Not .EOF
            patientID = .fields("PatientID")
            
            response.write "<tr>"
              response.write "<td>" & GetComboName("patient", patientID) & "</td>"
            response.write "</tr>"
          .MoveNext
          Loop
          response.write "</table>"
        Else
          response.write "No records found!"
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
2
2
