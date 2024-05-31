Sub printPatientData()
    Dim patientName, patientId, phoneNumber, phoneNumbers

    ' Construct SQL query
    Dim sql
    sql = "SELECT Patient.PatientName, Patient.PatientID, Patient.BusinessPhone, Visitation.VisitDate FROM "
    sql = sql & " Patient LEFT OUTER JOIN Visitation ON Visitation.PatientId = Patient.PatientID"
    sql = sql & " WHERE Visitation.VisitDate BETWEEN '2024-03-01' AND '2024-03-31'"

    ' Create ADO Recordset
    Dim rst
    Set rst = Server.CreateObject("ADODB.Recordset")

    On Error GoTo ErrorHandler

    ' Open Recordset
    rst.Open sql, conn, 3, 4

    If Not rst.EOF Then
        ' Output table header
        Response.Write "<table cellpadding='1' border='1' width='100%' cellspacing='0'>"
        Response.Write "<tr><th>Name</th><th>Patient ID</th><th>Phone Number 1</th><th>Phone Number 2</th><th>Phone Number 3</th></tr>"

        ' Loop through the recordset
        Do While Not rst.EOF
            ' Retrieve data from the recordset
            patientName = rst("PatientName")
            patientId = rst("PatientID")
            phoneNumber = rst("BusinessPhone")

            ' Check if phoneNumber is not null
            If Not IsNull(phoneNumber) Then
                ' Split phone numbers
                phoneNumbers = Split(phoneNumber, "/")
                
                ' Output table row
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
            Else
                ' Output table row with empty cells for phone numbers
                Response.Write "<tr>"
                Response.Write "<td align='center'>" & patientName & "</td>"
                Response.Write "<td align='center'>" & patientId & "</td>"
                Response.Write "<td align='center'></td>"
                Response.Write "<td align='center'></td>"
                Response.Write "<td align='center'></td>"
                Response.Write "</tr>"
            End If

            ' Move to the next record
            rst.MoveNext
        Loop

        ' Close the table
        Response.Write "</table>"
    Else
        ' No records found
        Response.Write "No records found"
    End If

    ' Close the recordset
    rst.Close
    Set rst = Nothing

Exit Sub

ErrorHandler:
    ' Handle errors
    Response.Write "An error occurred: " & Err.Description
End Sub
