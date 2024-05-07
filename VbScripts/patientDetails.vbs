'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
' This code was given to me by chatgpt
displayTopTenPatients

Sub displayTopTenPatients()
    Dim sql, patientID
    Dim conn, rst
    Dim strConn, connectionString
    
    ' Define connection string
    strConn = "your_connection_string_here"
    
    ' Create connection object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Open connection
    conn.Open strConn
    
    ' Check if connection is open
    If conn.State = 1 Then
        ' SQL query to select top 10 patients based on visitation date
        sql = "SELECT TOP 10 Patient.PatientID, Patient.PatientName " & _
              "FROM Patient INNER JOIN Visitation ON Patient.PatientID = Visitation.PatientID " & _
              "WHERE Visitation.VisitDate = '2024-05-02'"
        
        ' Create recordset object
        Set rst = CreateObject("ADODB.Recordset")
        
        ' Open recordset
        rst.Open sql, conn
        
        ' Check if recordset is not empty
        If Not rst.EOF Then
            ' Begin HTML output
            Response.Write "<table>"
            Response.Write "<tr><th>Patient ID</th><th>Patient Name</th></tr>"
            
            ' Loop through recordset and display patient details
            Do While Not rst.EOF
                Response.Write "<tr>"
                Response.Write "<td>" & rst("PatientID") & "</td>"
                Response.Write "<td>" & rst("PatientName") & "</td>"
                Response.Write "</tr>"
                rst.MoveNext
            Loop
            
            ' End HTML output
            Response.Write "</table>"
        Else
            ' No records found message
            Response.Write "No records found!"
        End If
        
        ' Close recordset
        rst.Close
    Else
        ' Connection failed message
        Response.Write "Failed to establish database connection!"
    End If
    
    ' Close connection
    conn.Close
    
End Sub
'<<--END_CODE_SEGMENT_PRINTHEADER-->>

Function
