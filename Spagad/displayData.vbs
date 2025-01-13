'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Response.Write "<body>"
Response.Write "   <label for='From'>From</label>"
Response.Write "   <input type='date' name='From' id='From' />"

Response.Write "   <label for='To'>To</label>"
Response.Write "   <input type='date' name='To' id='To' />"

Response.Write "   <button id='process'>Process</button>"

Response.Write " <script language='JavaScript'>"
Response.Write "   const fromDate = document.getElementById('From');"
Response.Write "   const toDate = document.getElementById('To');"
Response.Write "   document.getElementById('process').onclick = function() {"
'displayData fromDate, toDate
Response.Write "   };"
Response.Write "</script>"
Response.Write "</body>"


displayData '2018-05-28', '2024-06-19'
Sub displayData(fromDate, toDate)
    Dim sql, numberOfPrescriptions, prescribedDrug

    ' Construct the SQL query with parameters
    sql = "WITH DispensedCTE AS (" & _
          "SELECT Drug.Drugname AS DispensedDrug, DrugSaleItems.DispenseDate, DrugsaleItems.VisitationID " & _
          "FROM DrugsaleItems " & _
          "INNER JOIN Drug ON Drug.DrugID = DrugsaleItems.DrugID " & _
          "WHERE DispenseDate BETWEEN '" & fromDate & "' AND '" & toDate & "'), " & _
          "PrescriptionCTE AS (" & _
          "SELECT Drug.Drugname AS PrescribedDrug, Prescription.PrescriptionDate, Prescription.VisitationID " & _
          "FROM Drug " & _
          "INNER JOIN Prescription ON Drug.DrugID = Prescription.DrugID " & _
          "LEFT JOIN DrugsaleItems ON DrugsaleItems.VisitationID = Prescription.VisitationID " & _
          "WHERE PrescriptionDate BETWEEN '" & fromDate & "' AND '" & toDate & "'), " & _
          "NOTDISPENSEDCTE AS (" & _
          "SELECT PrescriptionCTE.PrescribedDrug, DispensedCTE.DispensedDrug, " & _
          "PrescriptionCTE.PrescriptionDate, DispensedCTE.DispenseDate, " & _
          "PrescriptionCTE.VisitationID AS PresVstID, DispensedCTE.VisitationID AS DisVstID " & _
          "FROM PrescriptionCTE " & _
          "LEFT JOIN DispensedCTE ON PrescriptionCTE.VisitationID = DispensedCTE.VisitationID) " & _
          "SELECT COUNT(*) AS NumberOfPrescriptions, PrescribedDrug " & _
          "FROM NOTDISPENSEDCTE " & _
          "WHERE DispensedDrug IS NULL " & _
          "GROUP BY PrescribedDrug " & _
          "ORDER BY NumberOfPrescriptions DESC, PrescribedDrug;"

    ' Output the SQL query (for debugging purposes)
    Response.Write "<pre>" & sql & "</pre>"

    ' Execute the SQL query and display results
    Dim conn, rst
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.open "your_connection_string_here" ' Replace with your actual connection string

    Set rst = Server.CreateObject("ADODB.Recordset")
    rst.open sql, conn, 3, 4

    If Not rst.EOF Then
        Response.Write "<table border='1'>"
        Response.Write "<tr><th>Number of Prescriptions</th><th>Prescribed Drug</th></tr>"
        Do While Not rst.EOF
            numberOfPrescriptions = rst("NumberOfPrescriptions").value
            prescribedDrug = rst("PrescribedDrug").value
            
            Response.Write "<tr>"
            Response.Write "<td>" & numberOfPrescriptions & "</td>"
            Response.Write "<td>" & prescribedDrug & "</td>"
            Response.Write "</tr>"
            
            rst.MoveNext
        Loop
        Response.Write "</table>"
    Else
        Response.Write "No records found"
    End If

    ' Clean up
    rst.Close
    Set rst = Nothing
    conn.Close
    Set conn = Nothing
End Sub






   




'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
'
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
