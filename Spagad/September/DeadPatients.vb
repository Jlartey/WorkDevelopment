'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim periodStart, periodEnd
periodStart = Trim(Request.querystring("printfilter"))
periodEnd = Trim(Request.querystring("printfilter1"))

tableStyles
dispDeadPatients

Sub dispDeadPatients()
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT visitation.PatientID, Patient.Age, Patient.GenderID, SpecialistTypeID, SpecialistID "
    sql = sql & "FROM Visitation "
    sql = sql & "JOIN Patient ON Patient.PatientID = Visitation.PatientID "
    sql = sql & "WHERE medicaloutcomeID = 'M002' "
    sql = sql & "AND Visitdate Between '" & periodStart & "' AND '" & periodEnd & "'"


    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Patient</th>"
             response.write "<th class='myth'>Age</th>"
            response.write "<th class='myth'>Gender</th>"
            response.write "<th class='myth'>Specialist Type</th>"
            response.write "<th class='myth'>Specialist</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & GetComboName("Patient", .fields("PatientID")) & "</td>"
                response.write "<td class='mytd'>" & .fields("Age") & "</td>"
                response.write "<td class='mytd'>" & GetComboName("Gender", .fields("GenderID")) & "</td>"
                response.write "<td class='mytd'>" & GetComboName("SpecialistType", .fields("SpecialistTypeID")) & "</td>"
                response.write "<td class='mytd'>" & GetComboName("Specialist", .fields("SpecialistID")) & "</td>"
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
