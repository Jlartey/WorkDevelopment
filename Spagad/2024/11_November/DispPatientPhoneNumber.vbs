'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim specialistGroup
specialistGroup = Trim(Request.querystring("PrintFilter"))

tableStyles
dispPatientPhoneNumber

Sub dispPatientPhoneNumber()
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT Patient.PatientName, STRING_AGG(CASE WHEN LEN(ResidencePhone) = 10 THEN ResidencePhone "
    sql = sql & "ELSE NULL  END, ', ' ) AS PhoneNumbers From Patient Join Visitation "
    sql = sql & "ON Patient.PatientID = Visitation.PatientID "
    sql = sql & "WHERE SpecialistGroupID = '" & specialistGroup & "' "
    sql = sql & "GROUP BY Patient.PatientName "
    sql = sql & "HAVING COUNT(CASE WHEN LEN(ResidencePhone) = 10 THEN 1 ELSE NULL END) > 0"
      
    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth' style='text-align: center;'>No.</th>"
            response.write "<th class='myth'>Patient</th>"
            response.write "<th class='myth'>Phone Number(s)</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd' style='text-align: center;'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("PatientName") & "</td>"
                response.write "<td class='mytd' style='width: 40%'>" & .fields("PhoneNumbers") & "</td>"
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
        response.write "    width: 75vw;"
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
