'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
'Dim datePeriod
'datePeriod = Trim(Request.querystring("PrintFilter1"))
'
'If datePeriod <> "" Then
'    dateArr = Split(datePeriod, "||")
'    periodStart = dateArr(0)
'    periodEnd = dateArr(1)
'End If

tableStyles
dispRegisteredPatientDetails

Sub dispRegisteredPatientDetails()
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT PatientID, PatientName, Age, GenderName, CountryName "
    sql = sql & "FROM Patient "
    sql = sql & "JOIN Country ON Country.CountryID = Patient.CountryID "
    sql = sql & "JOIN Gender ON Gender.GenderID = Patient.genderID "
    sql = sql & "WHERE 1=1 "
    'sql = sql & "AND firstvisitdate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "AND firstYearID = 'YRS2023' "
    sql = sql & ""
    sql = sql & "ORDER BY CountryName ASC "



    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Patient ID</th>"
             response.write "<th class='myth'>Patient Name</th>"
            response.write "<th class='myth'>Age</th>"
            response.write "<th class='myth'>Gender</th>"
            response.write "<th class='myth'>Country</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("PatientID") & "</td>"
                response.write "<td class='mytd'>" & .fields("PatientName") & "</td>"
                response.write "<td class='mytd'>" & .fields("Age") & "</td>"
                response.write "<td class='mytd'>" & .fields("GenderName") & "</td>"
                response.write "<td class='mytd'>" & .fields("CountryName") & "</td>"
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
