'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim periodStart, periodEnd

datePeriod = Trim(Request.QueryString("PrintFilter"))

If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    periodStart = dateArr(0)
    periodEnd = dateArr(1)
End If

tableStyles
dispDuplicteVisits

Sub dispDuplicteVisits()
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT DISTINCT v1.PatientID, convert(VARCHAR(20), v1.VisitDate, 106)VisitDay, v1.VisitationID NHIS, v2.VisitationID OTHER "
    sql = sql & "FROM Visitation v1 "
    sql = sql & "INNER JOIN Visitation v2 ON v1.PatientID = v2.PatientID "
    sql = sql & "AND CAST(v1.VisitDate AS DATE) = CAST(v2.VisitDate AS DATE) "
    sql = sql & "AND v1.SponsorID = 'NHIS' "
    sql = sql & "AND v2.SponsorID != 'NHIS' "
    sql = sql & "AND v1.VisitationID != v2.VisitationID "
    sql = sql & "WHERE v1.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    
    With rst
        .open sql, conn, 3, 4
        
        If .recordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Patient ID</th>"
            response.write "<th class='myth'>Patient</th>"
            response.write "<th class='myth'>Visit Day</th>"
            response.write "<th class='myth'>NHIS VISIT</th>"
            response.write "<th class='myth'>OTHER VISIT</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("PatientID") & "</td>"
                response.write "<td class='mytd'>" & GetComboName("Patient", .fields("PatientID")) & "</td>"
                response.write "<td class='mytd'>" & .fields("VisitDay") & "</td>"
                response.write "<td class='mytd'>" & .fields("NHIS") & "</td>"
                response.write "<td class='mytd'>" & .fields("OTHER") & "</td>"
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
