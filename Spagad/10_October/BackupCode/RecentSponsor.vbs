'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

displaySponsors
tableStyles
Sub displaySponsors()
    Dim rst, sql, periodStart, periodEnd, datePeriod, rowNum
    datePeriod = Trim(request.QueryString("PrintFilter"))
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If
    
    Set rst = CreateObject("ADODB.RecordSet")
    
    sql = "SELECT Sponsor.SponsorName, CONVERT(VARCHAR(20), MAX(InsuredPatient.EntryDate), 103) AS EntryDate "
    sql = sql & "From Sponsor "
    sql = sql & "JOIN InsuredPatient ON Sponsor.SponsorID = InsuredPatient.SponsorID "
    sql = sql & "JOIN SponsorType ON SponsorType.SponsorTypeID = InsuredPatient.SponsorTypeID "
    sql = sql & "WHERE InsuredPatient.SponsorTypeID = 'S004' "
    sql = sql & "AND InsuredPatient.EntryDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "GROUP BY Sponsor.SponsorName "
    sql = sql & "ORDER BY MAX(InsuredPatient.EntryDate) DESC"
    
    'response.write sql
    
    With rst
        .open sql, conn, 3, 4
         
        If .RecordCount > 0 Then
           rowNum = 1
            .MoveFirst
            
            response.write "<table class='mytable 'width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr class='mytr'>"
                response.write "<th class = 'myth'> No. </th>"
                response.write "<th class = 'myth'> Sponsor </th>"
                response.write "<th class = 'myth'> Entry Date </th>"
            response.write "</tr>"
            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class = 'mytd'>" & rowNum & "</td>"
                response.write "<td class = 'mytd'>" & .fields("SponsorName") & "</td>"
                response.write "<td class = 'mytd'>" & .fields("EntryDate") & "</td>"
                response.write "</tr>"
                
                rowNum = rowNum + 1
                .MoveNext
            Loop
            response.write "</table>"
        Else
            response.write "No records found"
        End If
        .Close
    End With
    
End Sub

Sub tableStyles()
    response.write "<style>"
        response.write ".mytable {"
        response.write "    width: 50vw;"
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
        response.write "    background-color: #e0e0e0;"
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
'
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>

