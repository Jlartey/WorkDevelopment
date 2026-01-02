'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim datePeriod, periodStart, periodEnd, dateArr

datePeriod = Trim(Request.QueryString("PrintFilter"))

periodStart = ""
periodEnd = ""

If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    If UBound(dateArr) >= 1 Then
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If
End If
    
tableStyles
DoctorsConsultations periodStart, periodEnd

Sub DoctorsConsultations(periodStart, periodEnd)
    Dim count, sql, rst, systemUserID
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "WITH DoctorConsults AS (SELECT SystemUserID, COUNT(DISTINCT VisitationID)Consults FROM "
    sql = sql & "EMRRequestItems "
    sql = sql & "WHERE EMRDataID IN ('TH060', 'IM051') "
    sql = sql & "AND EMRDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "GROUP BY SystemUserID "
    sql = sql & ") "
    sql = sql & "SELECT Staff.StaffName, dc.SystemUserID, dc.Consults "
    sql = sql & "FROM DoctorConsults dc "
    sql = sql & "Join SystemUser "
    sql = sql & "ON SystemUser.SystemUserID = dc.SystemUserID "
    sql = sql & "Join Staff "
    sql = sql & "ON Staff.StaffID = SystemUser.StaffID "
    sql = sql & "ORDER BY dc.Consults DESC"
    
    
    With rst
        .open sql, conn, 3, 4
        
        If .recordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Doctor</th>"
            response.write "<th class='myth'>Consultations</th>"
            response.write "<th class='myth'>View More</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
            systemUserID = .fields("SystemUserID")
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("StaffName") & "</td>"
                response.write "<td class='mytd'>" & .fields("Consults") & "</td>"
                response.write "<td class='mytd' style='cursor: pointer; color: blue' onclick='redirectToConsultDetails(""" & server.HTMLEncode(systemUserID) & """, """ & server.HTMLEncode(periodStart) & """, """ & server.HTMLEncode(periodEnd) & """)'><u>View More</u></td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop
            response.write "</table>"
            response.write "<script>"
            response.write "    function redirectToConsultDetails(systemUserID, periodStart, periodEnd) {"
            response.write "        const baseUrl = 'http://192.168.5.11/rmchms01/wpgPrtPrintLayoutAll.asp';"
            response.write "        const params = new URLSearchParams({"
            response.write "            PrintLayoutName: 'DoctorConsultationDetailsII',"
            response.write "            PositionForTableName: 'WorkingDay',"
            response.write "            WorkingDayID: '',"
            response.write "            SystemUserID: systemUserID,"
            response.write "            periodStart: periodStart,"
            response.write "            periodEnd: periodEnd"
            response.write "        });"
            response.write "        const newUrl = baseUrl + '?' + params.toString();"
            response.write "        window.open(newUrl, '_blank');"
            response.write "    }"
            response.write "</script>"
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


