'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim datePeriod, periodStart, periodEnd, dateArr

datePeriod = Trim(Request.QueryString("PrintFilter"))
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If
    
tableStyles
DoctorsConsultations

Sub DoctorsConsultations()
    Dim count, sql, rst, staffName
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT "
    sql = sql & "Staff.StaffName Doctor, "
    sql = sql & "COUNT(DISTINCT Diagnosis.VisitationId) Consults, "
    sql = sql & "VisitTotals.VisitAmount "
    sql = sql & "From diagnosis "
    sql = sql & "Join SystemUser "
    sql = sql & "ON SystemUser.SystemUserID = Diagnosis.SystemUserID "
    sql = sql & "Join Staff "
    sql = sql & "ON Staff.StaffID = SystemUser.StaffID "
    sql = sql & "JOIN ( "
    sql = sql & "SELECT "
    sql = sql & "v.VisitationID, "
    sql = sql & "MIN(v.visitCost) As visitCost "
    sql = sql & "FROM Visitation v "
    sql = sql & "WHERE EXISTS ( "
    sql = sql & "SELECT 1 "
    sql = sql & "FROM Diagnosis d "
    sql = sql & "Where d.visitationID = v.visitationID "
    sql = sql & ") "
    sql = sql & "GROUP BY v.VisitationID "
    sql = sql & ") UniqueVisits "
    sql = sql & "ON Diagnosis.VisitationId = UniqueVisits.VisitationID "
    sql = sql & "JOIN ( "
    sql = sql & "SELECT StaffName, SUM(VisitCost) AS VisitAmount "
    sql = sql & "FROM ( "
    sql = sql & "SELECT DISTINCT s.StaffName, v.VisitationID, v.VisitCost "
    sql = sql & "FROM Diagnosis d "
    sql = sql & "JOIN SystemUser su ON su.SystemUserID = d.SystemUserID "
    sql = sql & "JOIN Staff s ON s.StaffID = su.StaffID "
    sql = sql & "JOIN Visitation v ON d.VisitationId = v.VisitationID "
    sql = sql & ") DistinctVisits "
    sql = sql & "GROUP BY StaffName "
    sql = sql & ") VisitTotals "
    sql = sql & "ON Staff.StaffName = VisitTotals.StaffName "
    sql = sql & "WHERE ConsultReviewDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "GROUP BY Staff.StaffName, VisitTotals.VisitAmount "
    sql = sql & "ORDER BY Consults DESC"
    
    With rst
        .open sql, conn, 3, 4
        
        If .recordCount > 0 Then
            
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Doctor</th>"
            response.write "<th class='myth'>Consultations</th>"
            response.write "<th class='myth'>Visit Cost Total</th>"
            response.write "<th class='myth'>View More</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
            staffName = .fields("Doctor")
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("Doctor") & "</td>"
                response.write "<td class='mytd'>" & .fields("Consults") & "</td>"
                response.write "<td class='mytd'>" & FormatNumber(.fields("VisitAmount")) & "</td>"
                response.write "<td class='mytd' style='cursor: pointer; color: blue' onclick='redirectToVisitation(""" & staffName & """)'><u>View More</u></td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop
            response.write "</table>"
            response.write "<script>"
            response.write "    function redirectToVisitation(staffName) {"
            response.write "        const baseUrl = 'http://192.168.5.11/rmchms01/wpgPrtPrintLayoutAll.asp';"
            response.write "        const params = new URLSearchParams({"
            response.write "            PrintLayoutName: 'DoctorConsultationDetails',"
            response.write "            PositionForTableName: 'WorkingDay',"
            response.write "            WorkingDayID: '',"
            response.write "            staffName: staffName"
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
