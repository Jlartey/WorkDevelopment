
'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim datePeriod, periodStart, periodEnd, dateArr

datePeriod = Trim(Request.queryString("PrintFilter"))

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
VisitCount
AppointmentCount
AppointmentsByStaff

Sub VisitCount()
    Dim count, sql, rst, totalVisits
    count = 1
    totalVisits = 0

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT SpecialistGroupID, COUNT(*)Visits FROM Visitation "
    sql = sql & "WHERE VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "GROUP BY SpecialistGroupID"
    
'    response.write sql

    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            
            response.write "<h3>SHOWING VISITS BETWEEN " & periodStart & " AND " & periodEnd & " </h3>"
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Department</th>"
            response.write "<th class='myth'>Visits</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                totalVisits = totalVisits + .fields("Visits")
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & GetComboName("SpecialistGroup", .fields("SpecialistGroupID")) & "</td>"
                response.write "<td class='mytd'>" & .fields("Visits") & "</td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop
            response.write "<tr>"
                response.write "<td colspan='2' class='mytd' style='text-align:right; font-weight: bold'> TOTAL VISITS</td>"
                response.write "<td class='mytd'>" & totalVisits & "</td>"
            response.write "</tr>"

            response.write "</table>"
        Else
            response.write "<h1>No records found</h1>"
        End If
        
        .Close
    End With
    
    Set rst = Nothing
End Sub

Sub AppointmentCount()
    Dim count, sql, rst, totalAppointments
    
    count = 1
    totalAppointments = 0

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT AppointmentStatusID, COUNT(*)Appointments FROM Appointment "
    sql = sql & "WHERE AppointDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "GROUP BY AppointmentStatusID"
    

    'response.write sql

    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
        
            response.write "<h3>SHOWING APPOINTMENTS GROUPED BY APPOINTMENT STATUS BETWEEN " & periodStart & " AND " & periodEnd & " </h3>"
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Appointment Status</th>"
            response.write "<th class='myth'>Appointments</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                totalAppointments = totalAppointments + .fields("Appointments")
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & GetComboName("AppointmentStatus", .fields("AppointmentStatusID")) & "</td>"
                response.write "<td class='mytd'>" & .fields("Appointments") & "</td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop
            
            response.write "<tr>"
                response.write "<td colspan='2' class='mytd' style='text-align:right; font-weight: bold'> TOTAL APPOINTMENTS</td>"
                response.write "<td class='mytd'>" & totalAppointments & "</td>"
            response.write "</tr>"
            
            response.write "</table>"
        Else
            response.write "<h1>No records found</h1>"
        End If
        
        .Close
    End With
    
    Set rst = Nothing
End Sub

Sub AppointmentsByStaff()
    Dim count, sql, rst, totalAppointments
    
    count = 1
    totalAppointments = 0

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT SystemUserID, COUNT(*)Appointments FROM Appointment "
    sql = sql & "WHERE AppointDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "GROUP BY SystemUserID"
    

'    response.write sql

    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            response.write "<h3>SHOWING APPOINTMENTS BY STAFF BETWEEN " & periodStart & " AND " & periodEnd & " </h3>"
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Staff</th>"
            response.write "<th class='myth'>Appointments</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                totalAppointments = totalAppointments + .fields("Appointments")
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & GetComboName("Staff", GetComboNameFld("SystemUser", .fields("SystemUserID"), "StaffID")) & "</td>"
                response.write "<td class='mytd'>" & .fields("Appointments") & "</td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop
            
            response.write "<tr>"
                response.write "<td colspan='2' class='mytd' style='text-align:right; font-weight: bold'> TOTAL APPOINTMENTS</td>"
                response.write "<td class='mytd'>" & totalAppointments & "</td>"
            response.write "</tr>"
            
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
        response.write "h3 {"
        response.write "    font-size: 22px;"
        response.write "    color: #000;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "    margin: 15px 0;"
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

