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
VisitCount
AppointmentCount
AppointmentsByStaff
VisitAppointments

Sub VisitCount()
    Dim count, sql, rst, totalVisits
    count = 1
    totalVisits = 0

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT SpecialistGroupID, COUNT(*)Visits FROM Visitation "
    sql = sql & "WHERE VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "AND SpecialistGroupID NOT IN ('CD012', 'CD020') "
    sql = sql & "GROUP BY SpecialistGroupID"
    
'    response.write sql

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
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
                response.write "<td class='mytd' style = 'font-weight: bold'>" & totalVisits & "</td>"
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
    sql = sql & "Join AppointmentCatType "
    sql = sql & "ON Appointment.AppointmentCatTypeID = AppointmentCatType.AppointmentCatTypeID "
    sql = sql & "WHERE EntryDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "AND Appointment.AppointmentCatID NOT IN ('CD012', 'A007') "
    sql = sql & "AND AppointmentCatType.AppointmentCatTypeName NOT LIKE '%Physio%' "
    sql = sql & "AND AppointmentCatType.AppointmentCatTypeName NOT LIKE '%Diet%' "
    sql = sql & "GROUP BY AppointmentStatusID "
    sql = sql & "ORDER BY AppointmentStatusID"
    

    'response.write sql

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
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
                response.write "<td class='mytd' >" & .fields("Appointments") & "</td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop
            
            response.write "<tr>"
                response.write "<td colspan='2' class='mytd' style='text-align:right; font-weight: bold'> TOTAL APPOINTMENTS</td>"
                response.write "<td class='mytd' style = 'font-weight: bold'>" & totalAppointments & "</td>"
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

'    sql = "SELECT SystemUserID, COUNT(*)Appointments FROM Appointment "
'    sql = sql & "WHERE EntryDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
'    sql = sql & "AND AppointmentCatID <> 'CD012' "
'    sql = sql & "GROUP BY SystemUserID"
    
    sql = "SELECT SystemUserID, COUNT(*)Appointments FROM Appointment "
    sql = sql & "Join AppointmentCatType "
    sql = sql & "ON Appointment.AppointmentCatTypeID = AppointmentCatType.AppointmentCatTypeID "
    sql = sql & "WHERE EntryDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "AND Appointment.AppointmentCatID NOT IN ('CD012', 'A007') "
    sql = sql & "AND AppointmentCatType.AppointmentCatTypeName NOT LIKE '%Physio%' "
    sql = sql & "AND AppointmentCatType.AppointmentCatTypeName NOT LIKE '%Diet%' "
    sql = sql & "GROUP BY SystemUserID "
    sql = sql & "ORDER BY SystemUserID"

'    response.write sql

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
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
                response.write "<td class='mytd' style = 'font-weight: bold'>" & totalAppointments & "</td>"
            response.write "</tr>"
            
            response.write "</table>"
        Else
            response.write "<h1>No records found</h1>"
        End If
        
        .Close
    End With
    
    Set rst = Nothing
End Sub

Sub VisitAppointments()
    Dim count, sql, rst
    
    count = 1

    Set rst = CreateObject("ADODB.Recordset")
    
    sql = sql & "WITH RankedData AS ( " & vbCrLf
    sql = sql & "    SELECT " & vbCrLf
    sql = sql & "        A.PatientID, " & vbCrLf
    sql = sql & "        P.PatientName, " & vbCrLf
    sql = sql & "        V.VisitDate, " & vbCrLf
    sql = sql & "        A.AppointDate, " & vbCrLf
    sql = sql & "        ROW_NUMBER() OVER ( " & vbCrLf
    sql = sql & "            PARTITION BY CAST(V.VisitDate AS DATE) " & vbCrLf
    sql = sql & "            ORDER BY CAST(A.AppointDate AS DATE) ASC " & vbCrLf
    sql = sql & "        ) AS rn " & vbCrLf
    sql = sql & "    FROM Appointment A " & vbCrLf
    sql = sql & "    JOIN Patient P ON A.PatientID = P.PatientID " & vbCrLf
    sql = sql & "    JOIN Visitation V ON A.PatientID = V.PatientID " & vbCrLf
    sql = sql & "    WHERE  " & vbCrLf
    sql = sql & "        CAST(A.EntryDate AS DATE) = CAST(V.VisitDate AS DATE) " & vbCrLf
    sql = sql & "        AND V.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' " & vbCrLf
    sql = sql & ") " & vbCrLf
    sql = sql & "SELECT  " & vbCrLf
    sql = sql & "    PatientID, " & vbCrLf
    sql = sql & "    PatientName, " & vbCrLf
    sql = sql & "    CONVERT(VARCHAR(20), VisitDate, 106) AS VisitDate, " & vbCrLf
    sql = sql & "    CONVERT(VARCHAR(20), AppointDate, 106) AS AppointDate " & vbCrLf
    sql = sql & "FROM RankedData " & vbCrLf
    sql = sql & "WHERE rn = 1 " & vbCrLf
    sql = sql & "ORDER BY CAST(VisitDate AS DATE) DESC;"

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
        If .RecordCount > 0 Then
            response.write "<h3>SHOWING VISITS, APPOINTMENTS BY STAFF BETWEEN " & periodStart & " AND " & periodEnd & " </h3>"
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>PatientID</th>"
            response.write "<th class='myth'>PatientName</th>"
            response.write "<th class='myth'>Visit Date</th>"
            response.write "<th class='myth'>Review Date</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("PatientID") & "</td>"
                response.write "<td class='mytd'>" & .fields("PatientName") & "</td>"
                response.write "<td class='mytd'>" & .fields("VisitDate") & "</td>"
                response.write "<td class='mytd'>" & .fields("AppointDate") & "</td>"
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
