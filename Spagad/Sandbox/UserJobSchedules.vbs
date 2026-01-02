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
dispUserJobSchedules

Sub dispUserJobSchedules()
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT PerformVar12.Description AS StaffID, "
    sql = sql & "Staff.StaffName, "
    sql = sql & "STRING_AGG(JobSchedule.JobScheduleName, ', ') AS JobSchedules "
    sql = sql & "FROM PerformVar12 "
    sql = sql & "JOIN Staff "
    sql = sql & "ON PerformVar12.Description = Staff.StaffID "
    sql = sql & "JOIN JobSchedule "
    sql = sql & "ON PerformVar12.KeyPrefix = JobSchedule.JobScheduleID "
    sql = sql & "GROUP BY PerformVar12.Description, Staff.StaffName "
    sql = sql & "ORDER BY Staff.StaffName"


    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Staff ID</th>"
            response.write "<th class='myth'>Staff Name</th>"
            response.write "<th class='myth'>Job Schedules</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("StaffID") & "</td>"
                response.write "<td class='mytd'>" & .fields("StaffName") & "</td>"
                response.write "<td class='mytd'>" & .fields("JobSchedules") & "</td>"
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

