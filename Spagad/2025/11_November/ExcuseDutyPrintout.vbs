'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim datePeriod, periodStart, periodEnd, dateArr
datePeriod = Trim(Request.QueryString("PrintFilter"))
If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    periodStart = dateArr(0)
    periodEnd = dateArr(1)
End If

tableStyles
ExcuseDutyPrintOut

Sub ExcuseDutyPrintOut()
    Dim count, sql, rst, emrRequestID
    count = 1
    Set rst = CreateObject("ADODB.Recordset")
    sql = "SELECT ri.PatientID, ri.VisitationID, ri.EMRRequestID, ri.SystemUserID, "
    sql = sql & "CASE WHEN er.Column2 IS NULL THEN NULL "
    sql = sql & "ELSE CONVERT(varchar(12), CONVERT(datetime, CAST(er.Column2 AS nvarchar(50))), 106) END as DateBegan, "
    sql = sql & "CASE WHEN er.Column5 IS NULL THEN NULL "
    sql = sql & "ELSE CONVERT(varchar(12), CONVERT(datetime, CAST(er.Column5 AS nvarchar(50))), 106) END as DateEnded, "
    sql = sql & "CASE WHEN er.Column2 IS NULL OR er.Column5 IS NULL THEN 0 "
    sql = sql & "ELSE DATEDIFF(day, CONVERT(datetime, CAST(er.Column2 AS nvarchar(50))), CONVERT(datetime, CAST(er.Column5 AS nvarchar(50)))) END as NoOfDays "
    sql = sql & "From EMRRequestItems ri "
    sql = sql & "LEFT JOIN emrresults er ON ri.EMRRequestID = er.emrrequestid "
    sql = sql & "AND er.emrdataid = 'IM081' AND er.emrcomponentid = 'IM081.2' "
    sql = sql & "WHERE ri.EMRDataID = 'IM081' "
    sql = sql & "AND ri.EMRDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    
    With rst
        .open sql, conn, 3, 4
       
        If .recordCount > 0 Then
           
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>VisitationID</th>"
            response.write "<th class='myth'>Patient</th>"
            response.write "<th class='myth'>Date Began</th>"
            response.write "<th class='myth'>Date Ended</th>"
            response.write "<th class='myth'>No. Of Days</th>"
            response.write "<th class='myth'>Doctor</th>"
            response.write "</tr class='mytr'>"
            Do While Not .EOF
                emrRequestID = .fields("EMRRequestID")
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("VisitationID") & "</td>"
                response.write "<td class='mytd'>" & GetComboName("Patient", .fields("PatientID")) & "</td>"
                response.write "<td class='mytd'>" & .fields("DateBegan") & "</td>"
                response.write "<td class='mytd'>" & .fields("DateEnded") & "</td>"
                response.write "<td class='mytd'>" & .fields("NoOfDays") & " days</td>"
                response.write "<td class='mytd'>" & GetComboName("Staff", GetComboNameFld("SystemUser", .fields("SystemUserID"), "StaffID")) & "</td>"
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
    response.write " width: 75vw;"
    response.write " border-collapse: collapse;"
    response.write " margin: 20px 0;"
    response.write " font-size: 16px;"
    response.write " font-family: Arial, sans-serif;"
    response.write "}"
    response.write ".mytable, .myth, .mytd {"
    response.write " border: 1px solid #dddddd;"
    response.write "}"
    response.write ".myth, .mytd {"
    response.write " padding: 12px;"
    response.write " text-align: left;"
    response.write "}"
    response.write ".myth {"
    response.write " background-color: #f2f2f2;"
    response.write " color: #333;"
    response.write " font-weight: bold;"
    response.write "}"
    response.write ".mytr:nth-child(even) {"
    response.write " background-color: #f9f9f9;"
    response.write "}"
    response.write ".mytr:hover {"
    response.write " background-color: #f1f1f1;"
    response.write "}"
    response.write ".myth {"
    response.write " text-transform: uppercase;"
    response.write "}"
    response.write "h1 {"
    response.write " font-size: 18px;"
    response.write " color: #555;"
    response.write " font-family: Arial, sans-serif;"
    response.write " margin: 20px 0;"
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

