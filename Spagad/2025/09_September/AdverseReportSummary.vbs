'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim datePeriod, periodStart, periodEnd, dateArr

datePeriod = Trim(Request.QueryString("PrintFilter"))
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If
    
    
tableStyles
AdverseReportSummary

Sub AdverseReportSummary()
    Dim count, sql, rst, emrRequestID
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT PatientID, VisitationID, EMRRequestID "
    sql = sql & "From EMRRequestItems WHERE EMRDataID = 'ADV001' "
    sql = sql & "AND EMRDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    
    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Patient</th>"
            response.write "<th class='myth'>VisitationID</th>"
            response.write "<th class='myth'>Reporter Name</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                emrRequestID = .fields("EMRRequestID")
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & GetComboName("Patient", .fields("PatientID")) & "</td>"
                response.write "<td class='mytd'>" & .fields("VisitationID") & "</td>"
                response.write "<td class='mytd'>" & getEMRResult(emrRequestID, "ADV001", "ADV001.23", "Column2") & "</td>"
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

Function getEMRResult(emrRequestID, emrDataID, CompID, column)
    Dim sql, rst, emrValue
    Set rst = Server.CreateObject("ADODB.Recordset")
    emrValue = ""
    sql = "SELECT * FROM emrresults WHERE emrrequestid ='" & emrRequestID & "'"
    sql = sql & " AND emrdataid ='" & emrDataID & "' AND emrcomponentid='" & CompID & "'"
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            If Not IsNull(.fields(column)) Then
                emrValue = Trim(.fields(column))
            End If
            If IsNull(.fields(column)) Then
                emrValue = "Null"
            End If
        End If
        .Close
    End With
    getEMRResult = emrValue
    Set rst = Nothing
End Function


'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>


