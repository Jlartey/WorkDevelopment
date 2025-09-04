'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim periodStart, periodEnd, patientID

periodStart = Trim(Request.QueryString("periodStart"))
periodEnd = Trim(Request.QueryString("periodEnd"))
patientID = Trim(Request.QueryString("patientID"))

tableStyles
PharmacistNotesDetails

Sub PharmacistNotesDetails()
    Dim count, sql, rst, emrRequestID
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT EMRRequestID "
    sql = sql & "From EMRRequestItems WHERE EMRDataID = 'E000052' "
    sql = sql & " AND EMRDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    sql = sql & " AND PatientID = '" & patientID & "'"
    sql = sql & "ORDER BY EMRRequestID"
    
    response.write sql
    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth' style='min-width: 150px; text-align: center'>Time</th>"
            response.write "<th class='myth' style='text-align: center'>Notes</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                emrRequestID = .fields("EMRRequestID")
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & getEMRResult(emrRequestID, "E000052", "E00005201", "Column2") & "</td>"
                response.write "<td class='mytd'>" & getEMRResult(emrRequestID, "E000052", "E00005202", "Column2") & "</td>"
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

Sub tableStyles()
    response.write "<style>"
        response.write ".mytable {"
        response.write "    width: 75vw;"
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

