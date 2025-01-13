'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim periodStart, periodEnd
periodStart = Trim(Request.querystring("printfilter"))
periodEnd = Trim(Request.querystring("printfilter1"))

tableStyles
MainReport

Sub MainReport()
    dim rst, sql
    set rst = CreateObject("ADODB.RecordSet")
    sql = "SELECT visitationid FROM Visitation WHERE visitdate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"

    with rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Visit Date</th>"
             response.write "<th class='myth'>Payment Receipt Date</th>"
            response.write "<th class='myth'>Begin Consultation Date</th>"
            response.write "<th class='myth'>EMRRequest Date</th>"
            response.write "<th class='myth'>LabRequest Date</th>"
            response.write "<th class='myth'>Time Between Consultations</th>"
            response.write "</tr class='mytr'>"

            visitationID = .fields("visitationID")
            Do While Not .EOF
                dispDeadPatients(visitationID)
                .MoveNext
                count = count + 1
            Loop

            response.write "</table>"
        Else
            response.write "<h1>No records found</h1>"
        End If
        
        .Close
    end with

    Set rst = Nothing
End Sub

Sub dispDeadPatients(visitationID)
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT DISTINCT "
    sql = sql & " v.visitdate AS VisitDate, "
    sql = sql & "     pr.ReceiptDate AS PaymentReceiptDate, "
    sql = sql & "     sl.logdate AS BeginConsultationDate, "
    sql = sql & "     er.EMRDate AS EMRRequestDate, "
    sql = sql & "     lr.requestdate AS LabRequestDate,     "
    sql = sql & "     DATEDIFF(minute, v.visitdate, sl.logdate) AS TotTimeBtwnVstCons "
    sql = sql & " FROM Visitation v "
    sql = sql & " LEFT JOIN PatientReceipt2 pr ON pr.visitationid = v.visitationID AND pr.tableid = 'visitation' "
    sql = sql & " LEFT JOIN EMRRequest er ON er.visitationid = v.visitationID "
    sql = sql & " JOIN EMRResults emr ON emr.EMRRequestID = er.EMRRequestID AND emr.EMRDataID = 'EMR050' "
    sql = sql & " LEFT JOIN LabRequest lr ON lr.visitationid = v.visitationID "
    sql = sql & " LEFT JOIN SystemLog sl ON sl.keyvalue = v.visitationID AND sl.keyprefix = 'BEGIN_CONSULTATION' "
    sql = sql & " WHERE v.visitationID = '" & visitationID & "' "

 
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
        If .RecordCount > 0 Then

            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("VisitDate") & "</td>"
                response.write "<td class='mytd'>" & .fields("PaymentReceiptDate") & "</td>"
                response.write "<td class='mytd'>" & .fields("BeginConsultationDate") & "</td>"
                response.write "<td class='mytd'>" & .fields("EMRRequestDate") & "</td>"
                response.write "<td class='mytd'>" & .fields("LabRequestDate") & "</td>"
                response.write "<td class='mytd'>" & .fields("TotTimeBtwnVstCons") & "</td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop
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
