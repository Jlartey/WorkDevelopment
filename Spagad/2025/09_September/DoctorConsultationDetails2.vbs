'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim systemUserID, periodStart, periodEnd
systemUserID = Request.QueryString("SystemUserID")
periodStart = Request.QueryString("periodStart")
periodEnd = Request.QueryString("periodEnd")

tableStyles
DoctorConsultationDetails


Sub DoctorConsultationDetails()
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT distinct VisitationID, p.PatientName FROM "
    sql = sql & "EMRRequestItems emr "
    sql = sql & "Join "
    sql = sql & "Patient p "
    sql = sql & "ON p.PatientID = emr.PatientID "
    sql = sql & "WHERE EMRDataID IN ('TH060', 'IM051') "
    sql = sql & "AND EMRDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "AND emr.SystemUserID = '" & systemUserID & "'"
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
        If .recordCount > 0 Then
            response.write "<h1 style='text-align: center; color: #000;'>CONSULTATIONS DONE BY " & GetComboName("Staff", GetComboNameFld("SystemUser", systemUserID, "StaffID")) & "</h1>"
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>VisitationID</th>"
            response.write "<th class='myth'>Patient</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                visitationID = .fields("VisitationID")
                response.write "<tr class='mytr' onclick='redirectToVisitation(""" & visitationID & """)' style='cursor: pointer;'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("VisitationID") & "</td>"
                response.write "<td class='mytd'>" & .fields("PatientName") & "</td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop

            response.write "</table>"
             response.write "<script>"
            response.write "    function redirectToVisitation(visitationID) {"
            response.write "        const baseUrl = 'http://192.168.5.11/rmchms01/wpgPrtPrintLayoutAll.asp';"
            response.write "        const params = new URLSearchParams({"
            response.write "            PrintLayoutName: 'VisitationRCP',"
            response.write "            PositionForTableName: 'Visitation',"
            response.write "            VisitationID: visitationID,"
            response.write "            WorkFlowNav: 'POP'"
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


