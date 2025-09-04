'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim treatmentID, periodStart, periodEnd
treatmentID = Request.QueryString("treatmentID")
periodStart = Request.QueryString("periodStart")
periodEnd = Request.QueryString("periodEnd")

tableStyles
TreatmentByDoctor treatmentID, periodStart, periodEnd

Sub TreatmentByDoctor(treatmentID, periodStart, periodEnd)
    Dim count, sql, rst, systemUserID, totalRevenue, patients
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "WITH DoctorTreatAnalysis AS ( "
    sql = sql & "SELECT "
    sql = sql & "SystemUserID, "
    sql = sql & "SUM(UnitCost) AS TotalCost, "
    sql = sql & "COUNT(TreatmentID) AS Count "
    sql = sql & "FROM TreatCharges "
    sql = sql & "WHERE ConsultReviewDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "AND TreatmentID = '" & treatmentID & "' "
    sql = sql & "GROUP BY SystemUserID "
    sql = sql & ") "
    sql = sql & "SELECT "
    sql = sql & "Staff.StaffName, "
    sql = sql & "DoctorTreatAnalysis.SystemUserID, "
    sql = sql & "DoctorTreatAnalysis.TotalCost, "
    sql = sql & "DoctorTreatAnalysis.Count "
    sql = sql & "FROM Staff "
    sql = sql & "JOIN SystemUser ON SystemUser.StaffID = Staff.StaffID "
    sql = sql & "JOIN DoctorTreatAnalysis ON DoctorTreatAnalysis.SystemUserID = SystemUser.SystemUserID"

    With rst
        .open sql, conn, 3, 4
        
        If .recordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Doctor</th>"
            response.write "<th class='myth'>Revenue</th>"
            response.write "<th class='myth'>Patients</th>"
            response.write "<th class='myth'>View More</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                systemUserID = .fields("SystemUserID")
                doctor = .fields("StaffName")
                totalRevenue = totalRevenue + CDbl(.fields("TotalCost"))
                patients = patients + .fields("Count")
                response.write "<tr class='mytr' style='cursor: pointer;'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("StaffName") & "</td>"
                response.write "<td class='mytd'>" & FormatNumber(.fields("TotalCost"), 2) & "</td>"
                response.write "<td class='mytd'>" & .fields("Count") & "</td>"
                response.write "<td class='mytd' style='cursor: pointer; color: blue' onclick='redirectToDetails(""" & Server.HTMLEncode(systemUserID) & """, """ & Server.HTMLEncode(treatmentID) & """, """ & Server.HTMLEncode(periodStart) & """, """ & Server.HTMLEncode(periodEnd) & """)'><u>View More</u></td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop
            
            response.write "<tr>"
                response.write "<td colspan='2' class='mytd' style='text-align:right; font-weight: bold'> TOTAL </td>"
                response.write "<td class='mytd'>" & FormatNumber(totalRevenue, 2) & "</td>"
                response.write "<td class='mytd'>" & patients & "</td>"
                response.write "<td class='mytd'></td>"
            response.write "</tr>"
            
            response.write "</table>"
            response.write "<script>"
            response.write "    function redirectToDetails(systemUserID, treatmentID, periodStart, periodEnd) {"
            response.write "        const baseUrl = 'http://192.168.5.11/rmchms01/wpgPrtPrintLayoutAll.asp';"
            response.write "        const params = new URLSearchParams({"
            response.write "            PrintLayoutName: 'PatientTreatmentDetails',"
            response.write "            PositionForTableName: 'WorkingDay',"
            response.write "            WorkingDayID: '',"
            response.write "            SystemUserID: systemUserID,"
            response.write "            TreatmentID: treatmentID,"
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
