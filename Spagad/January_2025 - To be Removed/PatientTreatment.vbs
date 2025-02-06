'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim yearID

yearID = Trim(Request("PrintFilter"))
tableStyles
dispTreatmentSummary

Sub dispTreatmentSummary()
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT workingmonthID AS month, "
    sql = sql & " SUM(CASE WHEN treatmentID = 'T000011' THEN FinalAmt ELSE 0 END) AS wound_dressing, "
    sql = sql & " SUM(CASE WHEN treatmentID = 'T000012' THEN FinalAmt ELSE 0 END) AS bed_occupancy, "
    sql = sql & " SUM(FinalAmt) AS total_cost, "
    sql = sql & " COUNT(DISTINCT patientID) AS patient_tot, "
    sql = sql & " COUNT(VisitationID) AS visit_tot "
    sql = sql & "FROM treatCharges "
    sql = sql & "WHERE treatmentID IN ('T000011', 'T000012') AND workingYearID = '" & yearID & "' "
    sql = sql & "GROUP BY workingmonthID "
    sql = sql & "ORDER BY month;"

    With rst
        .Open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Month</th>"
            response.write "<th class='myth'>Wound Dressing</th>"
            response.write "<th class='myth'>Bed Occupancy</th>"
            response.write "<th class='myth'>Total Cost</th>"
            response.write "<th class='myth'>Patient Total</th>"
            response.write "<th class='myth'>Visit Total</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("month") & "</td>"
                response.write "<td class='mytd'>" & .fields("wound_dressing") & "</td>"
                response.write "<td class='mytd'>" & .fields("bed_occupancy") & "</td>"
                response.write "<td class='mytd'>" & .fields("total_cost") & "</td>"
                response.write "<td class='mytd'>" & .fields("patient_tot") & "</td>"
                response.write "<td class='mytd'>" & .fields("visit_tot") & "</td>"
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
