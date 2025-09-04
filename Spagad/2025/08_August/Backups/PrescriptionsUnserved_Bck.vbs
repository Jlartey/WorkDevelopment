'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim periodStart, periodEnd

datePeriod = Trim(Request.QueryString("PrintFilter"))

If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    periodStart = dateArr(0)
    periodEnd = dateArr(1)
End If

tableStyles
dispUndispensedDrugs

Sub dispUndispensedDrugs()
    Dim count, sql, rst, prescriptions
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT "
    sql = sql & "Visitation.PatientID, "
    sql = sql & "Visitation.VisitationID, "
    sql = sql & "STRING_AGG(Drug.DrugName, ', ') AS [Drugs Prescribed] "
    sql = sql & "From Visitation "
    sql = sql & "LEFT JOIN Prescription "
    sql = sql & "ON Visitation.VisitationID = Prescription.VisitationID "
    sql = sql & "LEFT JOIN Drug "
    sql = sql & "ON Prescription.DrugID = Drug.DrugID "
    sql = sql & "WHERE VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    sql = sql & "GROUP BY Visitation.PatientID, Visitation.VisitationID "
    sql = sql & "HAVING STRING_AGG(CAST(Prescription.DrugID AS VARCHAR), ', ') IS NOT NULL"
    
    With rst
        .open sql, conn, 3, 4
        
        If .recordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Visitation ID</th>"
            response.write "<th class='myth'>Patient</th>"
            response.write "<th class='myth'>Prescriptions</th>"
            response.write "<th class='myth'>Drugs Sold</th>"
            response.write "<th class='myth'>Prescriptions Not Served</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                prescriptions = .fields("Drugs Prescribed")
                VisitationID = .fields("VisitationID")
                response.write "<tr class='mytr' onclick='redirectToVisitation(""" & VisitationID & """)' style='cursor: pointer;'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & VisitationID & "</td>"
                response.write "<td class='mytd'>" & GetComboName("Patient", .fields("PatientID")) & "</td>"
                response.write "<td class='mytd'>" & FormatPrescriptions(prescriptions) & "</td>"
                response.write "<td class='mytd'>" & DrugsSold(VisitationID) & "</td>"
                response.write "<td class='mytd'>" & PrescriptionDrugsNotSold(VisitationID) & "</td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop

            response.write "</table>"
            
            response.write "<script>"
            response.write "    function redirectToVisitation(visitationId) {"
            response.write "        const baseUrl = 'http://172.19.0.36/hms/wpgPrtPrintLayoutAll.asp';"
            response.write "        const params = new URLSearchParams({"
            response.write "            PrintLayoutName: 'VisitationRCP',"
            response.write "            PositionForTableName: 'Visitation',"
            response.write "            VisitationID: visitationId,"
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

Function FormatPrescriptions(drugString)
    Dim drugArray
    drugArray = Split(drugString, ",")
    
    Dim htmlOutput
    htmlOutput = "<ol>" & vbCrLf

    Dim i
    For i = 0 To UBound(drugArray)
        Dim drug
        drug = Trim(drugArray(i))
        If Len(drug) > 0 Then
            htmlOutput = htmlOutput & "    <li>" & drug & "</li>" & vbCrLf
        End If
    Next
    htmlOutput = htmlOutput & "</ol>"
    
    FormatPrescriptions = htmlOutput
End Function

Function DrugsSold(VisitationID)
    Dim sql, rst, cmd, htmlOutput

    Set rst = CreateObject("ADODB.Recordset")
    Set cmd = CreateObject("ADODB.Command")
    
    sql = "WITH CombinedDrugSales AS ( " & _
          "    SELECT DrugID, VisitationID " & _
          "    FROM DrugSaleItems " & _
          "    WHERE VisitationID = ? " & _
          "    UNION ALL " & _
          "    SELECT DrugID, VisitationID " & _
          "    FROM DrugSaleItems2 " & _
          "    WHERE VisitationID = ? " & _
          ") " & _
          "SELECT d.DrugName " & _
          "FROM CombinedDrugSales cds " & _
          "JOIN Drug d ON cds.DrugID = d.DrugID " & _
          "ORDER BY d.DrugName"
    
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1
        .Parameters.append .CreateParameter("@VisitationID1", 200, 1, 20, VisitationID)
        .Parameters.append .CreateParameter("@VisitationID2", 200, 1, 20, VisitationID)
    End With
    
    rst.open cmd, , 3, 4
    
    htmlOutput = "<ol>" & vbCrLf
    If Not rst.EOF Then
        While Not rst.EOF
            Dim drugName
            drugName = Trim(rst("DrugName"))
            If Len(drugName) > 0 Then
                htmlOutput = htmlOutput & "    <li>" & Server.HTMLEncode(drugName) & "</li>" & vbCrLf
            End If
            rst.MoveNext
        Wend
    End If
    htmlOutput = htmlOutput & "</ol>"
    
    rst.Close
    Set rst = Nothing
    Set cmd = Nothing
    
    DrugsSold = htmlOutput
End Function

Function PrescriptionDrugsNotSold(VisitationID)
    Dim sql, rst, htmlOutput

    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "WITH DrugsNotSold AS ( " & _
          "    SELECT Prescription.DrugID " & _
          "    FROM Prescription " & _
          "    WHERE VisitationID = '" & VisitationID & "' " & _
          "    AND Prescription.DrugID NOT IN ( " & _
          "        SELECT DrugSaleItems.DrugID FROM DrugSaleItems " & _
          "        WHERE VisitationID = '" & VisitationID & "' " & _
          "        UNION " & _
          "        SELECT DrugSaleItems2.DrugID FROM DrugSaleItems2 " & _
          "        WHERE VisitationID = '" & VisitationID & "' " & _
          "    ) " & _
          ") " & _
          "SELECT d.DrugName " & _
          "FROM Drug d " & _
          "JOIN DrugsNotSold dns " & _
          "    ON d.DrugID = dns.DrugID"
    
    rst.open sql, conn, 3, 4
    
    
    If rst.recordCount > 0 Then
        htmlOutput = "<ol>" & vbCrLf
        While Not rst.EOF
            Dim drugName
            drugName = Trim(rst("DrugName"))
            If Len(drugName) > 0 Then
                htmlOutput = htmlOutput & "    <li>" & Server.HTMLEncode(drugName) & "</li>" & vbCrLf
            End If
            rst.MoveNext
        Wend
        htmlOutput = htmlOutput & "</ol>"
    Else
        htmlOutput = "<p>All prescriptions served</p>"
        
    End If
    
    rst.Close
    Set rst = Nothing
    
    PrescriptionDrugsNotSold = htmlOutput
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
