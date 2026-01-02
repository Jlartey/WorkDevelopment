'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>



tableStyles
dispTotalDrugStock

Sub dispTotalDrugStock()
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT "
    sql = sql & " d.DrugName,"
    sql = sql & " COALESCE(SUM(CASE WHEN ds.DrugStoreID = 'EMERGENCY' THEN ds.AvailableQty ELSE 0 END), 0) AS 'EMERGENCY',"
    sql = sql & " COALESCE(SUM(CASE WHEN ds.DrugStoreID = 'M0310' THEN ds.AvailableQty ELSE 0 END), 0) AS 'M0310',"
    sql = sql & " COALESCE(SUM(CASE WHEN ds.DrugStoreID = 'S21' THEN ds.AvailableQty ELSE 0 END), 0) AS 'S21',"
    sql = sql & " COALESCE(SUM(CASE WHEN ds.DrugStoreID = 's21a' THEN ds.AvailableQty ELSE 0 END), 0) AS 'S21A',"
    sql = sql & " COALESCE(SUM(CASE WHEN ds.DrugStoreID = 'S21C' THEN ds.AvailableQty ELSE 0 END), 0) AS 'S21C',"
    sql = sql & " COALESCE(SUM(CASE WHEN ds.DrugStoreID = 's22' THEN ds.AvailableQty ELSE 0 END), 0) AS 'S22',"
    sql = sql & " COALESCE(SUM(CASE WHEN ds.DrugStoreID = 'THEATRE1' THEN ds.AvailableQty ELSE 0 END), 0) AS 'THEATRE1',"
    sql = sql & " COALESCE(SUM(CASE WHEN ds.DrugStoreID = 'W011' THEN ds.AvailableQty ELSE 0 END), 0) AS 'W011',"
    sql = sql & " SUM(ds.AvailableQty) As [Total Drug Quantity]"
    sql = sql & " FROM Drug d"
    sql = sql & " INNER JOIN DrugStockLevel ds ON d.DrugID = ds.DrugID"
    sql = sql & " GROUP BY d.DrugName"
    sql = sql & " ORDER BY [Total Drug Quantity] DESC;"



    With rst
        .open sql, conn, 3, 4
        
        If .recordCount > 0 Then
            
            response.write "<h1> DRUG STOCK VALUES ACROSS THE DRUGSTORES</h1>"
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>Drug</th>"
            response.write "<th class='myth'>EMERGENCY</th>"
            response.write "<th class='myth'>M0310</th>"
            response.write "<th class='myth'>S21</th>"
            response.write "<th class='myth'>S21A</th>"
            response.write "<th class='myth'>S21C</th>"
            response.write "<th class='myth'>S22</th>"
            response.write "<th class='myth'>THEATRE1</th>"
            response.write "<th class='myth'>W011</th>"
            response.write "<th class='myth'>Drug Total</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & .fields("DrugName") & "</td>"
                response.write "<td class='mytd'>" & .fields("EMERGENCY") & "</td>"
                response.write "<td class='mytd'>" & .fields("M0310") & "</td>"
                response.write "<td class='mytd'>" & .fields("S21") & "</td>"
                response.write "<td class='mytd'>" & .fields("S21A") & "</td>"
                response.write "<td class='mytd'>" & .fields("S21C") & "</td>"
                response.write "<td class='mytd'>" & .fields("S22") & "</td>"
                response.write "<td class='mytd'>" & .fields("THEATRE1") & "</td>"
                response.write "<td class='mytd'>" & .fields("W011") & "</td>"
                response.write "<td class='mytd'>" & .fields("Total Drug Quantity") & "</td>"
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
        response.write "    width: 85vw;"
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


