'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>


tableStyles
dispDrugEOQ

Sub dispDrugEOQ()
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT d.DrugID, d.DrugName, d.ReOrderLevelQty, d.MaxStockQty, d.Rate2, dsl.AvailableQty,"
    sql = sql & "   CASE"
    sql = sql & "     WHEN (dSl.AvailableQty <= d.ReOrderLevelQty) OR (dSl.AvailableQty <= d.Rate2)"
    sql = sql & "     THEN (d.MaxStockQty - dsl.AvailableQty)"
    sql = sql & "    ELSE 0"
    sql = sql & "   END As eoq"
    sql = sql & " FROM Drug d"
    sql = sql & " JOIN DrugStockLevel dsl"
    sql = sql & " ON d.DrugID = dsl.DrugID"
    sql = sql & " WHERE d.DrugStatusID = 'IST001'"
    sql = sql & " AND dsl.DrugStoreID = 's22'"
    sql = sql & " AND dsl.AvailableQty > 0"
    sql = sql & " AND d.BillGroupID <> 'BG008'"
    sql = sql & " ORDER BY d.Drugname"

    With rst
        .open sql, conn, 0, 1
        
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Drug ID</th>"
            response.write "<th class='myth'>Drug</th>"
            response.write "<th class='myth'>Reorder Level</th>"
            response.write "<th class='myth'>Maximum Stock</th>"
            response.write "<th class='myth'>Minimum Stock</th>"
            response.write "<th class='myth'>Current Stock</th>"
            response.write "<th class='myth'>Economic Order Quantity</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("DrugID") & "</td>"
                response.write "<td class='mytd'>" & .fields("DrugName") & "</td>"
                response.write "<td class='mytd'>" & .fields("ReorderLevelQty") & "</td>"
                response.write "<td class='mytd'>" & .fields("MaxStockQty") & "</td>"
                response.write "<td class='mytd'>" & .fields("Rate2") & "</td>"
                response.write "<td class='mytd'>" & .fields("AvailableQty") & "</td>"
                response.write "<td class='mytd'>" & .fields("EOQ") & "</td>"
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
        response.write "  position: sticky;"
        response.write "  top: 0;"
        response.write "  background-color: #f2f2f2;"
        response.write "  color: #333;"
        response.write "  font-weight: bold;"
        response.write "  text-transform: uppercase;"
        response.write "  z-index: 10;"
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


