'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

tableStyles
dispTotalItemStock

Sub dispTotalItemStock()
    Dim count, sql, rst, storeArray, numStores, i, storeCols, rowBatch
    count = 1
    rowBatch = 50

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT DISTINCT ItemStoreID FROM ItemStockLevel"

    With rst
        .open sql, conn, 3, 4
        
        If .recordCount > 0 Then
            storeArray = .GetRows(.recordCount)
            numStores = UBound(storeArray, 2) + 1
            response.write "<h1>ITEM STOCK LEVELS ACROSS THE ITEM STORES</h1> "
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<thead><tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>ITEM</th>"
            
            For i = 0 To numStores - 1
                response.write "<th class='myth'>" & storeArray(0, i) & "</th>"
            Next
            response.write "<th class='myth'>Total</th>"
            
            response.write "</tr></thead>"
            response.write "<tbody>"
            
            response.Flush
        End If
        .Close
    End With
    Set rst = Nothing

    
    storeCols = ""
    For i = 0 To numStores - 1
        If i > 0 Then storeCols = storeCols & ", "
        storeCols = storeCols & "COALESCE(SUM(CASE WHEN isl.ItemStoreID = '" & Replace(storeArray(0, i), "'", "''") & "' THEN isl.AvailableQty ELSE 0 END), 0) AS '" & storeArray(0, i) & "'"
    Next

    Set rst2 = CreateObject("ADODB.Recordset")

    sql = "SELECT i.ItemName, " & storeCols
    sql = sql & ", SUM(isl.AvailableQty) AS [Total Item Quantity]"
    sql = sql & " FROM Items i"
    sql = sql & " INNER JOIN ItemStockLevel isl ON i.ItemID = isL.ItemID"
    sql = sql & " GROUP BY i.ItemName"
    sql = sql & " ORDER BY i.ItemName;"
    
    'response.write sql
    With rst2
        .open sql, conn, 3, 4

        If .recordCount > 0 Then
            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("ItemName") & "</td>"
                
                For i = 0 To numStores - 1
                    response.write "<td class='mytd'>" & .fields(storeArray(0, i)) & "</td>"
                Next
                response.write "<td class='mytd'>" & .fields("Total Item Quantity") & "</td>"

                response.write "</tr>"

                .MoveNext
                count = count + 1
                
                If (count - 1) Mod rowBatch = 0 And count > 1 Then
                    response.Flush
                End If
            Loop
            
            response.Flush
            response.write "</tbody></table>"
        Else
            response.write "<h1>No records found</h1>"
            response.write "</table>"
        End If
        
        .Close
    End With
    Set rst2 = Nothing
End Sub

Sub tableStyles()
    response.write "<style>"
    response.write ".mytable {"
    response.write "    width: 80vw;"
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
    response.write "    position: sticky;"
    response.write "    top: 0;"
    response.write "    z-index: 10;"
    response.write "    text-transform: uppercase;"
    response.write "}"
    response.write ".mytr:nth-child(even) {"
    response.write "    background-color: #f9f9f9;"
    response.write "}"
    response.write ".mytr:hover {"
    response.write "    background-color: #f1f1f1;"
    response.write "}"
    response.write "h1 {"
    response.write "    font-size: 18px;"
    response.write "    color: #000;"
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




