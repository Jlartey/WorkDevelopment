Sub printReport()
    Dim periodStart, periodEnd, drgStoreID, sql, rst
    Dim cnt, drugStoreID, stockDate, drugID, unitCost, availableQty, totalCost, reorderLevel

    ' Get period and store ID from request
    periodStart = getDatePeriodFromDelim(Trim(request.queryString("PrintFilter0")))(0)
    periodEnd = getDatePeriodFromDelim(Trim(request.queryString("PrintFilter0")))(1)
    drgStoreID = Trim(request.queryString("PrintFilter2"))

    ' Construct SQL query with parameters
    sql = "SELECT * FROM drugstocklevel AS d WHERE d.stockdate1 >= ? AND d.stockdate1 <= ?"
    If drgStoreID <> "" Then
        sql = sql & " AND drugstoreid = ?"
    End If

    ' Execute SQL query
    Set rst = server.CreateObject("ADODB.Recordset")
    With rst
        .Open sql, conn, 3, 4, , Array(periodStart, periodEnd, drgStoreID)
        If Not .EOF Then
            ' Output table header
            response.write "<table cellpadding='1' border='1' width='100%' cellspacing='0' >"
            response.write "<tr><th>No.</th><th>Store</th><th>Date</th><th>Drug</th><th>Unit Cost</th><th>Available Qty</th><th>Total</th><th>Reorder Level</th></tr>"

            ' Loop through recordset and output table rows
            .MoveFirst
            Do While Not .EOF
                cnt = cnt + 1
                drugStoreID = .fields("DrugStoreID")
                stockDate = .fields("StockDate1")
                drugID = .fields("DrugID")
                unitCost = .fields("BulkUnitCost")
                availableQty = .fields("AvailableQty")
                totalCost = .fields("TotalCost")
                reorderLevel = .fields("ReOrderLevel")

                ' Output table row
                response.write "<tr>"
                response.write "<td align='center'>" & cnt & "</td>"
                response.write "<td align='center'>" & GetComboName("DrugStore", drugStoreID) & "</td>"
                response.write "<td align='center'>" & FormatDateTime(stockDate, 2) & "</td>"
                response.write "<td align='center'>" & GetComboName("Drug", drugID) & "</td>"
                response.write "<td align='center'>" & FormatNumber(unitCost, 2) & "</td>"
                response.write "<td align='center'>" & FormatNumber(availableQty, 2) & "</td>"
                response.write "<td align='center'>" & FormatNumber(totalCost, 2) & "</td>"
                response.write "<td align='center'>" & FormatNumber(reorderLevel, 2) & "</td>"
                response.write "</tr>"

                .MoveNext
            Loop
            response.write "</table>"
        Else
            response.write "No records found."
        End If
        .Close
    End With
End Sub
