'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
addCSS
generateReport

Sub generateReport()
    Dim sql, rst, cnt, store
    Set rst = CreateObject("ADODB.Recordset")
    store = UCase(jSchd)
    

    sql = "SELECT dg.DrugName, ds.DrugID, ds.UnitOfMeasureID, ds.RetailUnitCost,"
    sql = sql & " dm.ItemUnitCost, ds.ReOrderLevel, ds.AvailableQty,"
    sql = sql & " CASE"
    sql = sql & "    WHEN ds.AvailableQty >= ds.ReOrderLevel + 20 THEN '#b2f2bb'"
    sql = sql & "    WHEN ds.AvailableQty <= ds.ReOrderLevel THEN '#ffc9c9'"
    sql = sql & "    WHEN ds.AvailableQty > ds.ReOrderLevel AND ds.AvailableQty < ds.ReOrderLevel + 20 THEN '#ffec99'"
    sql = sql & "     ELSE 'grey'"
    sql = sql & "   END AS ColorCode"
    sql = sql & " FROM DrugStockLevel  ds"
    sql = sql & " JOIN Drug dg ON dg.DrugID = ds.DrugID"
    sql = sql & " JOIN DrugPriceMatrix2 dm ON dm.DrugID = ds.DrugID"
    sql = sql & " WHERE ds.DrugStatusID = 'IST001'AND ds.DrugStoreID = '" & store & "' AND dm.InsuranceTypeID = 'I100'"
    sql = sql & " ORDER BY dg.DrugName"

    cnt = 0
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
            If .RecordCount > 0 Then
                response.write "<table class = 'anaesthesia' > "
                response.write "    <thead> "
                response.write "    <tr class = 'anaesthesia'>"
                response.write "        <th colspan = '7'>Generated Reorder Level Report for Drugs</th>"
                response.write "    </tr>"
                response.write "    <tr class = 'tHead'> "
                response.write "        <th>No.</th> "
                response.write "        <th>Drug</th> "
                response.write "        <th>UoM</th> "
                response.write "        <th>Retail Unit Cost</th> "
                response.write "        <th>Item Unit Cost</th> "
                response.write "        <th>Reorder Level</th> "
                response.write "        <th>Available Qty</th> "
                response.write "    </tr> "
                response.write "    </thead><tbody> "
                .MoveFirst
                Do While Not .EOF
                    cnt = cnt + 1
                    colorCode = .fields("ColorCode")
                    response.write "  <tr class = 'queryData' style=""background-color: " & colorCode & ";""> "
                    response.write "      <td>" & cnt & "</td> "
                    response.write "      <td>" & .fields("DrugName") & " [" & .fields("DrugID") & "]</td> "
                    response.write "      <td>" & GetComboName("UnitOfMeasure", .fields("UnitOfMeasureID")) & "</td> "
                    response.write "      <td>" & .fields("RetailUnitCost") & "</td> "
                    response.write "      <td>" & .fields("ItemUnitCost") & "</td> "
                    response.write "      <td>" & .fields("ReOrderLevel") & "</td> "
                    response.write "      <td>" & .fields("AvailableQty") & "</td> "
                    response.write "  </tr> "
                    .MoveNext
                Loop
            End If
            If .RecordCount = 0 Then
                response.write "<div class = 'errorMsg'>No data available to display.</div>"
            End If
            response.write "</tbody></table><br><br>"
        .Close
        Set rst = Nothing
    End With

End Sub

Sub addCSS()
  With response
    .write " <style> "
    .write "    .anaesthesia, .anaesthesia th, .anaesthesia td{ "
    .write "        border: 1px solid silver; "
    .write "        border-collapse: collapse; "
    .write "        padding: 5px; "
    .write "    } "
    .write "    .anaesthesia{ "
    .write "        width: 65vw; "
    .write "        margin: 0 auto; "
    .write "        font-family: sans-serif; "
    .write "        font-size: 13px; "
    .write "        box-sizing: border-box; "
    .write "    }"
    .write "    .anaesthesia tr{page-break-inside:avoid; "
    .write "        page-break-after:auto "
    .write "    } "
    .write "    .anaesthesia th, .anaesthesia td { "
    .write "        border: 1px solid silver; "
    .write "        text-align: center; "
    .write "        padding: 5px; "
    .write "        font-size:13px; "
    .write "        margin: 0 auto; "
    .write "    } "
    .write "    .tHead{ "
    .write "        position: sticky; top: 0; "
    .write "    }  "
    .write "    .queryData td{ "
    .write "        font-size: 12; "
    .write "        text-transform: uppercase; "
    .write "    }  "
    .write "    .anaesthesia th{ "
    .write "        background-color: #212529; "
    .write "        text-align: center; "
    .write "        font-weight: bold;"
    .write "        font-size: 14px;"
    .write "        color:#fff;"
    .write "   } "
    .write " </style> "
  End With
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
