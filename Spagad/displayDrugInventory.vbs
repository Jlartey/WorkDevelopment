'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Styling
displayDrugInventory

Sub displayDrugInventory()
    Dim sql, periodStart, periodEnd, datePeriod, count
    datePeriod = Trim(request.querystring("Dateperiod"))
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If
    
    Set rst = CreateObject("ADODB.RecordSet")
    
    sql = "SELECT "
    sql = sql & "Drug.DrugID, "
    sql = sql & "Drug.DrugName, "
    sql = sql & "UnitOfMeasure.UnitOfMeasureName, "
    sql = sql & "IncomingDrugItems.Qty, "
    sql = sql & "Drug.BulkUnitCost, "
    sql = sql & "IncomingDrugItems.DrugPurOrderID, "
    sql = sql & "IncomingDrugItems.ReturnQty, "
    sql = sql & "SystemUser.SystemUserName "
    sql = sql & "FROM Drug "
    sql = sql & "JOIN IncomingDrugItems "
    sql = sql & "ON Drug.DrugID = IncomingDrugItems.DrugID "
    sql = sql & "JOIN UnitOfMeasure "
    sql = sql & "ON Drug.UnitOfMeasureID = UnitOfMeasure.UnitOfMeasureID "
    sql = sql & "JOIN SystemUser "
    sql = sql & "ON SystemUser.SystemUserID = IncomingDrugItems.SystemUserID "
    
    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "WHERE IncomingDrugItems.EntryDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    Else
        sql = sql & "WHERE IncomingDrugItems.EntryDate BETWEEN '2018-01-01' AND '2024-01-01'"
    End If
    
    'response.write sql
    
    'Display the DatePicker
    response.write "<h2>Showing Data From " & periodStart & " To " & periodEnd & " </h2>"
    
    response.write "<form id='dateForm'> "
    response.write "    <div class='container' style='display: flex;'> "
    response.write "        <div> "
    response.write "            <label for='from'>From</label> "
    response.write "            <input type='date' name='from' id='from'> "
    response.write "        </div> "
    response.write "        <div> "
    response.write "            <label for='to' style='margin-left: 10px'>To</label> "
    response.write "            <input type='date' name='to' id='to'> "
    response.write "        </div> "
    response.write "        <div> "
    response.write "            <button type='button' onclick='updateUrl()'>Show Data</button> <br />"
    response.write "        </div>    "
    response.write "    </div> "
    response.write "</form> "

    response.write " <br />"
    response.write "<script> "
    response.write "    function updateUrl() { "
    response.write "        const fromDate = document.getElementById('from').value; "
    response.write "        const toDate = document.getElementById('to').value; "
    response.write "        const baseUrl = 'http://192.168.5.11/rmchms01/wpgPrtPrintLayoutAll.asp'; "
    response.write "        const params = new URLSearchParams({ "
    response.write "            PrintLayoutName: 'displayDrugInventory', "
    response.write "            PositionForTableName: 'WorkingDay', "
    response.write "            WorkingDayID: '' ,"
    response.write "            Dateperiod: fromDate + '||' + toDate"
    response.write "        }); "
    response.write "        const newUrl = baseUrl + '?' + params.toString(); "
    response.write "        console.log(newUrl); "
    response.write "        window.location.href = newUrl; "
    response.write "    } "
    response.write "</script> "
    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then

            .MoveFirst
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr>"
                response.write "<th class='myth'> Count </th>"
                response.write "<th class='myth'> Item Code </th>"
                response.write "<th class='myth'> Description of Medication </th>"
                response.write "<th class='myth'>Unit Of Measure</th>"
                response.write "<th class='myth'>Quantity Requisitioned</th>"
                response.write "<th class='myth'>Unit Cost</th>"
                response.write "<th class='myth'>Total Cost</th>"
                response.write "<th class='myth'>Pick List No</th>"
                response.write "<th class='myth'>Quantity Received</th>"
                response.write "<th class='myth'>Reviewed By</th>"
            response.write "</tr>"
            Do While Not .EOF
                count = count + 1
                response.write "<tr>"
                    response.write "<td class='mytd' align='center'>" & count & "</td>"
                    response.write "<td class='mytd'>" & .fields("DrugID") & "</td>"
                    response.write "<td class='mytd'>" & .fields("DrugName") & "</td>"
                    response.write "<td class='mytd'>" & .fields("UnitOfMeasureName") & "</td>"
                    response.write "<td class='mytd' align='center'>" & .fields("Qty") & "</td>"
                    response.write "<td class='mytd' align='center'>" & FormatNumber(.fields("BulkUnitCost"), 2) & "</td>"
                    response.write "<td class='mytd' align='center'>" & FormatNumber(.fields("Qty") * .fields("BulkUnitCost"), 2) & "</td>"
                    response.write "<td class='mytd'>" & .fields("DrugPurOrderID") & "</td>"
                    response.write "<td class='mytd' align='center'>" & .fields("ReturnQty") & "</td>"
                    response.write "<td class='mytd'>" & .fields("SystemUserName") & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
            response.write "</table>"
        End If
    End With
End Sub

Sub Styling()
    response.write " <style>"
        response.write " table {"
        response.write "     width: 75vw;"
        response.write "     font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
        response.write "     border-collapse: collapse;"
       
        response.write " }"
        
        response.write " .myth, .mytd {"
        response.write "     border: 1px solid #ddd;"
        response.write "     padding: 10px;"
        response.write " }"
        
        response.write " .mytd {"
        response.write "     text-alig: 1px solid #ddd;"
        response.write "     padding: 8px;"
        response.write " }"
        
        response.write "  tr:nth-child(even) {"
        response.write "    background-color: #f9f9f9;"
        response.write " } "
        
        response.write " .myth {"
        response.write "     background-color: #c2c2c2;"
        response.write "     color: black;"
        response.write "     text-align: center; "
        response.write "     text-transform: uppercase; "
        response.write "     font-size: 18px;"
        response.write " }"
        
        response.write "  button {"
        response.write "     background-color: #0236c4;"
        response.write "     border-radius: 5px;"
        response.write "     border: none;"
        response.write "     margin-left: 50px;"
        response.write "     padding: 5px 20px;"
        response.write "     color: white;"
        response.write "     cursor: pointer;"
        response.write "  }"
        
        response.write "  #to, #from {"
        response.write "    padding: 5px;"
        response.write "    border-radius: 5px;"
        response.write "    cursor: pointer;"
        response.write "  }"
    response.write " </style>"

End Sub
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
