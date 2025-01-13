Dim rst, rst2, sql, cnt, periodStart, periodEnd, datePeriod, selectedDrugStoreIDs, idsArr, formattedIDs, id
    
    datePeriod = Trim(request.querystring("Dateperiod"))
    selectedDrugStoreIDs = Trim(request.querystring("DrugStoreID"))
    
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If
    
    cnt = 0

    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    
    If selectedDrugStoreIDs <> "" Then
        idsArr = Split(selectedDrugStoreIDs, ",")
        For Each id In idsArr
            formattedIDs = formattedIDs & "'" & Trim(id) & "',"
        Next
        ' Remove the trailing comma
        formattedIDs = Left(formattedIDs, Len(formattedIDs) - 1)
    End If
    
    sql = "SELECT DrugStore.DrugStoreID, DrugStoreName, COUNT(DrugID)[StockLevel], "
    sql = sql & "CONVERT(VARCHAR(20), StockDate1, 103)[StockDate] "
    sql = sql & "FROM DrugStore JOIN DrugStockLevel "
    sql = sql & "ON DrugStockLevel.DrugStoreID = DrugStore.DrugStoreID "
    sql = sql & "WHERE CONVERT(DATE, StockDate1) "
    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    Else
        sql = sql & "BETWEEN '2018-01-01' AND '2018-01-31' "
    End If
    If selectedDrugStoreIDs <> "" Then
        sql = sql & " AND DrugStore.DrugStoreID IN (" & formattedIDs & ") "
    End If
    sql = sql & "GROUP BY DrugStore.DrugStoreID, DrugStoreName, CONVERT(VARCHAR(20), StockDate1, 103) "
    sql = sql & "ORDER BY [StockLevel] DESC, [StockDate]  "
     
    response.write "<h3>Stock Level Per Pharmacy</h3>"
    
    'response.write sql
    'response.write selectedDrugStoreIDs
    
    With rst
        .open sql, conn, 3, 4
        
        If .recordCount > 0 Then
            .movefirst
            
            response.write "    <div>"
            response.write "        <label for='pharmacy'>Select Pharmacy:</label><br>"
            response.write "        <select id='pharmacy' name='pharmacy' multiple class='mult-select-tag'>"
            ' Loop through the recordset and populate the dropdown
            Do Until .EOF
                response.write "            <option value='" & .fields("DrugStoreID") & "'>" & .fields("DrugStoreName") & "</option>"
                .MoveNext
            Loop
            response.write "        </select>"
            response.write "    </div>"
        End If
        .Close
    End With
            
            ' Output HTML form for date selection
            response.write "<form id='dateForm'>"
            response.write "<div class='form-container'>"
            response.write "    <div>"
            response.write "        <label for='from'>From</label>"
            response.write "        <input type='date' name='from' id='from'>"
            response.write "    </div>"
            response.write "    <div>"
            response.write "        <label for='to'>To</label>"
            response.write "        <input type='date' name='to' id='to'>"
            response.write "    </div>"
            
'            response.write "    <div>"
'            response.write "        <label for='pharmacy'>Select Pharmacy:</label><br>"
'            response.write "        <select id='pharmacy' name='pharmacy' multiple class='mult-select-tag'>"
'            ' Loop through the recordset and populate the dropdown
'            Do Until .EOF
'                response.write "            <option value='" & .fields("DrugStoreID") & "'>" & .fields("DrugStoreName") & "</option>"
'                .MoveNext
'            Loop
'            response.write "        </select>"
'            response.write "    </div>"
'        End If
'        .Close
'    End With
            

            response.write "    <div>"
            response.write "        <button type='button' onclick='updateUrl()'>Display Report</button>"
            response.write "    </div>"
            response.write "</div>"
            response.write "</form>"

            response.write "<script src='https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js'></script>"
            response.write "<script>"
            response.write "    new MultiSelectTag('pharmacy', {"
            response.write "        rounded: true,"
            response.write "        shadow: true,"
            response.write "        placeholder: 'Search',"
            response.write "        tagColor: {"
            response.write "            textColor: '#327b2c',"
            response.write "            borderColor: '#92e681',"
            response.write "            bgColor: '#eaffe6',"
            response.write "        },"
            response.write "        onChange: function (values) {"
            response.write "            console.log(values);"
            response.write "        },"
            response.write "    });"
            response.write "    function updateUrl() {"
            response.write "        const fromDate = document.getElementById('from').value;"
            response.write "        const toDate = document.getElementById('to').value;"
            response.write "        const pharmacy = Array.from(document.getElementById('pharmacy').selectedOptions).map(option => option.value).join(',');"
            response.write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp';"
            response.write "        const params = new URLSearchParams({"
            response.write "            PrintLayoutName: 'PharmStockLvlByBranchRay',"
            response.write "            PositionForTableName: 'WorkingDay',"
            response.write "            WorkingDayID: '',"
            response.write "            Dateperiod: `${fromDate}||${toDate}`,"
            response.write "            DrugStoreID: pharmacy"
            response.write "        });"
            response.write "        const newUrl = baseUrl + '?' + params.toString();"
            response.write "        window.location.href = newUrl;"
            response.write "        console.log(newUrl);"
            response.write "    }"
            response.write "</script>"

            
            If (periodStart <> "" And periodEnd <> "") Then
            response.write "<h5>FROM: " & periodStart & " TO: " & periodEnd & "</h5>"
            response.write "<h5>BRANCH: " & selectedDrugStoreIDs & "</h5>"
            Else
            response.write "<h5>FROM: 2018-01-01 TO: 2018-01-31</h5>"
            End If
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr>"
            response.write "<th class='myth'>Serial No.</th>"
            response.write "<th class='myth'>Pharmacy</th>"
            response.write "<th class='myth'>Stock Level</th>"
            response.write "<th class='myth'>Stock Date</th>"
            response.write "</tr>"
            
        With rst2
        .open sql, conn, 3, 4
        
        If .recordCount > 0 Then
            .movefirst
            Do While Not .EOF
                cnt = cnt + 1
                response.write "<tr>"
                response.write "<td class='mytd'>" & cnt & "</td>"
                response.write "<td class='mytd'>" & .fields("DrugStoreName") & "</td>"
                response.write "<td class='mytd'>" & .fields("StockLevel") & "</td>"
                response.write "<td class='mytd'>" & .fields("StockDate") & "</td>"
                response.write "</tr>"
                response.flush
                .MoveNext
            Loop
            response.write "</table>"
        End If
        .Close
    End With
End Sub
