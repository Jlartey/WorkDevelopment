Sub displayStockLevel03()
    
    Dim sql, periodStart, periodEnd, datePeriod, count, selectedDrugStoreIDs, idsArr, formattedIDs, id, dateArr
    'Dim conn, rst

    ' Retrieve query parameters
    datePeriod = Trim(request.querystring("Dateperiod"))
    selectedDrugStoreIDs = Trim(request.querystring("DrugStoreID"))

    ' Parse date period
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If

    ' Format selected drug store IDs
    If selectedDrugStoreIDs <> "" Then
        idsArr = Split(selectedDrugStoreIDs, ",")
        For Each id In idsArr
            formattedIDs = formattedIDs & "'" & Trim(id) & "',"
        Next
        ' Remove the trailing comma
        formattedIDs = Left(formattedIDs, Len(formattedIDs) - 1)
    End If

    ' Construct SQL query
    sql = "SELECT DrugStore.DrugStoreID, DrugStore.DrugStoreName, SUM(DrugStockLevel.AvailableQty) AS [StockLevel], "
    sql = sql & "CONVERT(VARCHAR(20), StockDate1, 103) AS [StockDate] "
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
    sql = sql & "GROUP BY DrugStore.DrugStoreID, DrugStore.DrugStoreName, CONVERT(VARCHAR(20), StockDate1, 103) "
    sql = sql & "ORDER BY [StockLevel] DESC, [StockDate]"

    'response.write sql

    ' Initialize and open database connection
    ' Set conn = CreateObject("ADODB.Connection")
    ' conn.Open "your_connection_string" ' Replace with your actual connection string

    Set rst = CreateObject("ADODB.Recordset")
    rst.Open sql, conn, 3, 4

    ' Populate dropdown and display table
    If rst.RecordCount > 0 Then
        rst.MoveFirst

        ' Populate dropdown
        Response.Write "<div>"
        Response.Write "        <label for='pharmacy' class='font-style'>Select Pharmacy:</label><br>"
        Response.Write "        <select id='pharmacy' name='pharmacy' multiple class='mult-select-tag'>"
        Do Until rst.EOF
            Response.Write "            <option value='" & rst("DrugStoreID") & "'>" & rst("DrugStoreName") & "</option>"
            rst.MoveNext
        Loop
        Response.Write "        </select>"
        Response.Write "</div>"

        rst.MoveFirst ' Reset recordset to the first record

        ' Display table
        Response.Write "<table class='mytable'>"
        Response.Write "<tr>"
        Response.Write "<th class='myth'>No.</th>"
        Response.Write "<th class='myth'>Pharmacy</th>"
        Response.Write "<th class='myth'>Stock Level</th>"
        Response.Write "<th class='myth'>Stock Date</th>"
        Response.Write "</tr>"

        count = 0
        Do While Not rst.EOF
            count = count + 1
            Response.Write "<tr>"
            Response.Write "<td class='mytd' align='center'>" & count & "</td>"
            Response.Write "<td class='mytd'>" & rst("DrugStoreName") & "</td>"
            Response.Write "<td class='mytd' align='center'>" & rst("StockLevel") & "</td>"
            Response.Write "<td class='mytd' align='center'>" & rst("StockDate") & "</td>"
            Response.Write "</tr>"
            Response.Flush
            rst.MoveNext
        Loop
        Response.Write "</table>"
    End If

    ' Close recordset and connection
    rst.Close
    conn.Close
    Set rst = Nothing
    Set conn = Nothing

    ' Output HTML Form for date selection
    Response.Write "    <form id='dateForm'>"
    Response.Write "    <div class='container' style='display: flex; align-items: center; justify-content: center'> "
    Response.Write "        <div> "
    Response.Write "            <label for='from'>From</label> "
    Response.Write "            <input type='date' name='from' id='from'> "
    Response.Write "        </div> "
    Response.Write "        <div> "
    Response.Write "            <label for='to' style='margin-left: 10px'>To</label> "
    Response.Write "            <input type='date' name='to' id='to'> "
    Response.Write "        </div> "
    Response.Write "        <div> "
    Response.Write "            <button type='button' onclick='updateUrl()'>Show Data</button> <br />"
    Response.Write "        </div>    "
    Response.Write "    </div> "
    Response.Write "   </form>"
    
    Response.Write "<script src='https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js'></script>"
    Response.Write "<script>"
        Response.Write "    new MultiSelectTag('pharmacy', {"
        Response.Write "        rounded: true,"
        Response.Write "        shadow: true,"
        Response.Write "        placeholder: 'Search',"
        Response.Write "        tagColor: {"
        Response.Write "            textColor: '#327b2c',"
        Response.Write "            borderColor: '#92e681',"
        Response.Write "            bgColor: '#eaffe6',"
        Response.Write "        },"
        Response.Write "        onChange: function (values) {"
        Response.Write "            console.log(values);"
        Response.Write "        },"
        Response.Write "    });"
        Response.Write "    function updateUrl() {"
        Response.Write "        const fromDate = document.getElementById('from').value;"
        Response.Write "        const toDate = document.getElementById('to').value;"
        Response.Write "        const pharmacy = Array.from(document.getElementById('pharmacy').selectedOptions).map(option => option.value).join(',');"
        Response.Write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp';"
        Response.Write "        const params = new URLSearchParams({"
        Response.Write "            PrintLayoutName: 'displayStockLevel',"
        Response.Write "            PositionForTableName: 'WorkingDay',"
        Response.Write "            WorkingDayID: '',"
        Response.Write "            Dateperiod: `${fromDate}||${toDate}`,"
        Response.Write "            DrugStoreID: pharmacy"
        Response.Write "        });"
        Response.Write "        const newUrl = baseUrl + '?' + params.toString();"
        Response.Write "        window.location.href = newUrl;"
        Response.Write "        console.log(newUrl);"
        Response.Write "    }"
    Response.Write "</script>"

    If (periodStart <> "" And periodEnd <> "") Then
        Response.Write "<h2 class='font-style'>FROM: " & periodStart & " TO: " & periodEnd & "</h2>"
    Else
        Response.Write "<h2 class='font-style'>FROM: 2018-01-01 TO: 2018-01-31</h2>"
    End If

End Sub
