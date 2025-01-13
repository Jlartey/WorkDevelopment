Sub displayStockLevel()

    Dim sql, periodStart, periodEnd, datePeriod, count, selectedDrugStoreIDs, idsArr, formattedIDs, id, dateArr
    Dim rstDropdown, rstMain
    Dim dropdownOptions, optionHTML

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

    ' Construct SQL query for dropdown options (all pharmacies)
    sql = "SELECT DrugStoreID, DrugStoreName FROM DrugStore"

    ' Initialize and open database connection for dropdown options
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.Open sql, conn, 3, 4

    ' Populate dropdown options
    dropdownOptions = ""

    With rstDropdown
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                optionHTML = "            <option value='" & .Fields("DrugStoreID") & "'>" & .Fields("DrugStoreName") & "</option>"
                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With

    ' Close dropdown recordset
    rstDropdown.Close
    Set rstDropdown = Nothing

    ' Construct SQL query for main data
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

    ' Initialize and open database connection for main data
    Set rstMain = CreateObject("ADODB.Recordset")
    rstMain.Open sql, conn, 3, 4

    ' Output dropdown
    Response.Write "<div>"
    Response.Write "        <label for='pharmacy' class='font-style'>Select Pharmacy:</label><br>"
    Response.Write "        <select id='pharmacy' name='pharmacy' multiple class='mult-select-tag'>"
    Response.Write dropdownOptions
    Response.Write "        </select>"
    Response.Write "</div>"
    
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

    If (periodStart <> "" And periodEnd <> "") Then
        Response.Write "<h2 class='font-style'>FROM: " & periodStart & " TO: " & periodEnd & "</h2>"
    Else
        Response.Write "<h2 class='font-style'>FROM: 2018-01-01 TO: 2018-01-31</h2>"
    End If

    ' Reset recordset to the first record
    rstMain.MoveFirst

    ' Display table
    Response.Write "<table class='mytable'>"
    Response.Write "<tr>"
    Response.Write "<th class='myth'>No.</th>"
    Response.Write "<th class='myth'>Pharmacy</th>"
    Response.Write "<th class='myth'>Stock Level</th>"
    Response.Write "<th class='myth'>Stock Date</th>"
    Response.Write "</tr>"

    count = 0

    With rstMain
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                count = count + 1
                Response.Write "<tr>"
                Response.Write "<td class='mytd' align='center'>" & count & "</td>"
                Response.Write "<td class='mytd'>" & .Fields("DrugStoreName") & "</td>"
                Response.Write "<td class='mytd' align='center'>" & .Fields("StockLevel") & "</td>"
                Response.Write "<td class='mytd' align='center'>" & .Fields("StockDate") & "</td>"
                Response.Write "</tr>"
                Response.Flush
                .MoveNext
            Loop
        End If
    End With

    Response.Write "</table>"

    ' Close main data recordset
    rstMain.Close
    Set rstMain = Nothing

    ' Output scripts
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
End Sub
