Styling
MultiSelectStyles
dispCatDischargeAmt

Sub dispCatDischargeAmt()
    Dim sql, periodStart, periodEnd, datePeriod, count, selectedTreatType, dateArr
    Dim rstDropdown, rstMain
    Dim dropdownOptions, optionHTML

    ' Retrieve query parameters
    datePeriod = Trim(request.querystring("Dateperiod"))
    selectedTreatType = Trim(request.querystring("TreatTypeID"))

    ' Parse date period
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If

    ' Construct SQL query for dropdown options (all TreatTypes)
    sql = "SELECT TreatTypeID, TreatTypeName FROM TreatType"

    ' Initialize and open database connection for dropdown options
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.open sql, conn, 3, 4

    ' Populate dropdown options
    dropdownOptions = ""

    With rstDropdown
        If .recordCount > 0 Then
            .movefirst
            Do Until .EOF
                optionHTML = "<option value='" & .fields("TreatTypeID") & "'"
                If .fields("TreatTypeID") = selectedTreatType Then
                    optionHTML = optionHTML & " selected"
                End If
                optionHTML = optionHTML & ">" & .fields("TreatTypeName") & "</option>"
                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With

    ' Close dropdown recordset
    rstDropdown.Close
    Set rstDropdown = Nothing

    ' Construct SQL query for main data
    sql = "WITH cateringCTE AS ("
    sql = sql & " SELECT SUM(FinalAmt) AS FinalAmt,"
    sql = sql & " CONVERT(date, ConsultReviewDate) AS TransactionDate"
    sql = sql & " FROM TreatCharges"
    sql = sql & " JOIN TreatType ON TreatCharges.TreatTypeID = TreatType.TreatTypeID"
    
    If selectedTreatType <> "" Then
        sql = sql & " WHERE TreatType.TreatTypeID = '" & selectedTreatType & "'"
    Else
        sql = sql & " WHERE TreatType.TreatTypeName = '%CATERING%'"
    End If
    
    If periodStart <> "" And periodEnd <> "" Then
        sql = sql & " AND (CONVERT(date, ConsultReviewDate) BETWEEN '" & periodStart & "' AND '" & periodEnd & "')"
    Else
        sql = sql & " AND (CONVERT(date, ConsultReviewDate) BETWEEN '2017-10-01' AND '2019-12-31')"
    End If
    'sql = sql & " AND (CONVERT(date, ConsultReviewDate) BETWEEN '2017-10-01' AND '2019-12-31')"
    sql = sql & " GROUP BY CONVERT(date, ConsultReviewDate)"
    sql = sql & " ), dischargeCTE AS ("
    sql = sql & " SELECT CONVERT(date, DischargeDate) AS TransactionDate,"
    sql = sql & " SUM(BedCharge * dbo.fn_NoOfDaysAdmitted(AdmissionDate, DischargeDate)) AS FinalAmt"
    sql = sql & " FROM Admission"
    sql = sql & " WHERE BedCharge > 0"
    sql = sql & " AND dbo.fn_NoOfDaysAdmitted(AdmissionDate, DischargeDate) IS NOT NULL"

    If periodStart <> "" And periodEnd <> "" Then
        sql = sql & " AND (CONVERT(date, DischargeDate) BETWEEN '" & periodStart & "' AND '" & periodEnd & "')"
    Else
        sql = sql & " AND (CONVERT(date, DischargeDate) BETWEEN '2017-10-01' AND '2019-12-31')"
    End If

    sql = sql & " GROUP BY CONVERT(date, DischargeDate)"
    sql = sql & " )"
    sql = sql & " SELECT"
    sql = sql & " FORMAT(ISNULL(cateringCTE.FinalAmt, 0), 'N2') AS CateringAmt,"
    sql = sql & " ISNULL(cateringCTE.TransactionDate, dischargeCTE.TransactionDate) AS CateringDate,"
    sql = sql & " FORMAT(ISNULL(dischargeCTE.FinalAmt, 0), 'N2') AS DischargeAmt,"
    sql = sql & " ISNULL(dischargeCTE.TransactionDate, cateringCTE.TransactionDate) AS DischargeDate,"
    sql = sql & " FORMAT((ISNULL(cateringCTE.FinalAmt, 0) + ISNULL(dischargeCTE.FinalAmt, 0)), 'N2') AS TotalAmt"
    sql = sql & " FROM dischargeCTE"
    sql = sql & " FULL OUTER JOIN cateringCTE ON dischargeCTE.TransactionDate = cateringCTE.TransactionDate"
    sql = sql & " ORDER BY CONVERT(date, dischargeCTE.TransactionDate) DESC"
    
    'response.write sql
    ' Initialize and open database connection for main data
    Set rstMain = CreateObject("ADODB.Recordset")
    rstMain.open sql, conn, 3, 4

    ' Output dropdown
    Response.Write "<div>"
    Response.Write "    <label for='treatType' class='font-style'>Select TreatType:</label><br>"
    Response.Write "    <select id='treatType' name='treatType' class='select-tag'>"
    Response.Write dropdownOptions
    Response.Write "    </select>"
    Response.Write "</div>"

    ' Output HTML Form for date selection
    Response.Write "<form id='dateForm'>"
    Response.Write "    <div class='container' style='display: flex; align-items: center; justify-content: center'>"
    Response.Write "        <div>"
    Response.Write "            <label for='from'>From</label>"
    Response.Write "            <input type='date' name='from' id='from'>"
    Response.Write "        </div>"
    Response.Write "        <div>"
    Response.Write "            <label for='to' style='margin-left: 10px'>To</label>"
    Response.Write "            <input type='date' name='to' id='to'>"
    Response.Write "        </div>"
    Response.Write "        <div>"
    Response.Write "            <button type='button' onclick='updateUrl()'>Show Data</button> <br />"
    Response.Write "        </div>"
    Response.Write "    </div>"
    Response.Write "</form>"

    If periodStart <> "" And periodEnd <> "" Then
        Response.Write "<h2 class='font-style'>FROM: " & periodStart & " TO: " & periodEnd & "</h2>"
    Else
        Response.Write "<h2 class='font-style'>FROM: 2018-01-01 TO: 2018-01-31</h2>"
    End If

    ' Reset recordset to the first record
    rstMain.movefirst

    ' Display table
    Response.Write "<table class='mytable'>"
    Response.Write "<tr>"
    Response.Write "<th class='myth'>No.</th>"
    Response.Write "<th class='myth'>Amount(Type)</th>"
    Response.Write "<th class='myth'>Discharge Amount</th>"
    Response.Write "<th class='myth'>Discharge Date</th>"
    Response.Write "<th class='myth'>Total Amount</th>"
    Response.Write "</tr>"

    count = 0

    With rstMain
        If .recordCount > 0 Then
            .movefirst
            Do While Not .EOF
                count = count + 1
                Response.Write "<tr>"
                Response.Write "<td class='mytd' align='center'>" & count & "</td>"
                Response.Write "<td class='mytd' align='center'>" & .fields("CateringAmt") & "</td>"
                Response.Write "<td class='mytd' align='center'>" & .fields("DischargeAmt") & "</td>"
                Response.Write "<td class='mytd' align='center'>" & .fields("DischargeDate") & "</td>"
                Response.Write "<td class='mytd' align='center'>" & .fields("TotalAmt") & "</td>"
                Response.Write "</tr>"
                Response.flush
                .MoveNext
            Loop
        End If
    End With

    Response.Write "</table>"

    ' Close main data recordset
    rstMain.Close
    Set rstMain = Nothing

    ' Output script for update URL function
    Response.Write "<script>"
    Response.Write "    function updateUrl() {"
    Response.Write "        const fromDate = document.getElementById('from').value;"
    Response.Write "        const toDate = document.getElementById('to').value;"
    Response.Write "        const treatType = Array.from(document.getElementById('treatType').selectedOptions).map(option => option.value).join(',');"
    Response.Write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp';"
    Response.Write "        const params = new URLSearchParams({"
    Response.Write "            PrintLayoutName: 'dispCatDischargeAmt',"
    Response.Write "            PositionForTableName: 'WorkingDay',"
    Response.Write "            WorkingDayID: '',"
    Response.Write "            Dateperiod: `${fromDate}||${toDate}`,"
    Response.Write "            TreatTypeID: treatType"
    Response.Write "        });"
    Response.Write "        const newUrl = baseUrl + '?' + params.toString();"
    Response.Write "        window.location.href = newUrl;"
    Response.Write "        console.log(newUrl);"
    Response.Write "    }"

    Response.Write "</script>"
End Sub


Sub Styling()
    Response.Write " <style>"
        Response.Write " .mytable {"
        Response.Write "     width: 65vw;"
        Response.Write "     font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
        Response.Write "     border-collapse: collapse;"
        Response.Write "     margin-top: 50px; "
        Response.Write " }"
        
        Response.Write " .container {"
        Response.Write "    display: flex"
        Response.Write "    margin-top: 50px !important;"
        Response.Write "    padding-top: 30px;"
        Response.Write " } "
        
        Response.Write " .myth, .mytd {"
        Response.Write "     border: 1px solid #ddd;"
        Response.Write "     padding: 10px;"
        Response.Write " }"
        
        Response.Write " .mytd {"
        Response.Write "     text-alig: 1px solid #ddd;"
        Response.Write "     padding: 8px;"
        Response.Write " }"
        
        Response.Write "  tr:nth-child(even) {"
        Response.Write "    background-color: #f9f9f9;"
        Response.Write " } "
        
        Response.Write " .myth {"
        Response.Write "     background-color: #c2c2c2;"
        Response.Write "     color: black;"
        Response.Write "     text-align: center; "
        Response.Write "     text-transform: uppercase; "
        Response.Write "     font-size: 18px;"
        Response.Write " }"
        
        Response.Write "  button {"
        Response.Write "     background-color: #0236c4;"
        Response.Write "     border-radius: 5px;"
        Response.Write "     border: none;"
        Response.Write "     margin-left: 50px;"
        Response.Write "     padding: 5px 20px;"
        Response.Write "     color: white;"
        Response.Write "     cursor: pointer;"
        Response.Write "  }"
        
        Response.Write "  #to, #from {"
        Response.Write "    padding: 5px;"
        Response.Write "    border-radius: 5px;"
        Response.Write "    cursor: pointer;"
        Response.Write "  }"
        
        Response.Write " .pagination {"
        Response.Write "    text-align: center;"
        Response.Write "    margin: 20px 0;"
        Response.Write " }"
        
        Response.Write " .pagination a {"
        Response.Write "    margin: 0 5px;"
        Response.Write "    padding: 10px 15px;"
        Response.Write "    background-color: #f1f1f1;"
        Response.Write "    border: 1px solid #ccc;"
        Response.Write "    text-decoration: none;"
        Response.Write "    color: #333;"
        Response.Write " }"
        
        Response.Write " .pagination a:hover {"
        Response.Write "    background-color: #ddd;"
        Response.Write " }"
        
        Response.Write " .font-style {"
        Response.Write "    font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif;"
        Response.Write " }"
        
        Response.Write " #pharmacy {"
        Response.Write "    padding-bottom: 10px;"
        Response.Write " }"
        Response.Write " </style>"
        
End Sub

Sub MultiSelectStyles()
     Response.Write "    <style>" & vbCrLf
    Response.Write "        .mult-select-tag {" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            width: 300px;" & vbCrLf
    Response.Write "            flex-direction: column;" & vbCrLf
    Response.Write "            align-items: center;" & vbCrLf
    Response.Write "            position: relative;" & vbCrLf
    Response.Write "            --tw-shadow: 0 1px 3px 0 rgb(0 0 0 / 0.1), 0 1px 2px -1px rgb(0 0 0 / 0.1);" & vbCrLf
    Response.Write "            --tw-shadow-color: 0 1px 3px 0 var(--tw-shadow-color), 0 1px 2px -1px var(--tw-shadow-color);" & vbCrLf
    Response.Write "            --border-color: rgb(218, 221, 224);" & vbCrLf
    Response.Write "            font-family: Verdana, sans-serif;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .wrapper {" & vbCrLf
    Response.Write "            width: 100%;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .body {" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            border: 1px solid var(--border-color);" & vbCrLf
    Response.Write "            background: #fff;" & vbCrLf
    Response.Write "            min-height: 2.15rem;" & vbCrLf
    Response.Write "            width: 100%;" & vbCrLf
    Response.Write "            min-width: 14rem;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .input-container {" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            flex-wrap: wrap;" & vbCrLf
    Response.Write "            flex: 1 1 auto;" & vbCrLf
    Response.Write "            padding: 0.1rem;" & vbCrLf
    Response.Write "            align-items: center;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .input-body {" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            width: 100%;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .input {" & vbCrLf
    Response.Write "            flex: 1;" & vbCrLf
    Response.Write "            background: 0 0;" & vbCrLf
    Response.Write "            border-radius: 0.25rem;" & vbCrLf
    Response.Write "            padding: 0.45rem;" & vbCrLf
    Response.Write "            margin: 10px;" & vbCrLf
    Response.Write "            color: #2d3748;" & vbCrLf
    Response.Write "            outline: 0;" & vbCrLf
    Response.Write "            border: 1px solid var(--border-color);" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .btn-container {" & vbCrLf
    Response.Write "            color: #e2ebf0;" & vbCrLf
    Response.Write "            padding: 0.5rem;" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            border-left: 1px solid var(--border-color);" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag button {" & vbCrLf
    Response.Write "            cursor: pointer;" & vbCrLf
    Response.Write "            width: 100%;" & vbCrLf
    Response.Write "            color: #718096;" & vbCrLf
    Response.Write "            outline: 0;" & vbCrLf
    Response.Write "            height: 100%;" & vbCrLf
    Response.Write "            border: none;" & vbCrLf
    Response.Write "            padding: 0;" & vbCrLf
    Response.Write "            background: 0 0;" & vbCrLf
    Response.Write "            background-image: none;" & vbCrLf
    Response.Write "            -webkit-appearance: none;" & vbCrLf
    Response.Write "            text-transform: none;" & vbCrLf
    Response.Write "            margin: 0;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag button:first-child {" & vbCrLf
    Response.Write "            width: 1rem;" & vbCrLf
    Response.Write "            height: 90%;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .drawer {" & vbCrLf
    Response.Write "            position: absolute;" & vbCrLf
    Response.Write "            background: #fff;" & vbCrLf
    Response.Write "            max-height: 15rem;" & vbCrLf
    Response.Write "            z-index: 40;" & vbCrLf
    Response.Write "            top: 98%;" & vbCrLf
    Response.Write "            width: 100%;" & vbCrLf
    Response.Write "            overflow-y: scroll;" & vbCrLf
    Response.Write "            border: 1px solid var(--border-color);" & vbCrLf
    Response.Write "            border-radius: 0.25rem;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag ul {" & vbCrLf
    Response.Write "            list-style-type: none;" & vbCrLf
    Response.Write "            padding: 0.5rem;" & vbCrLf
    Response.Write "            margin: 0;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag ul li {" & vbCrLf
    Response.Write "            padding: 0.5rem;" & vbCrLf
    Response.Write "            border-radius: 0.25rem;" & vbCrLf
    Response.Write "            cursor: pointer;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag ul li:hover {" & vbCrLf
    Response.Write "            background: rgb(243 244 246);" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .item-container {" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            justify-content: center;" & vbCrLf
    Response.Write "            align-items: center;" & vbCrLf
    Response.Write "            padding: 0.2rem 0.4rem;" & vbCrLf
    Response.Write "            margin: 0.2rem;" & vbCrLf
    Response.Write "            font-weight: 500;" & vbCrLf
    Response.Write "            border: 1px solid;" & vbCrLf
    Response.Write "            border-radius: 9999px;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .item-label {" & vbCrLf
    Response.Write "            max-width: 100%;" & vbCrLf
    Response.Write "            line-height: 1;" & vbCrLf
    Response.Write "            font-size: 0.75rem;" & vbCrLf
    Response.Write "            font-weight: 400;" & vbCrLf
    Response.Write "            flex: 0 1 auto;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .item-close-container {" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            flex: 1 1 auto;" & vbCrLf
    Response.Write "            flex-direction: row-reverse;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .item-close-svg {" & vbCrLf
    Response.Write "            width: 1rem;" & vbCrLf
    Response.Write "            margin-left: 0.5rem;" & vbCrLf
    Response.Write "            height: 1rem;" & vbCrLf
    Response.Write "            cursor: pointer;" & vbCrLf
    Response.Write "            border-radius: 9999px;" & vbCrLf
    Response.Write "            display: block;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .hidden {" & vbCrLf
    Response.Write "            display: none;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .shadow {" & vbCrLf
    Response.Write "            box-shadow: var(--tw-ring-offset-shadow, 0 0 #0000), var(--tw-ring-shadow, 0 0 #0000), var(--tw-shadow);" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .rounded {" & vbCrLf
    Response.Write "            border-radius: 0.375rem;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    </style>" & vbCrLf
End Sub
