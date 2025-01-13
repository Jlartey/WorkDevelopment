Sub dispCatDischargeAmt()

    Dim sql, periodStart, periodEnd, datePeriod, count, selectedTreatTypes, idsArr, formattedIDs, id, dateArr
    Dim rstDropdown, rstMain
    Dim dropdownOptions, optionHTML

    ' Retrieve query parameters
    datePeriod = Trim(request.querystring("Dateperiod"))
    selectedTreatTypes = Trim(request.querystring("TreatTypeID"))

    ' Parse date period
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If

    ' Format selected drug store IDs
    If selectedTreatTypes <> "" Then
        idsArr = Split(selectedTreatTypes, ",")
        For Each id In idsArr
            formattedIDs = formattedIDs & "'" & Trim(id) & "',"
        Next
        ' Remove the trailing comma
        formattedIDs = Left(formattedIDs, Len(formattedIDs) - 1)
    End If

    ' Construct SQL query for dropdown options (all TreatTypes)
    sql = "select TreatTypeID, TreatTypeName from TreatType"

    ' Initialize and open database connection for dropdown options
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.open sql, conn, 3, 4

    ' Populate dropdown options
    dropdownOptions = ""

    With rstDropdown
        If .recordCount > 0 Then
            .movefirst
            Do Until .EOF
                optionHTML = "<option value='" & .fields("TreatTypeID") & "'>" & .fields("TreatTypeName") & "</option>"
                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With

    ' Close dropdown recordset
    rstDropdown.Close
    Set rstDropdown = Nothing

    ' Construct SQL query for main data
    sql = "With cateringCTE"
    sql = sql & " as"
    sql = sql & " ("
    sql = sql & " select "
    sql = sql & " sum(FinalAmt) [FinalAmt] ,"
    sql = sql & " convert(date, ConsultReviewDate) [TransactionDate]"
    sql = sql & " from TreatCharges"
    sql = sql & " join TreatType"
    sql = sql & " on TreatCharges.TreatTypeID = TreatType.TreatTypeID"

    If selectedTreatTypes <> "" Then
        sql = sql & " WHERE TreatType.TreatTypeName LIKE '%" & formattedIDs & "%'"
    Else
        sql = sql & " where TreatType.TreatTypeName  like '%CATERING%' "
    End If

    'sql = sql & " where TreatType.TreatTypeName  like '%CATERING%' "
    sql = sql & " and (convert(date, ConsultReviewDate) between '2017-10-01' and '2019-12-31')"
    sql = sql & " group by "
    sql = sql & " convert(date, ConsultReviewDate)"
    sql = sql & " ),"
    sql = sql & " dischargeCTE as("
    sql = sql & " select   "
    sql = sql & " convert(date, DischargeDate) [TransactionDate],"
    sql = sql & " sum(BedCharge * dbo.fn_NoOfDaysAdmitted(AdmissionDate, DischargeDate)) [FinalAmt] "
    sql = sql & " from Admission"
    sql = sql & " where BedCharge > 0"
    sql = sql & " and dbo.fn_NoOfDaysAdmitted(AdmissionDate, DischargeDate) is not null"

    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & " and (convert(date, DischargeDate) between '" & periodStart & "' and '" & periodEnd & "')"
    Else
        sql = sql & " and (convert(date, DischargeDate) between '2017-10-01' and '2019-12-31')"
    End If

    sql = sql & " group by  "
    sql = sql & " convert(date, DischargeDate)"
    sql = sql & " )"
    sql = sql & " "
    sql = sql & " select "
    sql = sql & " format(isnull(cateringCTE.FinalAmt, 0), 'N2') [CateringAmt], "
    sql = sql & " isnull(cateringCTE.TransactionDate, dischargeCTE.TransactionDate) [CateringDate], "
    sql = sql & " format(isnull(dischargeCTE.FinalAmt, 0), 'N2') [DischargeAmt], "
    sql = sql & " isnull(dischargeCTE.TransactionDate, cateringCTE.TransactionDate) [DischargeDate],"
    sql = sql & " format((isnull(cateringCTE.FinalAmt, 0) + isnull(dischargeCTE.FinalAmt, 0)), 'N2') [TotalAmt]"
    sql = sql & " from dischargeCTE"
    sql = sql & " full outer join cateringCTE"
    sql = sql & " on dischargeCTE.TransactionDate = cateringCTE.TransactionDate"
    sql = sql & " order by convert(date, dischargeCTE.TransactionDate) desc"


    ' Initialize and open database connection for main data
    Set rstMain = CreateObject("ADODB.Recordset")
    rstMain.open sql, conn, 3, 4

    ' Output dropdown
    Response.Write "<div>"
    Response.Write "        <label for='pharmacy' class='font-style'>Select TreatType:</label><br>"
    Response.Write "        <select id='treatType' name='pharmacy' multiple class='mult-select-tag'>"
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
    rstMain.movefirst

    ' Display table
    Response.Write "<table class='mytable'>"
    Response.Write "<tr>"
    Response.Write "<th class='myth'>No.</th>"
    Response.Write "<th class='myth'>Catering Amount</th>"
    Response.Write "<th class='myth'>Catering Date</th>"
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
                Response.Write "<td class='mytd' align='center'>" & .fields("CateringDate") & "</td>"
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

    ' Output scripts
    Response.Write "<script src='https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js'></script>"
    Response.Write "<script>"
        Response.Write "    new MultiSelectTag('treatType', {"
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