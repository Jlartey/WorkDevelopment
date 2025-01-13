Sub dispCatDischargeAmt()
    Dim sql, periodStart, periodEnd, datePeriod, count, selectedTreatType, dateArr
    Dim rstDropdown, rstMain
    Dim dropdownOptions, optionHTML

    ' Retrieve query parameters
    datePeriod = Trim(Request.QueryString("Dateperiod"))
    selectedTreatType = Trim(Request.QueryString("TreatTypeID"))

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
    rstDropdown.Open sql, conn, 3, 4

    ' Populate dropdown options
    dropdownOptions = ""

    With rstDropdown
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                optionHTML = "<option value='" & .Fields("TreatTypeID") & "'"
                If .Fields("TreatTypeID") = selectedTreatType Then
                    optionHTML = optionHTML & " selected"
                End If
                optionHTML = optionHTML & ">" & .Fields("TreatTypeName") & "</option>"
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
        sql = sql & " WHERE TreatType.TreatTypeName LIKE '%CATERING%'"
    End If

    sql = sql & " AND (CONVERT(date, ConsultReviewDate) BETWEEN '2017-10-01' AND '2019-12-31')"
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

    ' Initialize and open database connection for main data
    Set rstMain = CreateObject("ADODB.Recordset")
    rstMain.Open sql, conn, 3, 4

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
    rstMain.MoveFirst

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
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                count = count + 1
                Response.Write "<tr>"
                Response.Write "<td class='mytd' align='center'>" & count & "</td>"
                Response.Write "<td class='mytd' align='center'>" & .Fields("CateringAmt") & "</td>"
                Response.Write "<td class='mytd' align='center'>" & .Fields("CateringDate") & "</td>"
                Response.Write "<td class='mytd' align='center'>" & .Fields("DischargeAmt") & "</td>"
                Response.Write "<td class='mytd' align='center'>" & .Fields("DischargeDate") & "</td>"
                Response.Write "<td class='mytd' align='center'>" & .Fields("TotalAmt") & "</td>"
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

    ' Output script for update URL function
    Response.Write "<script>"
    Response.Write "    function updateUrl() {"
    Response.Write "        var fromDate = document.getElementById('from').value;"
    Response.Write "        var toDate = document.getElementById('to').value;"
    Response.Write "        var treatType = document.getElementById('treatType').value;"
    Response.Write "        var newUrl = '?Dateperiod=' + fromDate + '||' + toDate + '&TreatTypeID=' + treatType;"
    Response.Write "        window.location.href = newUrl;"
    Response.Write "    }"
    Response.Write "</script>"
End Sub
