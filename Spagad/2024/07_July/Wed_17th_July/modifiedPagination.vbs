Sub dispPatientVisitationsAlternative02()
    Dim sql, periodStart, periodEnd, datePeriod, count, selectedWardIDs, idsArr, formattedIDs, id, dateArr
    Dim rstDropdown, rstMain
    Dim dropdownOptions, optionHTML

    ' Retrieve query parameters
    datePeriod = Trim(Request.QueryString("Dateperiod"))
    selectedWardIDs = Trim(Request.QueryString("WardID"))
    pageSize = 10 ' Default page size

    ' Parse date period
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If

    ' Format selected ward IDs
    If selectedWardIDs <> "" Then
        idsArr = Split(selectedWardIDs, ",")
        For Each id In idsArr
            formattedIDs = formattedIDs & "'" & Trim(id) & "',"
        Next
        ' Remove the trailing comma
        formattedIDs = Left(formattedIDs, Len(formattedIDs) - 1)
    End If

    ' Construct SQL query for dropdown options (all wards)
    sql = "SELECT WardID, Wardname FROM Ward"

    ' Initialize and open database connection for dropdown options
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.Open sql, conn, 3, 4

    ' Populate dropdown options
    dropdownOptions = ""

    With rstDropdown
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                optionHTML = "<option value='" & .Fields("WardID") & "'>" & .Fields("Wardname") & "</option>"
                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With

    ' Close dropdown recordset
    rstDropdown.Close
    Set rstDropdown = Nothing

    ' Construct SQL query for main data
    sql = "SELECT "
    sql = sql & "Patient.PatientID, "
    sql = sql & "Patient.Patientname, "
    sql = sql & "VisitType.VisitTypeName, "
    sql = sql & "CONVERT(varchar(20), Visitation.VisitDate, 103) AS VisitDate, "
    sql = sql & "AdmissionStatus.AdmissionStatusName, "
    sql = sql & "CONVERT(varchar(20), Admission.AdmissionDate, 103) AS AdmissionDate, "
    sql = sql & "CONVERT(varchar(20), Admission.DischargeDate, 103) AS DischargeDate, "
    sql = sql & "Ward.WardName, "
    sql = sql & "Bed.BedName "
    sql = sql & "FROM "
    sql = sql & "Patient "
    sql = sql & "INNER JOIN "
    sql = sql & "Visitation ON Visitation.PatientID = Patient.PatientID "
    sql = sql & "INNER JOIN "
    sql = sql & "VisitType ON Visitation.VisitTypeID = VisitType.VisitTypeID "
    sql = sql & "INNER JOIN "
    sql = sql & "Admission ON Admission.VisitationID = Visitation.VisitationID "
    sql = sql & "INNER JOIN "
    sql = sql & "AdmissionStatus ON AdmissionStatus.AdmissionStatusID = Admission.AdmissionStatusID "
    sql = sql & "INNER JOIN "
    sql = sql & "Ward ON Admission.WardID = Ward.WardID "
    sql = sql & "INNER JOIN "
    sql = sql & "Bed ON Admission.BedID = Bed.BedID "
    sql = sql & "WHERE "

    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "CONVERT(date, Visitation.VisitDate) BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    Else
        sql = sql & "CONVERT(date, Visitation.VisitDate) BETWEEN '2018-01-01' AND '2018-04-01' "
    End If

    If selectedWardIDs <> "" Then
        sql = sql & "AND Ward.WardID IN (" & formattedIDs & ")"
    Else
        sql = sql & "AND Ward.WardID IS NOT NULL "
    End If

    sql = sql & "ORDER BY "
    sql = sql & "CONVERT(date, Admission.AdmissionDate) DESC"

    ' Initialize and open database connection for main data (assuming it's already established)
    ' If the connection is managed externally, conn should already be defined and opened

    Set rstMain = CreateObject("ADODB.Recordset")
    rstMain.Open sql, conn, 3, 4

    ' Pagination variables
    Dim pageNum, totalPages
    pageNum = 1 ' Default page number

    ' Check if page number is specified
    If Request.QueryString("page") <> "" Then
        pageNum = CInt(Request.QueryString("page"))
    End If

    ' Move to the starting record of the current page
    rstMain.Move (pageNum - 1) * pageSize

    ' Output dropdown
    Response.Write "<div>"
    Response.Write "        <label for='ward' class='font-style'>Select Ward:</label><br>"
    Response.Write "        <select id='ward' name='ward' multiple class='mult-select-tag'>"
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
        Response.Write "<h2 class='font-style'>SHOWING DATA FROM: " & periodStart & " TO: " & periodEnd & "</h2>"
    Else
        Response.Write "<h2 class='font-style'>SHOWING DATA FROM: 2018-01-01 TO: 2018-01-31</h2>"
    End If

    ' Display table
    Response.Write "<table class='mytable'>"
    Response.Write "<tr>"
    Response.Write "<th class='myth'>No.</th>"
    Response.Write "<th class='myth'>Patient Name</th>"
    Response.Write "<th class='myth'>Visit Type</th>"
    Response.Write "<th class='myth'>Visit Date</th>"
    Response.Write "<th class='myth'>Admission Status</th>"
    Response.Write "<th class='myth'>Admission Date</th>"
    Response.Write "<th class='myth'>Discharge Date</th>"
    Response.Write "<th class='myth'>Ward</th>"
    Response.Write "<th class='myth'>Bed</th>"
    Response.Write "</tr>"

    count = 0

    ' Loop through the records for the current page
    Do While Not rstMain.EOF And count < pageSize
        count = count + 1
        Response.Write "<tr>"
        Response.Write "<td class='mytd' align='center'>" & count & "</td>"
        Response.Write "<td class='mytd'>" & rstMain.Fields("Patientname") & "</td>"
        Response.Write "<td class='mytd' align='center'>" & rstMain.Fields("VisitTypeName") & "</td>"
        Response.Write "<td class='mytd' align='center'>" & rstMain.Fields("VisitDate") & "</td>"
        Response.Write "<td class='mytd' align='center'>" & rstMain.Fields("AdmissionStatusName") & "</td>"
        Response.Write "<td class='mytd' align='center'>" & rstMain.Fields("AdmissionDate") & "</td>"
        Response.Write "<td class='mytd' align='center'>" & rstMain.Fields("DischargeDate") & "</td>"
        Response.Write "<td class='mytd' align='center'>" & rstMain.Fields("WardName") & "</td>"
        Response.Write "<td class='mytd' align='center'>" & rstMain.Fields("BedName") & "</td>"
        Response.Write "</tr>"
        rstMain.MoveNext
    Loop

    Response.Write "</table>"

    ' Calculate total pages
    totalPages = Int((rstMain.RecordCount + pageSize - 1) / pageSize)

    ' Close main data recordset
    rstMain.Close
    Set rstMain = Nothing

    ' Output pagination links
    Response.Write "<div class='pagination'>"
    If pageNum > 1 Then
        Response.Write "<a href='?page=1&Dateperiod=" & Server.UrlEncode(datePeriod) & "&WardID=" & Server.UrlEncode(selectedWardIDs) & "&pageSize=" & pageSize & "'>First</a>"
        Response.Write "<a href='?page=" & (pageNum - 1) & "&Dateperiod=" & Server.UrlEncode(datePeriod) & "&WardID=" & Server.UrlEncode(selectedWardIDs) & "&pageSize=" & pageSize & "'>Previous</a>"
    End If

    Response.Write "Page " & pageNum & " of " & totalPages

    If pageNum < totalPages Then
        Response.Write "<a href='?page=" & (pageNum + 1) & "&Dateperiod=" & Server.UrlEncode(datePeriod) & "&WardID=" & Server.UrlEncode(selectedWardIDs) & "&pageSize=" & pageSize & "'>Next</a>"
        Response.Write "<a href='?page=" & totalPages & "&Dateperiod=" & Server.UrlEncode(datePeriod) & "&WardID=" & Server.UrlEncode(selectedWardIDs) & "&pageSize=" & pageSize & "'>Last</a>"
    End If

    Response.Write "</div>"

    ' Output scripts
    Response.Write "<script src='https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js'></script>"
    Response.Write "<script>"
    Response.Write "function updateUrl() {"
    Response.Write "  var fromDate = document.getElementById('from').value;"
    Response.Write "  var toDate = document.getElementById('to').value;"
    Response.Write "  var urlParams = new URLSearchParams(window.location.search);"
    Response.Write "  urlParams.set('Dateperiod', fromDate + '||' + toDate);"
    Response.Write "  urlParams.delete('page');"
    Response.Write "  window.location.href = '?' + urlParams.toString();"
    Response.Write "}"
    Response.Write "</script>"

    ' Load more button using JavaScript
    Response.Write "<script>"
    Response.Write "function loadMore() {"
    Response.Write "  var nextPage = " & pageNum + 1 & ";"
    Response.Write "  var urlParams = new URLSearchParams(window.location.search);"
    Response.Write "  urlParams.set('page', nextPage);"
    Response.Write "  window.location.href = '?' + urlParams.toString();"
    Response.Write "}"
    Response.Write "</script>"

    Response.Write "<div style='text-align: center; margin-top: 20px;'>"
    If pageNum < totalPages Then
        Response.Write "<button onclick='loadMore()'>Load More</button>"
    End If
    Response.Write "</div>"

End Sub
