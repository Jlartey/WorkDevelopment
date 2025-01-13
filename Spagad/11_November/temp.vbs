displayYearSelection
pageScript
tableStyles
displaySponsors

Sub displayYearSelection()
    response.write "<label for='year'>Select Year: </label>"
    response.write "<select name='year' id='year'>"
    
    Dim year
    For year = 2018 To Year(Date)
        response.write "    <option value='" & year & "'>" & year & "</option>"
    Next
    
    response.write "</select>"
    response.write "<button type='button' onclick='processSelection()'>Process</button>"
End Sub

Sub displaySponsors()
    Dim rst, sql, periodStart, periodEnd, selectedYear, midYearDate, rowNum
    selectedYear = Trim(request.querystring("year"))
    If selectedYear = "" Then Exit Sub

    ' Define date ranges based on the selected year
    periodStart = selectedYear & "-01-01"
    midYearDate = selectedYear & "-06-30"
    periodEnd = selectedYear & "-12-31"

    ' Display records for the first half of the year
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "SELECT Sponsor.SponsorName, CONVERT(VARCHAR(20), MAX(InsuredPatient.EntryDate), 103) AS EntryDate " & _
          "FROM Sponsor " & _
          "JOIN InsuredPatient ON Sponsor.SponsorID = InsuredPatient.SponsorID " & _
          "JOIN SponsorType ON SponsorType.SponsorTypeID = InsuredPatient.SponsorTypeID " & _
          "WHERE InsuredPatient.SponsorTypeID = 'S004' " & _
          "AND InsuredPatient.EntryDate BETWEEN '" & periodStart & "' AND '" & midYearDate & "' " & _
          "GROUP BY Sponsor.SponsorName " & _
          "ORDER BY MAX(InsuredPatient.EntryDate) DESC"

    With rst
        .open sql, conn, 3, 4
        If .RecordCount > 0 Then
            rowNum = 1
            response.write "<h1>Records from Jan - June, " & selectedYear & "</h1>"
            response.write "<table class='mytable' width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr class='mytr'><th class='myth'>No.</th><th class='myth'>Sponsor</th><th class='myth'>Entry Date</th></tr>"
            Do While Not .EOF
                response.write "<tr class='mytr'><td class='mytd'>" & rowNum & "</td><td class='mytd'>" & .fields("SponsorName") & "</td><td class='mytd'>" & .fields("EntryDate") & "</td></tr>"
                rowNum = rowNum + 1
                .MoveNext
            Loop
            response.write "</table>"
        Else
            response.write "No records found for Jan - June, " & selectedYear
        End If
        .Close
    End With

    ' Display records for the second half of the year
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "SELECT Sponsor.SponsorName, CONVERT(VARCHAR(20), MAX(InsuredPatient.EntryDate), 103) AS EntryDate " & _
          "FROM Sponsor " & _
          "JOIN InsuredPatient ON Sponsor.SponsorID = InsuredPatient.SponsorID " & _
          "JOIN SponsorType ON SponsorType.SponsorTypeID = InsuredPatient.SponsorTypeID " & _
          "WHERE InsuredPatient.SponsorTypeID = 'S004' " & _
          "AND InsuredPatient.EntryDate BETWEEN '" & midYearDate & "' AND '" & periodEnd & "' " & _
          "GROUP BY Sponsor.SponsorName " & _
          "ORDER BY MAX(InsuredPatient.EntryDate) DESC"

    With rst
        .open sql, conn, 3, 4
        If .RecordCount > 0 Then
            rowNum = 1
            response.write "<h1>Records from July - Dec, " & selectedYear & "</h1>"
            response.write "<table class='mytable' width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr class='mytr'><th class='myth'>No.</th><th class='myth'>Sponsor</th><th class='myth'>Entry Date</th></tr>"
            Do While Not .EOF
                response.write "<tr class='mytr'><td class='mytd'>" & rowNum & "</td><td class='mytd'>" & .fields("SponsorName") & "</td><td class='mytd'>" & .fields("EntryDate") & "</td></tr>"
                rowNum = rowNum + 1
                .MoveNext
            Loop
            response.write "</table>"
        Else
            response.write "No records found for July - Dec, " & selectedYear
        End If
        .Close
    End With
End Sub

Sub pageScript()
    response.write "<script>"
    response.write "function processSelection() {"
    response.write "    var selectedYear = document.getElementById('year').value;"
    response.write "    if (selectedYear) {"
    response.write "        var baseUrl = 'http://192.168.5.11/rmchms01/wpgPrtPrintLayoutAll.asp?PrintLayoutName=RecentSponsors&PositionForTableName=WorkingDay&WorkingDayID=';"
    response.write "        window.location.href = baseUrl + '&year=' + selectedYear;"
    response.write "    } else {"
    response.write "        alert('Please select a year');"
    response.write "    }"
    response.write "}"
    response.write "</script>"
End Sub

Sub tableStyles()
    response.write "<style>"
    response.write ".mytable {width: 50vw; border-collapse: collapse; margin: 20px 0; font-size: 16px; font-family: Arial, sans-serif;}"
    response.write ".mytable, .myth, .mytd {border: 1px solid #dddddd;}"
    response.write ".myth, .mytd {padding: 12px; text-align: left;}"
    response.write ".myth {background-color: #e0e0e0; color: #333; font-weight: bold; text-transform: uppercase;}"
    response.write ".mytr:nth-child(even) {background-color: #f9f9f9;}"
    response.write ".mytr:hover {background-color: #f1f1f1;}"
    response.write "h1 {font-size: 18px; color: #555; font-family: Arial, sans-serif; margin: 20px 0;}"
    response.write "</style>"
End Sub
