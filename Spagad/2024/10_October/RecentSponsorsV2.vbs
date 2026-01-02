'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

displaySponsors
tableStyles

' Display the year selection dropdown
displayYearSelection

' Apply table styles and JavaScript for URL update functionality
tableStyles
pageScript

Sub displaySponsors()
    Dim rst, sql, periodStart, periodEnd, yearChoice, midYear, rowNum

    ' Retrieve year from query string
    yearChoice = Trim(request.QueryString("year"))
    If yearChoice = "" Then
        yearChoice = Year(Date) ' Default to current year if none selected
    End If
    
    ' Define date periods for first and second halves of the year
    periodStart = yearChoice & "-01-01"
    midYear = yearChoice & "-06-30"
    periodEnd = yearChoice & "-12-31"

    ' Set up database connection and query
    Set rst = CreateObject("ADODB.RecordSet")

    sql = "SELECT Sponsor.SponsorName, CONVERT(VARCHAR(20), MAX(InsuredPatient.EntryDate), 103) AS EntryDate "
    sql = sql & "FROM Sponsor "
    sql = sql & "JOIN InsuredPatient ON Sponsor.SponsorID = InsuredPatient.SponsorID "
    sql = sql & "JOIN SponsorType ON SponsorType.SponsorTypeID = InsuredPatient.SponsorTypeID "
    sql = sql & "WHERE InsuredPatient.SponsorTypeID = 'S004' "

    ' First half of the year
    sqlFirstHalf = sql & "AND InsuredPatient.EntryDate BETWEEN '" & periodStart & "' AND '" & midYear & "' "
    sqlFirstHalf = sqlFirstHalf & "GROUP BY Sponsor.SponsorName ORDER BY MAX(InsuredPatient.EntryDate) DESC"

    ' Second half of the year
    sqlSecondHalf = sql & "AND InsuredPatient.EntryDate BETWEEN '" & DateAdd("d", 1, midYear) & "' AND '" & periodEnd & "' "
    sqlSecondHalf = sqlSecondHalf & "GROUP BY Sponsor.SponsorName ORDER BY MAX(InsuredPatient.EntryDate) DESC"

    ' Display table for each half-year period
    DisplaySponsorTable rst, sqlFirstHalf, "Sponsors from Beginning to Mid-Year " & yearChoice
    DisplaySponsorTable rst, sqlSecondHalf, "Sponsors from Mid-Year to End of " & yearChoice
End Sub

' Helper Subroutine to Display Sponsor Table
Sub DisplaySponsorTable(rst, sql, headerTitle)
    With rst
        .open sql, conn, 3, 4
         
        If .RecordCount > 0 Then
            rowNum = 1
            response.write "<h1>" & headerTitle & "</h1>"
            response.write "<table class='mytable' width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Sponsor</th>"
            response.write "<th class='myth'>Entry Date</th>"
            response.write "</tr>"
            
            .MoveFirst
            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & rowNum & "</td>"
                response.write "<td class='mytd'>" & .fields("SponsorName") & "</td>"
                response.write "<td class='mytd'>" & .fields("EntryDate") & "</td>"
                response.write "</tr>"
                rowNum = rowNum + 1
                .MoveNext
            Loop
            response.write "</table>"
        Else
            response.write "<p>No records found for " & headerTitle & ".</p>"
        End If
        .Close
    End With  
End Sub

Sub tableStyles()
    response.write "<style>"
    response.write ".mytable { width: 50vw; border-collapse: collapse; margin: 20px 0; font-size: 16px; font-family: Arial, sans-serif; }"
    response.write ".mytable, .myth, .mytd { border: 1px solid #dddddd; }"
    response.write ".myth, .mytd { padding: 12px; text-align: left; }"
    response.write ".myth { background-color: #e0e0e0; color: #333; font-weight: bold; }"
    response.write ".mytr:nth-child(even) { background-color: #f9f9f9; }"
    response.write ".mytr:hover { background-color: #f1f1f1; }"
    response.write ".myth { text-transform: uppercase; }"
    response.write "h1 { font-size: 18px; color: #555; font-family: Arial, sans-serif; margin: 20px 0; }"
    response.write "</style>"
End Sub

Sub displayYearSelection()
    response.write "<form method='GET' action=''>"
    response.write "Select Year: <select name='year'>"
    Dim year, currentYear
    currentYear = Year(Date)
    For year = 2018 To currentYear
        response.write "<option value='" & year & "'"
        If year = CInt(request.QueryString("year")) Then response.write " selected"
        response.write ">" & year & "</option>"
    Next
    response.write "</select>"
    response.write "<button type='submit'>Process</button>"
    response.write "</form>"
End Sub

Sub pageScript()
    response.write "<script>"
    response.write "function updateURL() {"
    response.write "    const year = document.getElementsByName('year')[0].value;"
    response.write "    const url = new URL(window.location.href);"
    response.write "    url.searchParams.set('year', year);"
    response.write "    window.location.href = url;"
    response.write "}"
    response.write "</script>"
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>



'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

' Call subroutines in a logical order
Call displayYearSelection
Call pageScript
Call tableStyles
Call displaySponsors

'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Sub displaySponsors()
    Dim rst, sql, periodStart, periodEnd, selectedYear, midYearDate, rowNum
    selectedYear = Trim(request.QueryString("year"))
    If selectedYear = "" Then Exit Sub
    
    ' Set start and end periods for the selected year
    periodStart = selectedYear & "-01-01"
    midYearDate = selectedYear & "-06-30"
    periodEnd = selectedYear & "-12-31"

    ' First half of the year query
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
        .Open sql, conn, 3, 4
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

    ' Second half of the year query
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
        .Open sql, conn, 3, 4
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

Sub displayYearSelection()
    response.write "<form method='get' action=''>"
    response.write "    <label for='year'>Select Year: </label>"
    response.write "    <select name='year' id='year'>"
    
    Dim year
    For year = 2018 To Year(Date)
        response.write "        <option value='" & year & "'>" & year & "</option>"
    Next
    
    response.write "    </select>"
    response.write "    <button type='submit'>Process</button>"
    response.write "</form>"
End Sub

Sub pageScript()
    response.write "<script>"
    response.write "document.getElementById('year').value = new URLSearchParams(window.location.search).get('year') || '';"
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

'<<--END_CODE_SEGMENT_PRINTHEADER-->>




























































































































































































































































































































' <%
' ' Define connection object here if necessary
' Dim conn, currentYear
' currentYear = Year(Now())
' %>

' <!DOCTYPE html>
' <html lang="en">
' <head>
'     <title>Display Sponsors by Year</title>
'     <script type="text/javascript">
'         function processYearSelection() {
'             var selectedYear = document.getElementById("yearSelect").value;
'             if (selectedYear) {
'                 var url = window.location.pathname + "?year=" + selectedYear;
'                 window.location.href = url;
'             }
'         }
'     </script>
'     <style>
'         /* Table styling */
'         .mytable {
'             width: 50vw;
'             border-collapse: collapse;
'             margin: 20px 0;
'             font-size: 16px;
'             font-family: Arial, sans-serif;
'         }
'         .mytable, .myth, .mytd {
'             border: 1px solid #dddddd;
'         }
'         .myth, .mytd {
'             padding: 12px;
'             text-align: left;
'         }
'         .myth {
'             background-color: #e0e0e0;
'             color: #333;
'             font-weight: bold;
'             text-transform: uppercase;
'         }
'         .mytr:nth-child(even) {
'             background-color: #f9f9f9;
'         }
'         .mytr:hover {
'             background-color: #f1f1f1;
'         }
'         h2 {
'             font-size: 18px;
'             color: #555;
'             font-family: Arial, sans-serif;
'             margin: 20px 0;
'         }
'     </style>
' </head>
' <body>

'     <!-- Year Selection Dropdown -->
'     <label for="yearSelect">Select Year:</label>
'     <select id="yearSelect">
'         <option value="">-- Select Year --</option>
'         <% 
'         Dim year
'         For year = 2018 To currentYear
'             Response.Write("<option value='" & year & "'>" & year & "</option>")
'         Next
'         %>
'     </select>

'     <!-- Process Button -->
'     <button onclick="processYearSelection()">Process</button>

'     <!-- VBScript to call displaySponsors with the selected year -->
'     <%
'     Dim selectedYear
'     selectedYear = Request.QueryString("year")
    
'     If selectedYear <> "" Then
'         ' Define periods for the selected year
'         Dim periodStart1, periodEnd1, periodStart2, periodEnd2
'         periodStart1 = selectedYear & "-01-01"
'         periodEnd1 = selectedYear & "-06-30"
'         periodStart2 = selectedYear & "-07-01"
'         periodEnd2 = selectedYear & "-12-31"
        
'         ' Call displaySponsors for each period
'         Response.Write "<h2>Sponsors from January to June " & selectedYear & "</h2>"
'         Call displaySponsors(periodStart1, periodEnd1)
        
'         Response.Write "<h2>Sponsors from July to December " & selectedYear & "</h2>"
'         Call displaySponsors(periodStart2, periodEnd2)
'     End If

'     ' Subroutine to display sponsors within a specified period
'     Sub displaySponsors(periodStart, periodEnd)
'         Dim rst, sql, rowNum
'         Set rst = CreateObject("ADODB.RecordSet")
        
'         sql = "SELECT Sponsor.SponsorName, CONVERT(VARCHAR(20), MAX(InsuredPatient.EntryDate), 103) AS EntryDate "
'         sql = sql & "FROM Sponsor "
'         sql = sql & "JOIN InsuredPatient ON Sponsor.SponsorID = InsuredPatient.SponsorID "
'         sql = sql & "JOIN SponsorType ON SponsorType.SponsorTypeID = InsuredPatient.SponsorTypeID "
'         sql = sql & "WHERE InsuredPatient.SponsorTypeID = 'S004' "
'         sql = sql & "AND InsuredPatient.EntryDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
'         sql = sql & "GROUP BY Sponsor.SponsorName "
'         sql = sql & "ORDER BY MAX(InsuredPatient.EntryDate) DESC"
        
'         With rst
'             .open sql, conn, 3, 4
             
'             If .RecordCount > 0 Then
'                 rowNum = 1
'                 .MoveFirst
                
'                 response.write "<table class='mytable' cellspacing='0' cellpadding='2' border='1'>"
'                 response.write "<tr class='mytr'>"
'                     response.write "<th class='myth'> No. </th>"
'                     response.write "<th class='myth'> Sponsor </th>"
'                     response.write "<th class='myth'> Entry Date </th>"
'                 response.write "</tr>"
'                 Do While Not .EOF
'                     response.write "<tr class='mytr'>"
'                     response.write "<td class='mytd'>" & rowNum & "</td>"
'                     response.write "<td class='mytd'>" & .fields("SponsorName") & "</td>"
'                     response.write "<td class='mytd'>" & .fields("EntryDate") & "</td>"
'                     response.write "</tr>"
                    
'                     rowNum = rowNum + 1
'                     .MoveNext
'                 Loop
'                 response.write "</table>"
'             Else
'                 response.write "<p>No records found for this period.</p>"
'             End If
'             .Close
'         End With
'     End Sub
'     %>

' </body>
' </html>
