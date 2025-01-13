'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim periodStart, periodEnd, dateArr, datePeriod, selectedVisitStatusID

'Retrieve query parameters
datePeriod = Trim(Request.QueryString("Dateperiod"))
selectedVisitStatusID = Trim(Request.QueryString("visitStatusID"))

'Parse date period
If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    periodStart = dateArr(0)
    periodEnd = dateArr(1)
End If

If selectedVisitStatusID <> "" Then
    idsArr = Split(selectedVisitStatusID, ",")
    For Each id In idsArr
        formattedIDs = formattedIDs & "'" & Trim(id) & "',"
    Next
    formattedIDs = Left(formattedIDs, Len(formattedIDs) - 1)
End If

Styling
MultiSelectStyles

Response.Write "<!DOCTYPE html>"
Response.Write "<html lang='en'>"
Response.Write "<head>"
Response.Write "<meta charset='UTF-8'>"
Response.Write "<meta name='viewport' content='width=device-width, initial-scale=1.0'>"
Response.Write "<title>Visitation Analysis</title>"

Response.Write "<script src='https://cdn.plot.ly/plotly-2.32.0.min.js'></script>"
'Response.Write "<script src='https://cdn.plot.ly/plotly-latest.min.js'></script>"

Response.Write "    <link href=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css"" rel=""stylesheet"""
Response.Write "        integrity=""sha384-9ndCyUaIbzAi2FUVXJi0CjmCapSmO7SnpJef0486qhLnuZ2cdeRhO02iuK6FUUVM"" crossorigin=""anonymous"">"
Response.Write "    <script src=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"""
Response.Write "        integrity=""sha384-geWF76RCwLtnZ8qwWowPQNguL3RmwHVBC9FhGdlKrxdiJJigb/j/68SIy3Te4Bkz"""
Response.Write "        crossorigin=""anonymous""></script>"
Response.Write " <link href=""https://cdn.datatables.net/v/bs5/jq-3.6.0/jszip-2.5.0/dt-1.13.5/af-2.6.0/b-2.4.0/b-colvis-2.4.0/b-html5-2.4.0/b-print-2.4.0/cr-1.7.0/date-1.5.0/fc-4.3.0/fh-3.4.0/kt-2.10.0/r-2.5.0/rg-1.4.0/rr-1.4.0/sc-2.2.0/sb-1.5.0/sp-2.2.0/sl-1.7.0/sr-1.3.0/datatables.min.css"" rel=""stylesheet""/>"
Response.Write " <script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js""></script>"
Response.Write " <script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js""></script>"
Response.Write " <script src=""https://cdn.datatables.net/v/bs5/jq-3.6.0/jszip-2.5.0/dt-1.13.5/af-2.6.0/b-2.4.0/b-colvis-2.4.0/b-html5-2.4.0/b-print-2.4.0/cr-1.7.0/date-1.5.0/fc-4.3.0/fh-3.4.0/kt-2.10.0/r-2.5.0/rg-1.4.0/rr-1.4.0/sc-2.2.0/sb-1.5.0/sp-2.2.0/sl-1.7.0/sr-1.3.0/datatables.min.js""></script>"
'response.write "<script src='https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js'></script>"

Response.Write "<style>"
Response.Write "  .chart-container {"
Response.Write "    display: flex;"
Response.Write "    justify-content: center;"
Response.Write "  }"
Response.Write "  .chart {"
Response.Write "    flex: 1;"
Response.Write "    margin: 10px;"
Response.Write "    width: 80%;"
Response.Write "  }"
Response.Write "  .tab-header {"
Response.Write "    display: flex;"
Response.Write "    justify-content: center;"
Response.Write "    background-color: #007bff;"
Response.Write "    border: 1px solid #ddd;"
Response.Write "    border-radius: 5px;"
Response.Write "  }"
Response.Write "  .tab-button {"
Response.Write "    flex: 1;"
Response.Write "    padding: 10px;"
Response.Write "    text-align: center;"
Response.Write "    cursor: pointer;"
Response.Write "    font-weight: bold;"
Response.Write "    color: #fff;"
Response.Write "    border-right: 1px solid #ddd;"
Response.Write "  }"
Response.Write "  .tab-button:last-child {"
Response.Write "    border-right: none;"
Response.Write "  }"
Response.Write "  .tab-button.active {"
Response.Write "    background-color: #0056b3;"
Response.Write "  }"
Response.Write "  .tab-content {"
Response.Write "    display: none;"
Response.Write "    padding: 20px;"
Response.Write "    border: 1px solid #ddd;"
Response.Write "    border-radius: 5px;"
Response.Write "    background-color: #f9f9f9;"
Response.Write "    margin-top: 10px;"
Response.Write "  }"
Response.Write "  .tab-content.active {"
Response.Write "    display: block;"
Response.Write "  }"
Response.Write "</style>"


Response.Write "<div id='myDiv'><!-- Plotly chart will be drawn inside this DIV --></div>"
Response.Write "</head>"
Response.Write "<body>"


' Construct SQL query for dropdown options (all pharmacies)
    sql = "select VisitStatusId, VisitStatusName from VisitStatus"

    ' Initialize and open database connection for dropdown options
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.Open sql, conn, 3, 4

    ' Populate dropdown options
    dropdownOptions = ""

    With rstDropdown
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                optionHTML = "<option value='" & .Fields("VisitStatusId") & "'>" & .Fields("VisitStatusName") & "</option>"
                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With

    ' Close dropdown recordset
    rstDropdown.Close
    Set rstDropdown = Nothing

' Output dropdown
Response.Write "<div class='header'>"
    Response.Write "<div>"
        Response.Write "        <label for='visitStatus' class='font-style'>Select VisitStatus:</label><br>"
        Response.Write "        <select id='visitStatus' name='visitStatus' multiple class='mult-select-tag'>"
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
        Response.Write "        </div>"
        Response.Write "    </div> "
        Response.Write "   </form>"
    Response.Write "</div>"
  
    If (periodStart <> "" And periodEnd <> "") Then
        Response.Write "<h2 class='font-style'>SHOWING DATA FROM: " & periodStart & " TO: " & periodEnd & "</h2>"
    Else
        Response.Write "<h2 class='font-style'>SHOWING DATA FROM: 2018-01-01 TO: 2018-01-31</h2>"
    End If

Response.Write "<div id='yearlyTab' class='tab-content active'>"
Response.Write "  <div class='chart-container'>"
Response.Write "    <div id='yearlyChartDiv' class='chart'></div>"
Response.Write "  </div>"

' table

  Response.Write "      <table style=""width:100%"" id=""yearlyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
  Response.Write "      <thead class=""table-dark"">"
  Response.Write "        <tr>"
  Response.Write "             <th>No.</th>"
  Response.Write "             <th>Visit Status</th>"
  Response.Write "             <th>Year</th>"
  Response.Write "             <th>Number Of Visits</th>"
  Response.Write "             <th>Previous Visits</th>"
  Response.Write "             <th>Difference</th>"
  Response.Write "             <th>Percent Change</th>"
  Response.Write "             <th>VisitStatusName Contribution (%)</th>"
  Response.Write "             <th>Overall Visits Contributioon (%)</th>"
  Response.Write "        </tr>"
  Response.Write "       </thead>"
  Response.Write "    </table>"
Response.Write "</div>"

Response.Write "</body>"
Response.Write "</html>"

dispVisitationAnalysis

Sub dispVisitationAnalysis()
    Dim sql, count
    Dim dropdownOptions, optionHTML

    ' Construct SQL query for main data
    sql = "WITH selectCTE " & vbCrLf
    sql = sql & "AS " & vbCrLf
    sql = sql & "( " & vbCrLf
    sql = sql & "    SELECT " & vbCrLf
    sql = sql & "        --VisitStatus.VisitStatusID, " & vbCrLf
    sql = sql & "        VisitStatus.VisitStatusName, " & vbCrLf
    sql = sql & "        DATENAME(year, VisitDate) AS VisitYear, " & vbCrLf
    sql = sql & "        COUNT(*) AS NumberOfVisits, " & vbCrLf
    sql = sql & "        LAG(COUNT(*)) OVER(PARTITION BY VisitStatus.VisitStatusName " & vbCrLf
    sql = sql & "        ORDER BY VisitStatus.VisitStatusName, DATENAME(year, VisitDate)) AS [PrevVisits], " & vbCrLf
    sql = sql & "        COUNT(*) - LAG(COUNT(*)) OVER(PARTITION BY VisitStatus.VisitStatusName " & vbCrLf
    sql = sql & "        ORDER BY VisitStatus.VisitStatusName, DATENAME(year, VisitDate)) AS [Diff] " & vbCrLf
    sql = sql & "    FROM Visitation " & vbCrLf
    sql = sql & "    JOIN VisitStatus ON Visitation.VisitStatusID = VisitStatus.VisitStatusID " & vbCrLf
    sql = sql & "    WHERE " & vbCrLf

    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "     CONVERT(date, Visitation.VisitDate) BETWEEN '" & periodStart & "' AND '" & periodEnd & "' " & vbCrLf
    Else
        sql = sql & "     CONVERT(date, Visitation.VisitDate) BETWEEN '2018-01-01' AND '2022-12-31' " & vbCrLf
    End If

    If selectedVisitStatusID <> "" Then
        sql = sql & " and VisitStatus.VisitStatusID IN (" & formattedIDs & ") "
    Else
        sql = sql & " and VisitStatus.VisitStatusID IS NOT NULL "
    End If

    sql = sql & "    GROUP BY " & vbCrLf
    sql = sql & "        VisitStatus.VisitStatusID, " & vbCrLf
    sql = sql & "        VisitStatus.VisitStatusName, " & vbCrLf
    sql = sql & "        DATENAME(year, VisitDate) " & vbCrLf
    sql = sql & ") " & vbCrLf
    sql = sql & "--ORDER BY " & vbCrLf
    sql = sql & "--    VisitStatus.VisitStatusName, VisitYear " & vbCrLf
    sql = sql & "SELECT " & vbCrLf
    sql = sql & "    VisitStatusName, VisitYear, NumberOfVisits, " & vbCrLf
    sql = sql & "    PrevVisits, Diff, " & vbCrLf
    sql = sql & "    Diff * 100.00 / NumberOfVisits AS [PercentChange], " & vbCrLf
    sql = sql & "    NumberOfVisits * 100.00 / SUM(NumberOfVisits) " & vbCrLf
    sql = sql & "        OVER (PARTITION BY VisitStatusName) AS [PercentContToVisitStatusName], " & vbCrLf
    sql = sql & "    NumberOfVisits * 100.00 / SUM(NumberOfVisits) OVER () AS [PercentContToOverallVisits] " & vbCrLf
    sql = sql & "FROM selectCTE"

    'Initialize and open database connection for main data
    Set rstMain = CreateObject("ADODB.Recordset")
    rstMain.Open sql, conn, 3, 4

    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    If rstMain.RecordCount > 0 Then
        rstMain.MoveFirst
        Do While Not rstMain.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """VisitStatusName"":""" & rstMain.Fields("VisitStatusName").Value & ""","
            jsonData = jsonData & """VisitYear"":""" & rstMain.Fields("VisitYear").Value & ""","
            jsonData = jsonData & """NumberOfVisits"":""" & rstMain.Fields("NumberOfVisits").Value & ""","
            jsonData = jsonData & """PrevVisits"":""" & rstMain.Fields("PrevVisits").Value & ""","
            jsonData = jsonData & """Diff"":""" & rstMain.Fields("Diff").Value & ""","
            jsonData = jsonData & """PercentChange"":""" & rstMain.Fields("PercentChange").Value & ""","
            jsonData = jsonData & """PercentContToVisitStatusName"":""" & rstMain.Fields("PercentContToVisitStatusName").Value & ""","
            jsonData = jsonData & """PercentContToOverallVisits"":""" & rstMain.Fields("PercentContToOverallVisits").Value & """"
            jsonData = jsonData & "},"
             rstMain.MoveNext
            counter = counter + 1
        Loop
        jsonData = Left(jsonData, Len(jsonData) - 1) ' Remove the trailing comma
    End If

    jsonData = jsonData & "]}"

    rstMain.Close
    Set rstMain = Nothing
    ' Send the data to the client-side
    
    Response.Write "<script>"
    Response.Write "var dbDataYearly = " & jsonData & ";"
    Response.Write "document.addEventListener('DOMContentLoaded', function() {"
    Response.Write "    var visitData = dbDataYearly.data;"
    
    ' Extract data for plotting
    Response.Write "    var years = [...new Set(visitData.map(item => item.VisitYear))];"
    Response.Write "    var visitCounts = years.map(year => {"
    Response.Write "        return visitData.filter(item => item.VisitYear == year).reduce((sum, item) => sum + parseInt(item.NumberOfVisits), 0);"
    Response.Write "    });"

    ' Create a trace for the bar chart
    Response.Write "    var trace = {"
    Response.Write "        x: years,"
    Response.Write "        y: visitCounts,"
    Response.Write "        type: 'bar',"
    Response.Write "        text: visitCounts.map(count => 'Number of Visits: ' + count),"
    Response.Write "        hovertemplate: '%{text}',"
    Response.Write "    };"

    ' Define the layout for the bar chart
    Response.Write "    var layout = {"
    Response.Write "        title: 'Number of Visits per Year',"
    Response.Write "        xaxis: { title: 'Year' },"
    Response.Write "        yaxis: { title: 'Number of Visits' }"
    Response.Write "    };"

    ' Plot the bar chart
    Response.Write "    Plotly.newPlot('myDiv', [trace], layout);"
    Response.Write "});"
    Response.Write "</script>"


    Response.Write "<script src='https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js'></script>"
    Response.Write "<script>"
        Response.Write "    new MultiSelectTag('visitStatus', {"
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
        Response.Write "        const visitStatusArr = Array.from(document.getElementById('visitStatus').selectedOptions).map(option => option.value).join(',');"
        Response.Write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp';"
        Response.Write "        const params = new URLSearchParams({"
        Response.Write "            PrintLayoutName: 'dispVisitationAnalysis',"
        Response.Write "            PositionForTableName: 'WorkingDay',"
        Response.Write "            WorkingDayID: '',"
        Response.Write "            Dateperiod: `${fromDate}||${toDate}`,"
        Response.Write "            VisitStatusID: visitStatusArr"
        Response.Write "        });"
        Response.Write "        const newUrl = baseUrl + '?' + params.toString();"
        Response.Write "        window.location.href = newUrl;"
        Response.Write "        console.log(newUrl);"
        Response.Write "    }"

'        Response.Write "var data = ["
'        Response.Write "{"
'        Response.Write "   x: jsonData.visitYear,"
'        Response.Write "   y: jsonData.NumberOfVisits,"
'        Response.Write "   type: 'bar'"
'        Response.Write " }"
'        Response.Write "];"
'
'        Response.Write " Plotly.newPlot('myDiv', data);"



    Response.Write "</script>"

    ' DataTable Initialization
    Response.Write "<script>"
    Response.Write "var dbDataYearly = " & jsonData & ";"
    Response.Write "    new DataTable('#yearlyTable', {"
    Response.Write "        data: dbDataYearly.data,"
    Response.Write "        columns: ["
    Response.Write "            { data: 'counter' },"
    Response.Write "            { data: 'VisitStatusName' },"
    Response.Write "            { data: 'VisitYear' },"
    Response.Write "            { data: 'NumberOfVisits' },"
    Response.Write "            { data: 'PrevVisits' },"
    Response.Write "            { data: 'Diff' },"
    Response.Write "            { data: 'PercentChange' },"
    Response.Write "            { data: 'PercentContToVisitStatusName' },"
    Response.Write "            { data: 'PercentContToOverallVisits' }"
    Response.Write "        ],"


    Response.Write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    Response.Write "        dom: 'lBfrtip',"
    Response.Write "        buttons: ["
    Response.Write "            {"
    Response.Write "                extend: 'csv',"
    Response.Write "                text: 'CSV',"
    Response.Write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'excel',"
    Response.Write "                text: 'EXCEL',"
    Response.Write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'pdf',"
    Response.Write "                text: 'PDF',"
    Response.Write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'print',"
    Response.Write "                text: 'PRINT',"
    Response.Write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            }"
    Response.Write "        ]"
    Response.Write "    });"


'    Response.Write " function plotData() {"
'        Response.Write "// Extract data from dbDataYearly"
'        Response.Write "         const years = dbDataYearly.data.map(item => item.VisitYear);"
'        Response.Write "         const visits = dbDataYearly.data.map(item => item.NumberOfVisits);"
'
'        Response.Write "           // Create a trace for the bar chart"
'        Response.Write "          const trace = {"
'        Response.Write "              x: years,"
'        Response.Write "             y: visits,"
'        Response.Write "             type: 'bar'"
'        Response.Write "         };"
'
'        Response.Write "         // Define the layout for the bar chart"
'        Response.Write "         const layout = {"
'        Response.Write "         title: 'Number of Visits per Year',"
'        Response.Write "               xaxis: { title: 'Year' },"
'        Response.Write "               yaxis: { title: 'Number of Visits' }"
'        Response.Write "           };"
'
'        Response.Write "          // Plot the bar chart"
'        Response.Write "          Plotly.newPlot('myDiv', [trace], layout);"
'        Response.Write "       }"
'
'        Response.Write "        // Ensure dbDataYearly is defined and plot the data"
'        Response.Write "        if (typeof dbDataYearly !== 'undefined') {"
'        Response.Write "            plotData();"
'        Response.Write "       } else {"
'        Response.Write "          console.error('Data not available');"
'        Response.Write "    }"
'
'    Response.Write "</script>"
End Sub

Sub dispVisitationAnalysis01()
    Dim sql, count
    Dim dropdownOptions, optionHTML

    ' Construct SQL query for main data
    sql = "WITH selectCTE " & vbCrLf
    sql = sql & "AS " & vbCrLf
    sql = sql & "( " & vbCrLf
    sql = sql & "    SELECT " & vbCrLf
    sql = sql & "        VisitStatus.VisitStatusName, " & vbCrLf
    sql = sql & "        DATENAME(year, VisitDate) AS VisitYear, " & vbCrLf
    sql = sql & "        COUNT(*) AS NumberOfVisits, " & vbCrLf
    sql = sql & "        LAG(COUNT(*)) OVER(PARTITION BY VisitStatus.VisitStatusName " & vbCrLf
    sql = sql & "        ORDER BY VisitStatus.VisitStatusName, DATENAME(year, VisitDate)) AS [PrevVisits], " & vbCrLf
    sql = sql & "        COUNT(*) - LAG(COUNT(*)) OVER(PARTITION BY VisitStatus.VisitStatusName " & vbCrLf
    sql = sql & "        ORDER BY VisitStatus.VisitStatusName, DATENAME(year, VisitDate)) AS [Diff] " & vbCrLf
    sql = sql & "    FROM Visitation " & vbCrLf
    sql = sql & "    JOIN VisitStatus ON Visitation.VisitStatusID = VisitStatus.VisitStatusID " & vbCrLf
    sql = sql & "    WHERE " & vbCrLf

    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "     CONVERT(date, Visitation.VisitDate) BETWEEN '" & periodStart & "' AND '" & periodEnd & "' " & vbCrLf
    Else
        sql = sql & "     CONVERT(date, Visitation.VisitDate) BETWEEN '2018-01-01' AND '2022-12-31' " & vbCrLf
    End If

    If selectedVisitStatusID <> "" Then
        sql = sql & " and VisitStatus.VisitStatusID IN (" & formattedIDs & ") "
    Else
        sql = sql & " and VisitStatus.VisitStatusID IS NOT NULL "
    End If

    sql = sql & "    GROUP BY " & vbCrLf
    sql = sql & "        VisitStatus.VisitStatusID, " & vbCrLf
    sql = sql & "        VisitStatus.VisitStatusName, " & vbCrLf
    sql = sql & "        DATENAME(year, VisitDate) " & vbCrLf
    sql = sql & ") " & vbCrLf
    sql = sql & "SELECT " & vbCrLf
    sql = sql & "    VisitStatusName, VisitYear, NumberOfVisits, " & vbCrLf
    sql = sql & "    PrevVisits, Diff, " & vbCrLf
    sql = sql & "    Diff * 100.00 / NumberOfVisits AS [PercentChange], " & vbCrLf
    sql = sql & "    NumberOfVisits * 100.00 / SUM(NumberOfVisits) " & vbCrLf
    sql = sql & "        OVER (PARTITION BY VisitStatusName) AS [PercentContToVisitStatusName], " & vbCrLf
    sql = sql & "    NumberOfVisits * 100.00 / SUM(NumberOfVisits) OVER () AS [PercentContToOverallVisits] " & vbCrLf
    sql = sql & "FROM selectCTE"

    ' Initialize and open database connection for main data
    Set rstMain = CreateObject("ADODB.Recordset")
    rstMain.Open sql, conn, 3, 4

    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    If rstMain.RecordCount > 0 Then
        rstMain.MoveFirst
        Do While Not rstMain.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """VisitStatusName"":""" & rstMain.Fields("VisitStatusName").Value & ""","
            jsonData = jsonData & """VisitYear"":""" & rstMain.Fields("VisitYear").Value & ""","
            jsonData = jsonData & """NumberOfVisits"":""" & rstMain.Fields("NumberOfVisits").Value & ""","
            jsonData = jsonData & """PrevVisits"":""" & rstMain.Fields("PrevVisits").Value & ""","
            jsonData = jsonData & """Diff"":""" & rstMain.Fields("Diff").Value & ""","
            jsonData = jsonData & """PercentChange"":""" & rstMain.Fields("PercentChange").Value & ""","
            jsonData = jsonData & """PercentContToVisitStatusName"":""" & rstMain.Fields("PercentContToVisitStatusName").Value & ""","
            jsonData = jsonData & """PercentContToOverallVisits"":""" & rstMain.Fields("PercentContToOverallVisits").Value & """"
            jsonData = jsonData & "},"
            rstMain.MoveNext
            counter = counter + 1
        Loop
        jsonData = Left(jsonData, Len(jsonData) - 1) ' Remove the trailing comma
    End If

    jsonData = jsonData & "]}"

    rstMain.Close
    Set rstMain = Nothing

    ' Send the data to the client-side
    Response.Write "<script>"
    Response.Write "var dbDataYearly = " & jsonData & ";"
    Response.Write "document.addEventListener('DOMContentLoaded', function() {"
    Response.Write "    var visitData = dbDataYearly.data;"
    
    ' Extract data for plotting
    Response.Write "    var years = [...new Set(visitData.map(item => item.VisitYear))];"
    Response.Write "    var visitCounts = years.map(year => {"
    Response.Write "        return visitData.filter(item => item.VisitYear == year).reduce((sum, item) => sum + parseInt(item.NumberOfVisits), 0);"
    Response.Write "    });"

    ' Create a trace for the bar chart
    Response.Write "    var trace = {"
    Response.Write "        x: years,"
    Response.Write "        y: visitCounts,"
    Response.Write "        type: 'bar',"
    Response.Write "        text: visitCounts.map(count => 'Number of Visits: ' + count),"
    Response.Write "        hovertemplate: '%{text}',"
    Response.Write "    };"

    ' Define the layout for the bar chart
    Response.Write "    var layout = {"
    Response.Write "        title: 'Number of Visits per Year',"
    Response.Write "        xaxis: { title: 'Year' },"
    Response.Write "        yaxis: { title: 'Number of Visits' }"
    Response.Write "    };"

    ' Plot the bar chart
    Response.Write "    Plotly.newPlot('myDiv', [trace], layout);"
    Response.Write "});"
    Response.Write "</script>"

    ' DataTable Initialization
    Response.Write "<script>"
    Response.Write "var dbDataYearly = " & jsonData & ";"
    Response.Write "    new DataTable('#yearlyTable', {"
    Response.Write "        data: dbDataYearly.data,"
    Response.Write "        columns: ["
    Response.Write "            { data: 'counter' },"
    Response.Write "            { data: 'VisitStatusName' },"
    Response.Write "            { data: 'VisitYear' },"
    Response.Write "            { data: 'NumberOfVisits' },"
    Response.Write "            { data: 'PrevVisits' },"
    Response.Write "            { data: 'Diff' },"
    Response.Write "            { data: 'PercentChange' },"
    Response.Write "            { data: 'PercentContToVisitStatusName' },"
    Response.Write "            { data: 'PercentContToOverallVisits' }"
    Response.Write "        ],"
    Response.Write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    Response.Write "        dom: 'lBfrtip',"
    Response.Write "        buttons: ["
    Response.Write "            {"
    Response.Write "                extend: 'csv',"
    Response.Write "                text: 'CSV',"
    Response.Write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'excel',"
    Response.Write "                text: 'EXCEL',"
    Response.Write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'pdf',"
    Response.Write "                text: 'PDF',"
    Response.Write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'print',"
    Response.Write "                text: 'PRINT',"
    Response.Write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            'colvis'"
    Response.Write "        ]"
    Response.Write "    });"
    Response.Write "</script>"

End Sub


Sub Styling()
    Response.Write " <style>"
        Response.Write " .mytable {"
        Response.Write "     width: 95vw;"
        Response.Write "     font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
        Response.Write "     border-collapse: collapse;"
        Response.Write "     margin-top: 50px; "
        Response.Write "     border-radius: 10px;"
        Response.Write " }"
        
        Response.Write " .header {"
        Response.Write "    display: flex;"
        Response.Write "    justify-content: center;"
        Response.Write "    align-items: center;"
        Response.Write " } "
        
        Response.Write " .font-style {"
        Response.Write "    text-align: center;"
        Response.Write " } "
        
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
        Response.Write "    background-color: rgba(249, 249, 249, 6);"
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


'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
