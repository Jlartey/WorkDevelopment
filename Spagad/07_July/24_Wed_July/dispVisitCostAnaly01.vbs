'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Response.Write "Hello Joe"

Dim periodStart, periodEnd, dateArr, datePeriod

'Retrieve query parameters
datePeriod = Trim(Request.QueryString("Dateperiod"))

'Parse date period
If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    periodStart = dateArr(0)
    periodEnd = dateArr(1)
End If

Styling

Response.Write "<!DOCTYPE html>"
Response.Write "<html lang='en'>"
Response.Write "<head>"
Response.Write "<meta charset='UTF-8'>"
Response.Write "<meta name='viewport' content='width=device-width, initial-scale=1.0'>"
Response.Write "<title>Visit Cost Analysis</title>"


Response.Write "<script src='https://cdn.plot.ly/plotly-2.34.0.min.js'></script>"

Response.Write "    <link href=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css"" rel=""stylesheet"""
Response.Write "        integrity=""sha384-9ndCyUaIbzAi2FUVXJi0CjmCapSmO7SnpJef0486qhLnuZ2cdeRhO02iuK6FUUVM"" crossorigin=""anonymous"">"
Response.Write "    <script src=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"""
Response.Write "        integrity=""sha384-geWF76RCwLtnZ8qwWowPQNguL3RmwHVBC9FhGdlKrxdiJJigb/j/68SIy3Te4Bkz"""
Response.Write "        crossorigin=""anonymous""></script>"
Response.Write " <link href=""https://cdn.datatables.net/v/bs5/jq-3.6.0/jszip-2.5.0/dt-1.13.5/af-2.6.0/b-2.4.0/b-colvis-2.4.0/b-html5-2.4.0/b-print-2.4.0/cr-1.7.0/date-1.5.0/fc-4.3.0/fh-3.4.0/kt-2.10.0/r-2.5.0/rg-1.4.0/rr-1.4.0/sc-2.2.0/sb-1.5.0/sp-2.2.0/sl-1.7.0/sr-1.3.0/datatables.min.css"" rel=""stylesheet""/>"
Response.Write " <script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js""></script>"
Response.Write " <script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js""></script>"
Response.Write " <script src=""https://cdn.datatables.net/v/bs5/jq-3.6.0/jszip-2.5.0/dt-1.13.5/af-2.6.0/b-2.4.0/b-colvis-2.4.0/b-html5-2.4.0/b-print-2.4.0/cr-1.7.0/date-1.5.0/fc-4.3.0/fh-3.4.0/kt-2.10.0/r-2.5.0/rg-1.4.0/rr-1.4.0/sc-2.2.0/sb-1.5.0/sp-2.2.0/sl-1.7.0/sr-1.3.0/datatables.min.js""></script>"
Response.Write "<script src='https://cdn.plot.ly/plotly-latest.min.js'></script>"

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

Response.Write "</head>"

Response.Write "<body>"

' Output dropdown
Response.Write "<div class='header'>"
  
    'Output HTML Form for date selection
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

Response.Write "<div id='visitCostTab' class='tab-content active'>"
Response.Write "  <div class='chart-container'>"
Response.Write "    <div id='visitCostChartDiv' class='chart'></div>"
Response.Write "  </div>"

' table

  Response.Write "      <table style=""width:100%"" id=""visitCostTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
  Response.Write "      <thead class=""table-dark"">"
  Response.Write "        <tr>"
  Response.Write "             <th>No.</th>"
  Response.Write "             <th>Visit Year</th>"
  Response.Write "             <th>Visitation CostF</th>"
  Response.Write "             <th>Visitation Cost</th>"
  Response.Write "             <th>Previous Cost</th>"
  Response.Write "             <th>Difference</th>"
  Response.Write "             <th>Percent Change</th>"
  Response.Write "             <th>Yearly Contribution (%)</th>"
  Response.Write "        </tr>"
  Response.Write "       </thead>"
  Response.Write "    </table>"
Response.Write "</div>"

'Response.Write " it works"
Response.Write "</body>"
Response.Write "</html>"

dispVisitCostAnalysis
'MedicalOutcomeYearly
Sub dispVisitCostAnalysis()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")

    ' Construct SQL query for main data
    sql = "WITH VisitCostCTE AS ("
    sql = sql & "SELECT "
    sql = sql & "    DATENAME(YEAR, Visitation.VisitDate) VisitYear, "
    sql = sql & "    SUM(Visitation.VisitCost) VisitationCost, "
    sql = sql & "    LAG(SUM(Visitation.VisitCost)) OVER(ORDER BY DATENAME(YEAR, Visitation.VisitDate)) AS [PrevCost], "
    sql = sql & "    SUM(Visitation.VisitCost) - (LAG(SUM(Visitation.VisitCost)) OVER(ORDER BY DATENAME(YEAR, Visitation.VisitDate))) AS [Diff] "
    sql = sql & "FROM Visitation "
    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "WHERE Visitation.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    Else
        sql = sql & "WHERE Visitation.VisitDate BETWEEN '2018-01-01' AND '2022-12-31' "
    End If
    sql = sql & "GROUP BY DATENAME(YEAR, Visitation.VisitDate) "
    sql = sql & "), "
    sql = sql & "GroupCTE AS ("
    sql = sql & "SELECT "
    sql = sql & "    VisitYear, "
    sql = sql & "    VisitationCost, "
    sql = sql & "    PrevCost, "
    sql = sql & "    Diff, "
    sql = sql & "    (Diff * 100.00) / VisitationCost AS [PercentageChange], "
    sql = sql & "    VisitationCost * 100.00 / SUM(VisitationCost) OVER() AS ContPercentage "
    sql = sql & "FROM VisitCostCTE) "
    sql = sql & "SELECT "
    sql = sql & "    VisitYear, "
    sql = sql & "    FORMAT(VisitationCost, 'N2') AS VisitationCostF, "
    sql = sql & "    VisitationCost, "
    sql = sql & "    FORMAT(PrevCost, 'N2') AS PrevCost, "
    sql = sql & "    FORMAT(Diff, 'N2') AS Diff, "
    sql = sql & "    FORMAT(PercentageChange, 'N2') AS PercentageChange, "
    sql = sql & "    FORMAT(ContPercentage, 'N2') AS ContPercentage "
    sql = sql & "FROM GroupCTE"

    rst.Open sql, conn, 3, 4

    ' Generate JSON data
    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    ' Check if the recordset has any records
    If Not rst.EOF Then
        rst.MoveFirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """VisitYear"":""" & rst.Fields("VisitYear").Value & ""","
            jsonData = jsonData & """VisitationCost"":""" & rst.Fields("VisitationCost").Value & ""","
            jsonData = jsonData & """PrevCost"":""" & rst.Fields("PrevCost").Value & ""","
            jsonData = jsonData & """Diff"":""" & rst.Fields("Diff").Value & ""","
            jsonData = jsonData & """PercentageChange"":""" & rst.Fields("PercentageChange").Value & ""","
            jsonData = jsonData & """ContPercentage"":""" & rst.Fields("ContPercentage").Value & """"
            jsonData = jsonData & "},"
            rst.MoveNext
            counter = counter + 1
        Loop
        jsonData = Left(jsonData, Len(jsonData) - 1) ' Remove the trailing comma
    End If

    jsonData = jsonData & "]}"

    rst.Close
    Set rst = Nothing
    
    Response.Write "<script>"
        Response.Write "    function updateUrl() {"
        Response.Write "        const fromDate = document.getElementById('from').value;"
        Response.Write "        const toDate = document.getElementById('to').value;"
        Response.Write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp';"
        Response.Write "        const params = new URLSearchParams({"
        Response.Write "            PrintLayoutName: 'dispVisitCostAnalysis',"
        Response.Write "            PositionForTableName: 'WorkingDay',"
        Response.Write "            WorkingDayID: '',"
        Response.Write "            Dateperiod: `${fromDate}||${toDate}`"
        Response.Write "        });"
        Response.Write "        const newUrl = baseUrl + '?' + params.toString();"
        Response.Write "        window.location.href = newUrl;"
        Response.Write "        console.log(newUrl);"
        Response.Write "    }"
    Response.Write "</script>"
    
    
    ' Send the data to the client-side
    Response.Write "<script>"
    Response.Write "var dbDataYearly = " & jsonData & ";"
    Response.Write "document.addEventListener('DOMContentLoaded', function() {"
    Response.Write "    var yearlyData = dbDataYearly.data;"

    ' Defining colors
    Response.Write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"

    ' Define the chart data
    Response.Write "    var data = yearlyData.reduce((acc, row) => {"
    Response.Write "        acc.x.push(row.VisitYear);"
    Response.Write "        acc.y.push(row.VisitationCost);"
    Response.Write "        return acc;"
    Response.Write "    }, { x: [], y: [], type: 'bar', marker: { color: colors[0] } });"

    ' Define the layout for the chart
    Response.Write "    var barLayout = {"
    Response.Write "        title: 'Visitation Cost Analysis Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & "',"
    Response.Write "        xaxis: { title: 'Visit Year' },"
    Response.Write "        yaxis: { title: 'Visitation Cost' },"
    Response.Write "        height: 600, width: window.innerWidth * 0.95,"
    Response.Write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    Response.Write "        text: yearlyData.map(pair => 'Visit Year: ' + pair.VisitYear + ' <br>Visitation Cost: ' + pair.VisitationCost + '<br>Prev Cost: ' + pair.PrevCost + '<br>Diff: ' + pair.Diff + '<br>Percentage Change: ' + pair.PercentageChange + '<br>Yearly Contribution: ' + pair.ContPercentage),"
    Response.Write "        textposition: 'auto',"
    Response.Write "        texttemplate: '%{y}',"
    Response.Write "        hovertemplate: '%{text}',"
    Response.Write "        legend: { orientation: 'h', y: -0.2, x: 0.5, xanchor: 'center', yanchor: 'top' }"
    Response.Write "    };"

    ' Create the bar chart
    Response.Write "    Plotly.newPlot('visitCostChartDiv', [data], barLayout);"
    Response.Write "});"
    Response.Write "</script>"

    ' DataTable Initialization
    Response.Write "<script>"
    Response.Write "    new DataTable('#visitCostTrendsTable', {"
    Response.Write "        data: dbDataYearly.data,"
    Response.Write "        columns: ["
    Response.Write "            { data: 'counter' },"
    Response.Write "            { data: 'VisitYear' },"
    Response.Write "            { data: 'VisitationCost' },"
    Response.Write "            { data: 'PrevCost' },"
    Response.Write "            { data: 'Diff' },"
    Response.Write "            { data: 'PercentageChange' },"
    Response.Write "            { data: 'ContPercentage' }"
    Response.Write "        ],"
    Response.Write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    Response.Write "        dom: 'lBfrtip',"
    Response.Write "        buttons: ["
    Response.Write "            {"
    Response.Write "                extend: 'csv',"
    Response.Write "                text: 'CSV',"
    Response.Write "                title: '" & brnchName & " Visit Cost Trends From: " & periodStart & " To: " & periodEnd & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'excel',"
    Response.Write "                text: 'EXCEL',"
    Response.Write "                title: '" & brnchName & " Visit Cost Trends From: " & periodStart & " To: " & periodEnd & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'pdf',"
    Response.Write "                text: 'PDF',"
    Response.Write "                title: '" & brnchName & " Visit Cost Trends From: " & periodStart & " To: " & periodEnd & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'print',"
    Response.Write "                text: 'PRINT',"
    Response.Write "                title: '" & brnchName & " Visit Cost Trends From: " & periodStart & " To: " & periodEnd & "'"
    Response.Write "            }"
    Response.Write "        ]"
    Response.Write "    });"
    Response.Write "</script>"
End Sub


Sub dispVisitCostAnalysis()
    Dim sql, count
    
    ' Construct SQL query for main data
    sql = "WITH VisitCostCTE AS ("
    sql = sql & "SELECT "
    sql = sql & "    DATENAME(YEAR, Visitation.VisitDate) VisitYear, "
    sql = sql & "    SUM(Visitation.VisitCost) VisitationCost, "
    sql = sql & "    LAG(SUM(Visitation.VisitCost)) OVER(ORDER BY DATENAME(YEAR, Visitation.VisitDate)) AS [PrevCost], "
    sql = sql & "    SUM(Visitation.VisitCost) - (LAG(SUM(Visitation.VisitCost)) OVER(ORDER BY DATENAME(YEAR, Visitation.VisitDate))) AS [Diff] "
    sql = sql & "FROM Visitation "
    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "WHERE Visitation.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    Else
        sql = sql & "WHERE Visitation.VisitDate BETWEEN '2018-01-01' AND '2026-12-31' "
    End If
    'sql = sql & "WHERE Visitation.VisitDate BETWEEN '2018-01-01' AND '2026-12-31' "
    sql = sql & "GROUP BY DATENAME(YEAR, Visitation.VisitDate) "
    sql = sql & "), "
    sql = sql & "GroupCTE AS ("
    sql = sql & "SELECT "
    sql = sql & "    VisitYear, "
    sql = sql & "    VisitationCost, "
    sql = sql & "    PrevCost, "
    sql = sql & "    Diff, "
    sql = sql & "    (Diff * 100.00) / VisitationCost AS [PercentageChange], "
    sql = sql & "    VisitationCost * 100.00 / SUM(VisitationCost) OVER() AS ContPercentage "
    sql = sql & "FROM VisitCostCTE) "
    sql = sql & "SELECT "
    sql = sql & "    VisitYear, "
    sql = sql & "    FORMAT(VisitationCost, 'N2') AS VisitationCostF, "
    sql = sql & "    VisitationCost, "
    sql = sql & "    FORMAT(PrevCost, 'N2') AS PrevCost, "
    sql = sql & "    FORMAT(Diff, 'N2') AS Diff, "
    sql = sql & "    FORMAT(PercentageChange, 'N2') AS PercentageChange, "
    sql = sql & "    FORMAT(ContPercentage, 'N2') AS ContPercentage "
    sql = sql & "FROM GroupCTE"
    
    'response.write sql
    
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
            jsonData = jsonData & """VisitYear"":""" & rstMain.Fields("VisitYear").Value & ""","
            jsonData = jsonData & """VisitationCostF"":""" & rstMain.Fields("VisitationCostF").Value & ""","
            jsonData = jsonData & """VisitationCost"":""" & rstMain.Fields("VisitationCost").Value & ""","
            jsonData = jsonData & """PrevCost"":""" & rstMain.Fields("PrevCost").Value & ""","
            jsonData = jsonData & """Diff"":""" & rstMain.Fields("Diff").Value & ""","
            jsonData = jsonData & """PercentageChange"":""" & rstMain.Fields("PercentageChange").Value & ""","
            jsonData = jsonData & """ContPercentage"":""" & rstMain.Fields("ContPercentage").Value & """"
            jsonData = jsonData & "},"
             rstMain.MoveNext
            counter = counter + 1
        Loop
        jsonData = Left(jsonData, Len(jsonData) - 1) ' Remove the trailing comma
    End If
    
    jsonData = jsonData & "]}"

    rstMain.Close
    Set rstMain = Nothing
    
    Response.Write "<script>"
        Response.Write "    function updateUrl() {"
        Response.Write "        const fromDate = document.getElementById('from').value;"
        Response.Write "        const toDate = document.getElementById('to').value;"
        Response.Write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp';"
        Response.Write "        const params = new URLSearchParams({"
        Response.Write "            PrintLayoutName: 'dispVisitCostAnalysis',"
        Response.Write "            PositionForTableName: 'WorkingDay',"
        Response.Write "            WorkingDayID: '',"
        Response.Write "            Dateperiod: `${fromDate}||${toDate}`"
        Response.Write "        });"
        Response.Write "        const newUrl = baseUrl + '?' + params.toString();"
        Response.Write "        window.location.href = newUrl;"
        Response.Write "        console.log(newUrl);"
        Response.Write "    }"
    Response.Write "</script>"
    
    ' DataTable Initialization
    Response.Write "<script>"
    Response.Write "var dbDataYearly = " & jsonData & ";"
    Response.Write "    new DataTable('#visitCostTable', {"
    Response.Write "        data: dbDataYearly.data,"
    Response.Write "        columns: ["
    Response.Write "            { data: 'counter' },"
    Response.Write "            { data: 'VisitYear' },"
    Response.Write "            { data: 'VisitationCostF' },"
    Response.Write "            { data: 'VisitationCost' },"
    Response.Write "            { data: 'PrevCost' },"
    Response.Write "            { data: 'Diff' },"
    Response.Write "            { data: 'PercentageChange' },"
    Response.Write "            { data: 'ContPercentage' },"
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
    Response.Write "</script>"
    
'    response.write "<script>"
'        response.write "var dbDataYearly = " & jsonData & ";"
'        response.write "document.addEventListener('DOMContentLoaded', function() {"
'        response.write "var yearlyData = dbDataYearly.data;"
'
'        ' Defining a color palette
'        response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
'
'        ' Define the chart data
'        response.write "    var data = yearlyData.reduce((acc, row) => {"
'        response.write "        acc.x.push(row.VisitYear);"
'        response.write "        acc.y.push(row.VisitationCost);"
'        response.write "        return acc;"
'        response.write "    }, { x: [], y: [], type: 'bar', marker: { color: colors[0] } });"
'
'        ' Define the layout for the chart
'        response.write "    var barLayout = {"
'        'Response.Write "        title: 'Visitation Cost Analysis Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & "',"
'        response.write "        xaxis: { title: 'Visit Year' },"
'        response.write "        yaxis: { title: 'Visitation Cost' },"
'        response.write "        height: 600, width: window.innerWidth * 0.95,"
'        response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
'        response.write "        text: yearlyData.map(pair => 'Visit Year: ' + pair.VisitYear + ' <br>Visitation Cost: ' + pair.VisitationCost + '<br>Prev Cost: ' + pair.PrevCost + '<br>Diff: ' + pair.Diff + '<br>Percentage Change: ' + pair.PercentageChange + '<br>Yearly Contribution: ' + pair.ContPercentage),"
'        response.write "        textposition: 'auto',"
'        response.write "        texttemplate: '%{y}',"
'        response.write "        hovertemplate: '%{text}',"
'        response.write "        legend: { orientation: 'h', y: -0.2, x: 0.5, xanchor: 'center', yanchor: 'top' }"
'        response.write "    };"
'
'        ' Create the bar chart
'        response.write "    Plotly.newPlot('visitCostChartDiv', [data], barLayout);"
'        response.write "});"
'response.write "</script>"

' Send the data to the client-side
    Response.Write "<script>"
    Response.Write "var dbDataYearly = " & jsonData & ";"
    Response.Write "document.addEventListener('DOMContentLoaded', function() {"
    Response.Write "    var yearlyData = dbDataYearly.data;"

    ' Defining colors
    Response.Write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"

    ' Define the chart data
    Response.Write "    var data = yearlyData.reduce((acc, row) => {"
    Response.Write "        acc.x.push(row.VisitYear);"
    Response.Write "        acc.y.push(row.VisitationCost);"
    Response.Write "        return acc;"
    Response.Write "    }, { x: [], y: [], type: 'bar', marker: { color: colors[0] } });"

    ' Define the layout for the chart
    Response.Write "    var barLayout = {"
    Response.Write "        title: 'Visitation Cost Analysis Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & "',"
    Response.Write "        xaxis: { title: 'Visit Year' },"
    Response.Write "        yaxis: { title: 'Visitation Cost' },"
    Response.Write "        height: 600, width: window.innerWidth * 0.95,"
    Response.Write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    Response.Write "        text: yearlyData.map(pair => 'Visit Year: ' + pair.VisitYear + ' <br>Visitation Cost: ' + pair.VisitationCost + '<br>Prev Cost: ' + pair.PrevCost + '<br>Diff: ' + pair.Diff + '<br>Percentage Change: ' + pair.PercentageChange + '<br>Yearly Contribution: ' + pair.ContPercentage),"
    Response.Write "        textposition: 'auto',"
    Response.Write "        texttemplate: '%{y}',"
    Response.Write "        hovertemplate: '%{text}',"
    Response.Write "        legend: { orientation: 'h', y: -0.2, x: 0.5, xanchor: 'center', yanchor: 'top' }"
    Response.Write "    };"

    ' Create the bar chart
    Response.Write "    Plotly.newPlot('visitCostChartDiv', [data], barLayout);"
    Response.Write "});"
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

Function FormatDateNew(dateString)
    Dim dateParts, yearPart, monthPart, dayPart, formatedDate
    dateParts = Split(dateString, "-")
    yearPart = dateParts(0)
    monthPart = dateParts(1)
    dayPart = dateParts(2)
    formatedDate = dayPart & "/" & monthPart & "/" & yearPart
    FormatDateNew = formatedDate
End Function


'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
