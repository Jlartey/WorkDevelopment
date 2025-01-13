'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

response.Clear
conn.commandTimeOut = 7200
Dim periodStart, periodEnd, datePeriod, dateArr

Styling

datePeriod = Trim(Request.QueryString("Dateperiod"))

' Parse date period
If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    periodStart = dateArr(0)
    periodEnd = dateArr(1)
End If

response.write "<!DOCTYPE html>"
response.write "<html lang='en'>"
response.write "<head>"
response.write "<meta charset='UTF-8'>"
response.write "<meta name='viewport' content='width=device-width, initial-scale=1.0'>"
response.write "<title>Visit Cost Analysis</title>"

response.write "<script src='https://cdn.plot.ly/plotly-latest.min.js'></script>"

response.write "    <link href=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css"" rel=""stylesheet"""
response.write "        integrity=""sha384-9ndCyUaIbzAi2FUVXJi0CjmCapSmO7SnpJef0486qhLnuZ2cdeRhO02iuK6FUUVM"" crossorigin=""anonymous"">"

response.write "    <script src=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"""
response.write "        integrity=""sha384-geWF76RCwLtnZ8qwWowPQNguL3RmwHVBC9FhGdlKrxdiJJigb/j/68SIy3Te4Bkz"""
response.write "        crossorigin=""anonymous""></script>"
' Data Tables
response.write " <link href=""https://cdn.datatables.net/v/bs5/jq-3.6.0/jszip-2.5.0/dt-1.13.5/af-2.6.0/b-2.4.0/b-colvis-2.4.0/b-html5-2.4.0/b-print-2.4.0/cr-1.7.0/date-1.5.0/fc-4.3.0/fh-3.4.0/kt-2.10.0/r-2.5.0/rg-1.4.0/rr-1.4.0/sc-2.2.0/sb-1.5.0/sp-2.2.0/sl-1.7.0/sr-1.3.0/datatables.min.css"" rel=""stylesheet""/>"
response.write " <script src=""https://cdn.datatables.net/v/bs5/jq-3.6.0/jszip-2.5.0/dt-1.13.5/af-2.6.0/b-2.4.0/b-colvis-2.4.0/b-html5-2.4.0/b-print-2.4.0/cr-1.7.0/date-1.5.0/fc-4.3.0/fh-3.4.0/kt-2.10.0/r-2.5.0/rg-1.4.0/rr-1.4.0/sc-2.2.0/sb-1.5.0/sp-2.2.0/sl-1.7.0/sr-1.3.0/datatables.min.js""></script>"

'PDF Maker
response.write " <script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js""></script>"
response.write " <script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js""></script>"


response.write "<style>"
response.write "  .chart-container {"
response.write "    display: flex;"
response.write "    justify-content: center;"
response.write "  }"
response.write "  .chart {"
response.write "    flex: 1;"
response.write "    margin: 10px;"
response.write "    width: 80%;"
response.write "  }"

response.write "  .tab-header {"
response.write "    display: flex;"
response.write "    justify-content: center;"
response.write "    background-color: #007bff;"
response.write "    border: 1px solid #ddd;"
response.write "    border-radius: 5px;"
response.write "  }"
response.write "  .tab-button {"
response.write "    flex: 1;"
response.write "    padding: 10px;"
response.write "    text-align: center;"
response.write "    cursor: pointer;"
response.write "    font-weight: bold;"
response.write "    color: #fff;"
response.write "    border-right: 1px solid #ddd;"
response.write "  }"
response.write "  .tab-button:last-child {"
response.write "    border-right: none;"
response.write "  }"
response.write "  .tab-button.active {"
response.write "    background-color: #0056b3;"
response.write "  }"
response.write "  .tab-content {"
response.write "    display: none;"
response.write "    padding: 20px;"
response.write "    border: 1px solid #ddd;"
response.write "    border-radius: 5px;"
response.write "    background-color: #f9f9f9;"
response.write "    margin-top: 10px;"
response.write "  }"
response.write "  .tab-content.active {"
response.write "    display: block;"
response.write "  }"
response.write "</style>"


response.write "</head>"
response.write "<body>"


'calling InitPageScript sub
'InitPageScript

response.write "<script>"
response.write "function openTab(event, tabId) {"
response.write "  var i, tabcontent, tabbuttons;"
response.write "  tabcontent = document.getElementsByClassName('tab-content');"
response.write "  for (i = 0; i < tabcontent.length; i++) {"
response.write "    tabcontent[i].style.display = 'none';"
response.write "  }"
response.write "  tabbuttons = document.getElementsByClassName('tab-button');"
response.write "  for (i = 0; i < tabbuttons.length; i++) {"
response.write "    tabbuttons[i].className = tabbuttons[i].className.replace(' active', '');"
response.write "  }"
response.write "  document.getElementById(tabId).style.display = 'block';"
response.write "  event.currentTarget.className += ' active';"
response.write "}"
response.write "</script>"

response.write "<div class='tab-header'>"
'response.write "  <div class='tab-button' onclick='openTab(event, ""yearlySamePeriodTab"")'>Inter Visit Interval</div>"
response.write "  <div class='tab-button active' onclick='openTab(event, ""yearlyTab"")'>Yearly Visit Cost</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""quarterlyTab"")'>Quarterly Visit Cost</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""monthlyTab"")'>Monthly Visit Cost</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""weeklyTab"")'>Weekly Visit Cost</div>"

response.write "</div>"

'calling filters sub
'filters


 ' Output HTML Form for date selection
    response.write "    <form id='dateForm'>"
    response.write "    <div class='container' style='display: flex; align-items: center; justify-content: center'> "
    response.write "        <div> "
    response.write "            <label for='from'>From</label> "
    response.write "            <input type='date' name='from' id='from'> "
    response.write "        </div> "
    response.write "        <div> "
    response.write "            <label for='to' style='margin-left: 10px'>To</label> "
    response.write "            <input type='date' name='to' id='to'> "
    response.write "        </div> "
    response.write "        <div> "
    response.write "            <button type='button' onclick='updateUrl()'>Show Data</button> <br />"
    response.write "        </div>    "
    response.write "    </div> "
    response.write "   </form>"
    
    If (periodStart <> "" And periodEnd <> "") Then
        response.write "<h2 class='font-style'>SHOWING DATA FROM: " & periodStart & " TO: " & periodEnd & "</h2>"
    Else
        response.write "<h2 class='font-style'>SHOWING DATA FROM: 2018-01-01 TO: 2018-01-31</h2>"
    End If
    

'yearly tab starts here

response.write "<div id='yearlyTab' class='tab-content active'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='yearlyChartDiv' class='chart'></div>"

response.write "  </div>"


response.write "  <div class='chart-container'>"

response.write "    <div id='yearlyChartDivGender' class='chart'></div>"
response.write "  </div>"

    response.write "      <table style=""width:100%"" id=""visitCostTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    response.write "      <thead class=""table-dark"">"
    response.write "            <tr>"
    response.write "                <th>No.</th>"
    response.write "                <th>Visit Year</th>"
    response.write "                <th>Visitation CostF</th>"
    response.write "                <th>Visitation Cost</th>"
    response.write "                <th>Previous Year Cost</th>"
    response.write "                <th>Difference</th>"
    response.write "                <th>Percentage Change</th>"
    response.write "                <th>Yearly Contribution (%)</th>"
    response.write "            </tr>"
    response.write "        </thead>"
    response.write "    </table>"
    
response.write "</div>"


 response.write "<script>"
    response.write "    function updateUrl() {"
    response.write "        const fromDate = document.getElementById('from').value;"
    response.write "        const toDate = document.getElementById('to').value;"
    
    response.write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp';"
    response.write "        const params = new URLSearchParams({"
    response.write "            PrintLayoutName: 'dispVisitCostAnalysis',"
    response.write "            PositionForTableName: 'WorkingDay',"
    response.write "            WorkingDayID: '',"
    response.write "            Dateperiod: `${fromDate}||${toDate}`"
    response.write "        });"
    response.write "        const newUrl = baseUrl + '?' + params.toString();"
    response.write "        window.location.href = newUrl;"
    response.write "        console.log(newUrl);"
    response.write "    }"
response.write "</script>"

'yearly tab end here

'quarterly tab starts here
response.write "<div id='quarterlyTab' class='tab-content'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='quarterlyChartDiv' class='chart'></div>"
response.write "  </div>"

' quarterly table

response.write "      <table style=""width:100%"" id=""quarterlyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
response.write "      <thead class=""table-dark"">"
response.write "              <tr>"
response.write "                <th>No.</th>"
response.write "                <th>Visit Year</th>"
response.write "                <th>Visit Quarter</th>"
response.write "                <th>Visitation CostF</th>"
response.write "                <th>Visitation Cost</th>"
response.write "                <th>Previous Cost</th>"
response.write "                <th>Difference</th>"
response.write "                <th>Percentage Change</th>"
response.write "                <th>Yearly Contribution (%)</th>"
response.write "           </tr>"
response.write "        </thead>"
response.write "    </table>"

response.write "</div>"
'qurterly ends here




response.write "</body>"
response.write "</html>"


dispVisitCostAnalysis

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

    rst.open sql, conn, 3, 4

    ' Generate JSON data
    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    ' Check if the recordset has any records
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """VisitYear"":""" & CStr(rst.Fields("VisitYear").Value) & ""","
            jsonData = jsonData & """VisitationCostF"":""" & rst.Fields("VisitationCostF").Value & ""","
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
    
    ' Send the data to the client-side
    response.write "<script>"
    response.write "var dbDataYearly = " & jsonData & ";"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var revenueSourcesYearly = dbDataYearly.data;"
    
    ' Defining a color palette
    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    
    ' Define the chart data with different colors for each bar
    response.write "    var trace = {"
    response.write "        x: revenueSourcesYearly.map(pair => pair.VisitYear),"
    response.write "        y: revenueSourcesYearly.map(pair => pair.VisitationCost),"
    response.write "        type: 'bar',"
    response.write "        text: revenueSourcesYearly.map(pair => 'Visitation Cost: ' + pair.VisitationCostF + '<br>Previous Cost: ' + pair.PrevCost + '<br>Difference: ' + pair.Diff + '<br>Percentage Change: ' + pair.PercentageChange + '<br>Cont Percentage: ' + pair.ContPercentage + ' '),"
    response.write "        textposition: 'auto',"
    response.write "        texttemplate: '%{y}',"
    response.write "        hovertemplate: '%{text}',"
    response.write "        marker: {"
    response.write "            color: revenueSourcesYearly.map((_, index) => colors[index % colors.length])"
    response.write "        }"
    response.write "    };"

    ' Define the layout for the chart
    response.write "    var barLayout = {"
    response.write "        title: 'Visit Cost Analysis Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " ',"
    response.write "        xaxis: { title: 'Year' },"
    response.write "        yaxis: { title: 'Visitation Cost' },"
    response.write "        height: 600, width: window.innerWidth * 0.95,"
    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    response.write "    };"

    ' Create the bar chart
    response.write "    Plotly.newPlot('yearlyChartDiv', [trace], barLayout);"
    response.write "});"
    
        
    response.write "</script>"

    ' DataTable Initialization
    response.write "<script>"
    response.write "    new DataTable('#visitCostTable', {"
    response.write "        data: dbDataYearly.data,"
    response.write "        columns: ["
    response.write "            { data: 'counter' },"
    response.write "            { data: 'VisitYear' },"
    response.write "            { data: 'VisitationCostF' },"
    response.write "            { data: 'VisitationCost' },"
    response.write "            { data: 'PrevCost' },"
    response.write "            { data: 'Diff' },"
    response.write "            { data: 'PercentageChange' },"
    response.write "            { data: 'ContPercentage' }"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Visit Cost Analysis From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Visit Cost Analysis From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Visit Cost Analysis From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Visit Cost Analysis From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            }"
    response.write "        ]"
    response.write "    });"
    response.write "</script>"
    
End Sub

Sub quarterlyVisitCostAnalysis()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "WITH VisitCostCTE AS ("
sql = sql & " SELECT "
sql = sql & "     FORMAT(Visitation.VisitDate, 'yyyy') AS VisitYear, "
sql = sql & "     FORMAT(Visitation.VisitDate, 'yyyy') + '_Q' + DATENAME(QUARTER, Visitation.VisitDate) AS VisitQuarter, "
sql = sql & "     SUM(Visitation.VisitCost) AS VisitationCost, "
sql = sql & "     LAG(SUM(Visitation.VisitCost)) "
sql = sql & "         OVER(PARTITION BY FORMAT(Visitation.VisitDate, 'yyyy') "
sql = sql & "         ORDER BY FORMAT(Visitation.VisitDate, 'yyyy'), DATENAME(QUARTER, Visitation.VisitDate)) AS [PrevCost], "
sql = sql & "     SUM(Visitation.VisitCost) - LAG(SUM(Visitation.VisitCost)) "
sql = sql & "         OVER(PARTITION BY FORMAT(Visitation.VisitDate, 'yyyy') "
sql = sql & "         ORDER BY FORMAT(Visitation.VisitDate, 'yyyy'), DATENAME(QUARTER, Visitation.VisitDate)) AS [Diff] "
sql = sql & " FROM Visitation "
sql = sql & " WHERE "
sql = sql & "     Visitation.VisitDate BETWEEN '2018-01-01' AND '2022-12-31' "
sql = sql & " GROUP BY "
sql = sql & "     FORMAT(Visitation.VisitDate, 'yyyy'), "
sql = sql & "     DATENAME(QUARTER, Visitation.VisitDate) "
sql = sql & "), "
sql = sql & "GroupCTE AS ("
sql = sql & " SELECT "
sql = sql & "     VisitYear, "
sql = sql & "     VisitQuarter, "
sql = sql & "     VisitationCost, "
sql = sql & "     PrevCost, "
sql = sql & "     Diff, "
sql = sql & "     (Diff * 100.00)/VisitationCost AS [PercentageChange], "
sql = sql & "     VisitationCost * 100.00 / SUM(VisitationCost) OVER() AS ContPercentage "
sql = sql & " FROM VisitCostCTE "
sql = sql & ") "
sql = sql & "SELECT "
sql = sql & "     VisitYear, "
sql = sql & "     VisitQuarter, "
sql = sql & "     FORMAT(VisitationCost, 'N2') AS VisitationCostF, "
sql = sql & "     VisitationCost, "
sql = sql & "     FORMAT(PrevCost, 'N2') AS PrevCost, "
sql = sql & "     FORMAT(Diff, 'N2') AS Diff, "
sql = sql & "     FORMAT(PercentageChange, 'N2') AS PercentageChange, "
sql = sql & "     FORMAT(ContPercentage, 'N2') AS ContPercentage "
sql = sql & "FROM GroupCTE "
sql = sql & "ORDER BY "
sql = sql & "     VisitYear, "
sql = sql & "     CASE "
sql = sql & "         WHEN RIGHT(VisitQuarter, 2) = 'Q1' THEN 1 "
sql = sql & "         WHEN RIGHT(VisitQuarter, 2) = 'Q2' THEN 2 "
sql = sql & "         WHEN RIGHT(VisitQuarter, 2) = 'Q3' THEN 3 "
sql = sql & "         WHEN RIGHT(VisitQuarter, 2) = 'Q4' THEN 4 "
sql = sql & "     END"

response.write sql

rst.open sql, conn, 3, 4

    ' Generate JSON data
    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    ' Check if the recordset has any records
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """VisitYear"":""" & CStr(rst.Fields("VisitYear").Value) & ""","
            jsonData = jsonData & """VisitQuarter"":""" & rst.Fields("VisitQuarter").Value & ""","
            jsonData = jsonData & """VisitationCostF"":""" & rst.Fields("VisitationCostF").Value & ""","
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
    
'    ' Send the data to the client-side
'    response.write "<script>"
'    response.write "var dbDataYearly = " & jsonData & ";"
'    response.write "document.addEventListener('DOMContentLoaded', function() {"
'    response.write "    var revenueSourcesYearly = dbDataYearly.data;"
'
'    ' Defining a color palette
'    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
'
'    ' Define the chart data with different colors for each bar
'    response.write "    var trace = {"
'    response.write "        x: revenueSourcesYearly.map(pair => pair.VisitYear),"
'    response.write "        y: revenueSourcesYearly.map(pair => pair.VisitationCost),"
'    response.write "        type: 'bar',"
'    response.write "        text: revenueSourcesYearly.map(pair => 'Visitation Cost: ' + pair.VisitationCostF + '<br>Previous Cost: ' + pair.PrevCost + '<br>Difference: ' + pair.Diff + '<br>Percentage Change: ' + pair.PercentageChange + '<br>Cont Percentage: ' + pair.ContPercentage + ' '),"
'    response.write "        textposition: 'auto',"
'    response.write "        texttemplate: '%{y}',"
'    response.write "        hovertemplate: '%{text}',"
'    response.write "        marker: {"
'    response.write "            color: revenueSourcesYearly.map((_, index) => colors[index % colors.length])"
'    response.write "        }"
'    response.write "    };"
'
'    ' Define the layout for the chart
'    response.write "    var barLayout = {"
'    response.write "        title: 'Visit Cost Analysis Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " ',"
'    response.write "        xaxis: { title: 'Year' },"
'    response.write "        yaxis: { title: 'Visitation Cost' },"
'    response.write "        height: 600, width: window.innerWidth * 0.95,"
'    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
'    response.write "    };"
'
'    ' Create the bar chart
'    response.write "    Plotly.newPlot('yearlyChartDiv', [trace], barLayout);"
'    response.write "});"
'
'
'    response.write "</script>"

    ' DataTable Initialization
    response.write "<script>"
    response.write "    new DataTable('#quarterlyTable', {"
    response.write "        data: dbDataYearly.data,"
    response.write "        columns: ["
    response.write "            { data: 'counter' },"
    response.write "            { data: 'VisitYear' },"
    response.write "            { data: 'VisitQuarter' },"
    response.write "            { data: 'VisitationCostF' },"
    response.write "            { data: 'VisitationCost' },"
    response.write "            { data: 'PrevCost' },"
    response.write "            { data: 'Diff' },"
    response.write "            { data: 'PercentageChange' },"
    response.write "            { data: 'ContPercentage' }"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Visit Cost Analysis From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Visit Cost Analysis From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Visit Cost Analysis From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Visit Cost Analysis From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            }"
    response.write "        ]"
    response.write "    });"
    response.write "</script>"

End Sub
Sub Styling()
    response.write " <style>"
        response.write " .mytable {"
        response.write "     width: 95vw;"
        response.write "     font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
        response.write "     border-collapse: collapse;"
        response.write "     margin-top: 50px; "
        response.write "     border-radius: 10px;"
        response.write " }"
        
        response.write " .header {"
        response.write "    display: flex;"
        response.write "    justify-content: center;"
        response.write "    align-items: center;"
        response.write " } "
        
        response.write " .font-style {"
        response.write "    text-align: center;"
        response.write " } "
        
        response.write " .container {"
        response.write "    display: flex"
        response.write "    margin-top: 50px !important;"
        response.write "    padding-top: 30px;"
        response.write " } "
        
        response.write " .myth, .mytd {"
        response.write "     border: 1px solid #ddd;"
        response.write "     padding: 10px;"
        response.write " }"
        
        response.write " .mytd {"
        response.write "     text-alig: 1px solid #ddd;"
        response.write "     padding: 8px;"
        response.write " }"
        
        response.write "  tr:nth-child(even) {"
        response.write "    background-color: rgba(249, 249, 249, 6);"
        response.write " } "
        
        response.write " .myth {"
        response.write "     background-color: #c2c2c2;"
        response.write "     color: black;"
        response.write "     text-align: center; "
        response.write "     text-transform: uppercase; "
        response.write "     font-size: 18px;"
        response.write " }"
        
        response.write "  button {"
        response.write "     background-color: #0236c4;"
        response.write "     border-radius: 5px;"
        response.write "     border: none;"
        response.write "     margin-left: 50px;"
        response.write "     padding: 5px 20px;"
        response.write "     color: white;"
        response.write "     cursor: pointer;"
        response.write "  }"
        
        response.write "  #to, #from {"
        response.write "    padding: 5px;"
        response.write "    border-radius: 5px;"
        response.write "    cursor: pointer;"
        response.write "  }"
        
        response.write " .font-style {"
        response.write "    font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif;"
        response.write " }"
       
        response.write " </style>"
        
End Sub

Function FormatDate(dateValue)
    FormatDate = Year(dateValue) & "-" & Right("0" & Month(dateValue), 2) & "-" & Right("0" & day(dateValue), 2)
End Function

Function FormatDateNew(dateString)
    Dim dateParts, yearPart, monthPart, dayPart, formatedDate
    dateParts = Split(dateString, "-")
    yearPart = dateParts(0)
    monthPart = dateParts(1)
    dayPart = dateParts(2)

    ' Array of month names
    Dim monthNames
    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    Dim monthName
    monthName = monthNames(CInt(monthPart) - 1) ' Subtract 1 for zero-based index

    formatedDate = dayPart & "-" & monthName & "-" & yearPart
    FormatDateNew = formatedDate
End Function


Function GetComboName(table, id)
    GetComboName = "Branch Name"
End Function





'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>


