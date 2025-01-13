Sub quarterlyVisitCostAnalysis()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    ' Construct SQL query for quarterly data
    sql = "WITH VisitCostCTE AS ("
    sql = sql & "SELECT "
    sql = sql & "    FORMAT(Visitation.VisitDate, 'yyyy') AS VisitYear, "
    sql = sql & "    FORMAT(Visitation.VisitDate, 'yyyy') + '_Q' + DATENAME(QUARTER, Visitation.VisitDate) AS VisitQuarter, "
    sql = sql & "    SUM(Visitation.VisitCost) AS VisitationCost, "
    sql = sql & "    LAG(SUM(Visitation.VisitCost)) OVER(PARTITION BY FORMAT(Visitation.VisitDate, 'yyyy') ORDER BY FORMAT(Visitation.VisitDate, 'yyyy'), DATENAME(QUARTER, Visitation.VisitDate)) AS [PrevCost], "
    sql = sql & "    SUM(Visitation.VisitCost) - LAG(SUM(Visitation.VisitCost)) OVER(PARTITION BY FORMAT(Visitation.VisitDate, 'yyyy') ORDER BY FORMAT(Visitation.VisitDate, 'yyyy'), DATENAME(QUARTER, Visitation.VisitDate)) AS [Diff] "
    sql = sql & "FROM Visitation "
    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "WHERE Visitation.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    Else
        sql = sql & "WHERE Visitation.VisitDate BETWEEN '2018-01-01' AND '2022-12-31' "
    End If
    sql = sql & "GROUP BY FORMAT(Visitation.VisitDate, 'yyyy'), DATENAME(QUARTER, Visitation.VisitDate) "
    sql = sql & "), "
    sql = sql & "GroupCTE AS ("
    sql = sql & "SELECT "
    sql = sql & "    VisitYear, "
    sql = sql & "    VisitQuarter, "
    sql = sql & "    VisitationCost, "
    sql = sql & "    PrevCost, "
    sql = sql & "    Diff, "
    sql = sql & "    (Diff * 100.00) / VisitationCost AS [PercentageChange], "
    sql = sql & "    VisitationCost * 100.00 / SUM(VisitationCost) OVER() AS ContPercentage "
    sql = sql & "FROM VisitCostCTE) "
    sql = sql & "SELECT "
    sql = sql & "    VisitYear, "
    sql = sql & "    VisitQuarter, "
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
            jsonData = jsonData & """VisitQuarter"":""" & CStr(rst.Fields("VisitQuarter").Value) & ""","
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
    response.write "var dbDataQuarterly = " & jsonData & ";"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var revenueSourcesQuarterly = dbDataQuarterly.data;"
    
    ' Defining a color palette
    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    
    ' Define the chart data with different colors for each bar
    response.write "    var trace = {"
    response.write "        x: revenueSourcesQuarterly.map(pair => pair.VisitQuarter),"
    response.write "        y: revenueSourcesQuarterly.map(pair => pair.VisitationCost),"
    response.write "        type: 'bar',"
    response.write "        text: revenueSourcesQuarterly.map(pair => 'Visitation Cost: ' + pair.VisitationCostF + '<br>Previous Cost: ' + pair.PrevCost + '<br>Difference: ' + pair.Diff + '<br>Percentage Change: ' + pair.PercentageChange + '<br>Cont Percentage: ' + pair.ContPercentage + ' '),"
    response.write "        textposition: 'auto',"
    response.write "        texttemplate: '%{y}',"
    response.write "        hovertemplate: '%{text}',"
    response.write "        marker: {"
    response.write "            color: revenueSourcesQuarterly.map((_, index) => colors[index % colors.length])"
    response.write "        }"
    response.write "    };"

    ' Define the layout for the chart
    response.write "    var barLayout = {"
    response.write "        title: 'Visit Cost Analysis Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " ',"
    response.write "        xaxis: { title: 'Quarter' },"
    response.write "        yaxis: { title: 'Visitation Cost' },"
    response.write "        height: 600, width: window.innerWidth * 0.95,"
    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    response.write "    };"
    response.write "    Plotly.newPlot('quarterlyChartDiv', [trace], barLayout);"
    response.write "});"
    response.write "</script>"
End Sub

'This one shows a bar chart
Sub monthlyVisitCostAnalysis02()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    ' Construct SQL query for monthly data with division by zero handling
    sql = "WITH VisitCostCTE AS ("
    sql = sql & "SELECT "
    sql = sql & "    FORMAT(Visitation.VisitDate, 'yyyy') AS VisitYear, "
    sql = sql & "    FORMAT(Visitation.VisitDate, 'yyyy-MM') AS VisitMonth, "
    sql = sql & "    SUM(Visitation.VisitCost) AS VisitationCost, "
    sql = sql & "    LAG(SUM(Visitation.VisitCost)) OVER(PARTITION BY FORMAT(Visitation.VisitDate, 'yyyy') ORDER BY FORMAT(Visitation.VisitDate, 'yyyy-MM')) AS [PrevCost], "
    sql = sql & "    SUM(Visitation.VisitCost) - LAG(SUM(Visitation.VisitCost)) OVER(PARTITION BY FORMAT(Visitation.VisitDate, 'yyyy') ORDER BY FORMAT(Visitation.VisitDate, 'yyyy-MM')) AS [Diff] "
    sql = sql & "FROM Visitation "
    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "WHERE Visitation.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    Else
        sql = sql & "WHERE Visitation.VisitDate BETWEEN '2018-01-01' AND '2022-12-31' "
    End If
    sql = sql & "GROUP BY FORMAT(Visitation.VisitDate, 'yyyy'), FORMAT(Visitation.VisitDate, 'yyyy-MM') "
    sql = sql & "), "
    sql = sql & "GroupCTE AS ("
    sql = sql & "SELECT "
    sql = sql & "    VisitYear, "
    sql = sql & "    VisitMonth, "
    sql = sql & "    VisitationCost, "
    sql = sql & "    PrevCost, "
    sql = sql & "    Diff, "
    sql = sql & "    CASE WHEN VisitationCost != 0 THEN (Diff * 100.00) / VisitationCost ELSE NULL END AS [PercentageChange], "
    sql = sql & "    CASE WHEN SUM(VisitationCost) OVER() != 0 THEN (VisitationCost * 100.00) / SUM(VisitationCost) OVER() ELSE NULL END AS ContPercentage "
    sql = sql & "FROM VisitCostCTE) "
    sql = sql & "SELECT "
    sql = sql & "    VisitYear, "
    sql = sql & "    VisitMonth, "
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
            jsonData = jsonData & """VisitMonth"":""" & CStr(rst.Fields("VisitMonth").Value) & ""","
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
    response.write "var dbDataMonthly = " & jsonData & ";"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var revenueSourcesMonthly = dbDataMonthly.data;"
    
    ' Initialize DataTable with the JSON data
    response.write "    $('#monthlyTable').DataTable({"
    response.write "        data: revenueSourcesMonthly,"
    response.write "        columns: ["
    response.write "            { data: 'counter' },"
    response.write "            { data: 'VisitYear' },"
    response.write "            { data: 'VisitMonth' },"
    response.write "            { data: 'VisitationCostF' },"
    response.write "            { data: 'VisitationCost' },"
    response.write "            { data: 'PrevCost' },"
    response.write "            { data: 'Diff' },"
    response.write "            { data: 'PercentageChange' },"
    response.write "            { data: 'ContPercentage' }"
    response.write "        ],"
    response.write "        pageLength: 10,"
    response.write "        lengthMenu: [10, 25, 50, 75, 100],"
    response.write "        searching: true,"
    response.write "        ordering: true,"
    response.write "        info: true,"
    response.write "        responsive: true"
    response.write "    });"
    
    ' Define the chart data
    response.write "    var trace = {"
    response.write "        x: revenueSourcesMonthly.map(pair => pair.VisitMonth),"
    response.write "        y: revenueSourcesMonthly.map(pair => pair.VisitationCost),"
    response.write "        type: 'bar',"
    response.write "        text: revenueSourcesMonthly.map(pair => 'Visitation Cost: ' + pair.VisitationCostF + '<br>Previous Cost: ' + pair.PrevCost + '<br>Difference: ' + pair.Diff + '<br>Percentage Change: ' + pair.PercentageChange + '<br>Cont Percentage: ' + pair.ContPercentage + ' '),"
    response.write "        textposition: 'auto',"
    response.write "        texttemplate: '%{y}',"
    response.write "        hovertemplate: '%{text}'"
    response.write "    };"

    ' Define the layout for the chart
    response.write "    var barLayout = {"
    response.write "        title: 'Monthly Visit Cost Analysis Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " ',"
    response.write "        xaxis: { title: 'Month' },"
    response.write "        yaxis: { title: 'Visitation Cost' },"
    response.write "        height: 600, width: window.innerWidth * 0.95,"
    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    response.write "    };"
    response.write "    Plotly.newPlot('monthlyChartDiv', [trace], barLayout);"
    response.write "});"
    response.write "</script>"

    ' HTML for the table and chart
    response.write "<div>"
    response.write "<table id='monthlyTable' class='display' style='width:100%'></table>"
    response.write "<div id='monthlyChartDiv'></div>"
    response.write "</div>"
End Sub