Sub VisitStatusTrendYearly()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "WITH selectCTE AS ( "
    sql = sql & "SELECT VisitStatus.VisitStatusName, DATENAME(YEAR, VisitDate) VisitYear,  COUNT(VisitationID) [NumberOfVisits] "
    sql = sql & ", LAG(COUNT(VisitationID)) OVER(PARTITION BY VisitStatus.VisitStatusName ORDER BY VisitStatus.VisitStatusName, DATENAME(YEAR, VisitDate)) [PrevNoOfVisits], "
    sql = sql & "(COUNT(VisitationID) - LAG(COUNT(VisitationID)) OVER(PARTITION BY VisitStatus.VisitStatusName ORDER BY VisitStatus.VisitStatusName, DATENAME(YEAR, VisitDate))) [DiffInVisits] "
    sql = sql & "FROM Visitation JOIN VisitStatus ON Visitation.VisitStatusID = VisitStatus.VisitStatusID "
    
    ' Add date range filter
    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "WHERE CONVERT(DATE, Visitation.VisitDate) BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    Else
        sql = sql & "WHERE CONVERT(DATE, Visitation.VisitDate) BETWEEN '2018-01-01' AND '2022-12-31' "
    End If
    
    sql = sql & "GROUP BY VisitStatus.VisitStatusName, DATENAME(YEAR, VisitDate) "
    sql = sql & ") "
    sql = sql & "SELECT VisitStatusName, VisitYear, FORMAT(NumberOfVisits, 'N2') [NumberOfVisits], FORMAT(PrevNoOfVisits, 'N2') [PrevNoOfVisits] "
    sql = sql & ", FORMAT( DiffInVisits, 'N2') [DiffInVisits]  "
    sql = sql & ", CONVERT(DECIMAL(5, 2), DiffInVisits * 100.00/NumberOfVisits) [YoYChange] "
    sql = sql & ", CONVERT(DECIMAL(5, 2), NumberOfVisits * 100.00/SUM(NumberOfVisits) OVER(PARTITION BY VisitStatusName)) [PercentContToVisitStatusName] "
    sql = sql & ", CONVERT(DECIMAL(5, 2), NumberOfVisits * 100.00/SUM(NumberOfVisits) OVER()) [PercentContToOverallVisits] "
    sql = sql & "FROM selectCTE "

    rst.Open sql, conn, 3, 4

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
            jsonData = jsonData & """VisitStatusName"":""" & CStr(rst.Fields("VisitStatusName").value) & ""","
            jsonData = jsonData & """VisitYear"":""" & rst.Fields("VisitYear").value & ""","
            jsonData = jsonData & """NumberOfVisits"":""" & rst.Fields("NumberOfVisits").value & ""","
            jsonData = jsonData & """PrevNoOfVisits"":""" & rst.Fields("PrevNoOfVisits").value & ""","
            jsonData = jsonData & """DiffInVisits"":""" & rst.Fields("DiffInVisits").value & ""","
            jsonData = jsonData & """YoYChange"":""" & rst.Fields("YoYChange").value & ""","
            jsonData = jsonData & """PercentContToVisitStatusName"":""" & rst.Fields("PercentContToVisitStatusName").value & ""","
            jsonData = jsonData & """PercentContToOverallVisits"":""" & rst.Fields("PercentContToOverallVisits").value & """"
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
    response.Write "<script>"
    response.Write "var dbDataVisitStatus = " & jsonData & ";"
    response.Write "document.addEventListener('DOMContentLoaded', function() {"
    response.Write "    var visitStatus = dbDataVisitStatus.data;"

    ' Defining a color palette
    response.Write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"

    ' Define the chart data with different colors for each bar and group by VisitStatusName
    response.Write "    var data = visitStatus.reduce((acc, row) => {"
    response.Write "        if (!acc[row.VisitStatusName]) acc[row.VisitStatusName] = { x: [], y: [], type: 'bar', name: row.VisitStatusName, marker: { color: colors[Object.keys(acc).length % colors.length] } };"
    response.Write "        acc[row.VisitStatusName].x.push(row.VisitYear);"
    response.Write "        acc[row.VisitStatusName].y.push(row.NumberOfVisits);"
    response.Write "        return acc;"
    response.Write "    }, {});"
    response.Write "    data = Object.values(data);"

    ' Define the layout for the chart
    response.Write "    var barLayout = {"
    response.Write "        title: 'Visit Status Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " ',"
    response.Write "        xaxis: { title: 'Visit Year' },"
    response.Write "        yaxis: { title: 'No. Of Visits' },"
    response.Write "        height: 600, width: window.innerWidth * 0.95,"
    response.Write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    response.Write "        barmode: 'group',"
    
    
    response.Write "        text: visitStatus.map(pair => 'Visit Year: ' + pair.VisitYear + ' <br>Visits: ' + pair.NumberOfVisits + '<br>Prev Year Visits : ' + pair.PrevNoOfVisits + '<br>Diff In Visits: ' + pair.DiffInVisits + '<br>YoY Change: ' + pair.YoYChange + '<br>PercentContToVisitStatusName: ' + pair.PercentContToVisitStatusName + '<br>PercentContToOverallVisits: ' + pair.PercentContToOverallVisits + ' '),"
    
    response.Write "        textposition: 'auto',"
    response.Write "        texttemplate: '%{y}',"

    response.Write "        hovertemplate: '%{text}',"
    response.Write "        legend: { orientation: 'h', y: -0.2, x: 0.5, xanchor: 'center', yanchor: 'top' }" ' Horizontal legend
    response.Write "    };"
    
    ' Create the bar chart
    response.Write "    Plotly.newPlot('visitStatusTrendsChart', data, barLayout);"
    response.Write "});"
    response.Write "</script>"

    ' DataTable Initialization
    response.Write "<script>"
    response.Write "    new DataTable('#visitStatusTrendsTable', {"
    response.Write "        data: dbDataVisitStatus.data,"
    response.Write "        columns: ["
    response.Write "            { data: 'counter' },"
    response.Write "            { data: 'VisitStatusName' },"
    response.Write "            { data: 'VisitYear' },"
    response.Write "            { data: 'NumberOfVisits' },"
    response.Write "            { data: 'PrevNoOfVisits' },"
    response.Write "            { data: 'DiffInVisits' },"
    response.Write "            { data: 'YoYChange' },"
    response.Write "            { data: 'PercentContToVisitStatusName' },"
    response.Write "            { data: 'PercentContToOverallVisits' }"
    response.Write "        ],"
    response.Write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.Write "        dom: 'lBfrtip',"
    response.Write "        buttons: ["
    response.Write "            {"
    response.Write "                extend: 'csv',"
    response.Write "                text: 'CSV',"
    response.Write "                title: '" & brnchName & " Visit Status Trends From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.Write "            },"
    response.Write "            {"
    response.Write "                extend: 'excel',"
    response.Write "                text: 'EXCEL',"
    response.Write "                title: '" & brnchName & " Visit Status Trends From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.Write "            },"
    response.Write "            {"
    response.Write "                extend: 'pdf',"
    response.Write "                text: 'PDF',"
    response.Write "                title: '" & brnchName & " Visit Status Trends From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.Write "            },"
    response.Write "            {"
    response.Write "                extend: 'print',"
    response.Write "                text: 'PRINT',"
    response.Write "                title: '" & brnchName & " Visit Status Trends From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.Write "            }"
    response.Write "        ]"
    response.Write "    });"
    response.Write "</script>"
End Sub