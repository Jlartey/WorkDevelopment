
Sub MedicalOutcomeYearly()
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

    rst.Open sql, conn, 3, 4

    response.write sql

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
            jsonData = jsonData & """MedicalOutcomeID"":""" & CStr(rst.Fields("MedicalOutcomeName").value) & ""","
            jsonData = jsonData & """VisitYear"":""" & rst.Fields("VisitYear").value & ""","
            jsonData = jsonData & """Frequency"":""" & rst.Fields("Frequency").value & ""","
            jsonData = jsonData & """PrevNoOfMedOutcome"":""" & rst.Fields("PrevNoOfMedOutcome").value & ""","
            jsonData = jsonData & """DiffInVisits"":""" & rst.Fields("DiffInVisits").value & """"
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
    response.write "var dbDataVisitCost = " & jsonData & ";"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var visitCost = dbDataVisitCost.data;"

    ' Defining colors
    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"

    ' Define the chart data with different colors for each bar and group by VisitYear
    response.write "    var data = visitCost.reduce((acc, row) => {"
    response.write "        if (!acc[row.VisitYear]) acc[row.VisitYear] = { x: [], y: [], type: 'bar', name: row.VisitYear, marker: { color: colors[Object.keys(acc).length % colors.length] } };"
    response.write "        acc[row.VisitYear].x.push(row.MedicalOutcomeID);"
    response.write "        acc[row.VisitYear].y.push(row.Frequency);"
    response.write "        return acc;"
    response.write "    }, {});"
    response.write "    data = Object.values(data);"

    ' Define the layout for the chart
    response.write "    var barLayout = {"
    response.write "        title: 'Medical Outcome Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & "',"
    response.write "        xaxis: { title: 'Medical Outcome ID' },"
    response.write "        yaxis: { title: 'No. Of Visits' },"
    response.write "        height: 600, width: window.innerWidth * 0.95,"
    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    response.write "        barmode: 'group',"
    response.write "        text: visitCost.map(row => 'Visit Year: ' + row.VisitYear + ' <br>Visits: ' + row.Frequency + '<br>Prev Year Visits: ' + row.PrevNoOfMedOutcome + '<br>Diff In Visits: ' + row.DiffInVisits),"
    response.write "        textposition: 'auto',"
    response.write "        texttemplate: '%{y}',"
    response.write "        hovertemplate: '%{text}',"
    response.write "        legend: { orientation: 'h', y: -0.2, x: 0.5, xanchor: 'center', yanchor: 'top' }"
    response.write "    };"
    
    ' Create the bar chart
    response.write "    Plotly.newPlot('visitCostTrendsChart', data, barLayout);"
    response.write "});"
    response.write "</script>"

    ' DataTable Initialization
    response.write "<script>"
    response.write "    new DataTable('#visitCostTrendsTable', {"
    response.write "        data: dbDataVisitCost.data,"
    response.write "        columns: ["
    response.write "            { data: 'counter' },"
    response.write "            { data: 'MedicalOutcomeID' },"
    response.write "            { data: 'VisitYear' },"
    response.write "            { data: 'Frequency' },"
    response.write "            { data: 'PrevNoOfMedOutcome' },"
    response.write "            { data: 'DiffInVisits' }"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Visit Status Trends From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Visit Status Trends From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Visit Status Trends From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Visit Status Trends From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            }"
    response.write "        ]"
    response.write "    });"
    response.write "</script>"
End Sub