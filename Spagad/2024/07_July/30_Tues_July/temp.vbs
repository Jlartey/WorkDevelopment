' Send the data to the client-side
    response.write "<script>"
    response.write "var dbDataQuarterly = " & jsonData & ";"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var revenueSourcesQuarterly = dbDataQuarterly.data;"
    
    ' Defining a color palette
    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    
    ' Get unique quarters
    response.write "    var uniqueQuarters = [...new Set(revenueSourcesQuarterly.map(pair => pair.Quarter))];"
    
    ' Create traces for each quarter
    response.write "    var traces = uniqueQuarters.map((quarter, index) => {"
    response.write "        var filteredData = revenueSourcesQuarterly.filter(pair => pair.Quarter === quarter);"
    response.write "        return {"
    response.write "            x: filteredData.map(pair => pair.Quarter),"
    response.write "            y: filteredData.map(pair => pair.QuarterTotal),"
    response.write "            type: 'bar',"
   ' response.write "            text: filteredData.map(pair => 'Visitation Cost: ' + pair.VisitationCostF + '<br>Previous Cost: ' + pair.PrevCost + '<br>Difference: ' + pair.Diff + '<br>Percentage Change: ' + pair.PercentageChange + '<br>Cont Percentage: ' + pair.ContPercentage + ' '),"
    response.write "            textposition: 'auto',"
    response.write "            texttemplate: '%{y}',"
    response.write "            hovertemplate: '%{text}',"
    response.write "            marker: {"
    response.write "                color: colors[index % colors.length]"
    response.write "            },"
    response.write "            name: quarter"  ' Setting the quarter as the name for the legend
    response.write "        };"
    response.write "    });"
    
    ' Define the layout for the chart
    response.write "    var barLayout = {"
    response.write "        title: 'Quarterly Insurance Type Trend Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " ',"
    response.write "        xaxis: { title: 'Quarter' },"
    response.write "        yaxis: { title: 'Quarter Total' },"
    response.write "        height: 600, width: window.innerWidth * 0.95,"
    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    response.write "        legend: {"
    response.write "            orientation: 'h',"
    response.write "            x: 0.5,"
    response.write "            xanchor: 'center',"
    response.write "            y: -0.2"
    response.write "        }"
    response.write "    };"
    
    ' Create the bar chart
    response.write "    Plotly.newPlot('quarterlyChartDiv', traces, barLayout);"
    response.write "});"
