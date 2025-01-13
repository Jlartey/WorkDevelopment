'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

response.Clear
conn.commandTimeOut = 7200
Dim periodStart, periodEnd, brnchID
Dim page_title
page_title = ""

If Len(Trim(request.querystring("selectedValue"))) > 1 Then
    periodStart = Trim(request.querystring("selectedValue"))
    periodEnd = Trim(request.querystring("selectedValue1"))
    brnchID = Trim(request.querystring("branID"))

    periodStart = FormatDate(periodStart)
    periodEnd = FormatDate(periodEnd)
    brnchID = Trim(brnchID)
Else
    periodStart = FormatDate(Now - 1)
    periodEnd = FormatDate(Now)
    brnchID = "B001"
End If

'page_title = "Admission Analysis Between " & periodStart & " And " & periodEnd & " :" & GetComboName("Branch", brnchID) & " "

AddCss


response.write "<!DOCTYPE html>"
response.write "<html lang='en'>"
response.write "<head>"
response.write "<meta charset='UTF-8'>"
response.write "<meta name='viewport' content='width=device-width, initial-scale=1.0'>"
response.write "<title>Visitation  Analysis</title>"

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
InitPageScript

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
response.write "  <div class='tab-button active' onclick='openTab(event, ""yearlyTab"")'>Inter Visit Intervals</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""quarterlyTab"")'>Inter Visit Range Proportions</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""monthlyTab"")'>Visit Intervals By Gender</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""weeklyTab"")'>Visit Intervals By Age Group</div>"

response.write "</div>"

'calling filters sub
filters

'yearly tab starts here

response.write "<div id='yearlyTab' class='tab-content active'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='yearlyChartDiv' class='chart'></div>"

response.write "  </div>"

response.write "  <div class='chart-container'>"

response.write "    <div id='yearlyChartDivGender' class='chart'></div>"
response.write "  </div>"

response.write "      <table style=""width:100%"" id=""interVisitTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
response.write "      <thead class=""table-dark"">"
response.write "              <tr>"
response.write "                <th>S/No.</th>"
response.write "                <th>Inter Visit Interval</th>"
response.write "                <th>Frequency</th>"
response.write "                <th>Max. InterVisit Interval</th>"
response.write "                <th>Min. InterVisit Interval</th>"
response.write "                <th>Avg. InterVisit Interval</th>"
response.write "                <th>25th Percentile</th>"
response.write "                <th>75th Percentile</th>"
response.write "                <th>Median Inter Visit Interval</th>"
response.write "              </tr>"
response.write "        </thead>"
response.write "    </table>"
    
response.write "</div>"

'yearly tab end here

'quRTERly tab starts here
response.write "<div id='quarterlyTab' class='tab-content'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='quarterlyChartDiv' class='chart'></div>"
response.write "  </div>"

' quarterly table

   response.write "      <table style=""width:100%"" id=""proportionTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    response.write "      <thead class=""table-dark"">"
    response.write "              <tr>"
     response.write "                <th>S/No.</th>"
    response.write "                <th>Interval Range</th>"
     response.write "                <th>Frequency</th>"
      response.write "                <th>Proportion (%)</th>"
    
    
    response.write "                            </tr>"
    response.write "        </thead>"
    response.write "    </table>"

response.write "</div>"
'qurterly ends here


' monthly tab starts here
response.write "<div id='monthlyTab' class='tab-content'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='monthlyChartDiv' class='chart'></div>"
response.write "  </div>"
response.write "<br>"
response.write "    <div id='btnMonthDetails' ></div>"


' monthly table

   response.write "      <table style=""width:100%"" id=""monthlyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    response.write "      <thead class=""table-dark"">"
    response.write "              <tr>"
     response.write "                <th>S/No.</th>"
     response.write "                <th>Gender</th>"
    response.write "                <th>Visit Interval </th>"
     response.write "                <th>Frequency</th>"
    response.write "                <th>Max  Visit Interval</th>"
    response.write "                <th>Min Visit Interval</th>"
    response.write "                <th>Avg Visit Interval</th>"
    response.write "                <th>25TH Percentile</th>"
    response.write "                <th>75TH Percentile</th>"
     response.write "                <th>Median Visit Interval</th>"

    response.write "                            </tr>"
    response.write "        </thead>"
    response.write "    </table>"

response.write "  <div class='chart-container'>"
response.write "    <div id='monthlyGenderVisitsChartDiv' class='chart'></div>"
response.write "  </div>"

' monthly table

   response.write "      <table style=""width:100%"" id=""monthlyGenderTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    response.write "      <thead class=""table-dark"">"
    response.write "              <tr>"
     response.write "                <th>S/No.</th>"
    response.write "                <th>Year</th>"
     response.write "                <th>Month</th>"
    response.write "                <th>Gender</th>"
    response.write "                <th>No. Of Visits</th>"
    response.write "                <th>Prev No. Of Visits</th>"
    response.write "                <th>Difference</th>"
    response.write "                <th>% Change</th>"
     response.write "                <th>% Cont. To Age Group</th>"
    response.write "                <th>% Cont. To Inter Visit Interval</th>"
     response.write "                <th>Cumulative Monthly Visits </th>"
      response.write "                <th>Overall Visits</th>"
    response.write "                <th>% To Overall Visits</th>"
    
    response.write "                            </tr>"
    response.write "        </thead>"
    response.write "    </table>"


response.write "</div>"
'monthly tab ends here

'weekly tab starts here
response.write "<div id='weeklyTab' class='tab-content'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='weeklyVisitsChartDiv' class='chart'></div>"
response.write "  </div>"

' weekly table

   response.write "      <table style=""width:100%"" id=""weeklyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    response.write "      <thead class=""table-dark"">"
    response.write "              <tr>"
     response.write "                <th>S/No.</th>"
    response.write "                <th>Age Group</th>"
     response.write "                <th>Visit Interval</th>"
    response.write "                <th>Frequency</th>"
    response.write "                <th>Max Visit Interval</th>"
    response.write "                <th>Min Visit Interval</th>"
    response.write "                <th>Avg Visit Interval</th>"
    response.write "                <th>25TH Percentile</th>"
    response.write "                <th>75TH Percentile</th>"
     response.write "                <th>Median  Visit Interval</th>"
    response.write "                <th>Median  Visit Interval</th>"
   
     
    response.write "                            </tr>"
    response.write "        </thead>"
    response.write "    </table>"



response.write "  <div class='chart-container'>"
response.write "    <div id='weeklyGenderVisitsChartDiv' class='chart'></div>"
response.write "  </div>"


response.write "</div>"
'weekly tab ends here

'same period revenue-yearly
'response.write "<div id='yearlySamePeriodTab' class='tab-content'>"
'response.write "  <div class='chart-container'>"
'response.write "    <div id='yearlySamePeriodChartDiv' class='chart'></div>"
'response.write "  </div>"
'
'response.write "</div>"

response.write "</body>"
response.write "</html>"




get_inter_visit_intervals_by_age_group

get_inter_visit_proportions_analysis

get_inter_visit_intervals_by_gender
get_inter_visit_intervals



'========================================================
Sub get_inter_visit_proportions_analysis()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")

    sql = "   WITH SelectCTE AS ("
    sql = sql & " SELECT"
    sql = sql & "   LAG(v.VisitDate) OVER (PARTITION BY v.PatientID"
    sql = sql & "    ORDER BY v.VisitDate) AS PreviousVisitDate,"
    sql = sql & "   DATEDIFF(DAY, LAG(v.VisitDate)"
    sql = sql & "   OVER (PARTITION BY v.PatientID"
    sql = sql & "   ORDER BY v.VisitDate), v.VisitDate)"
    sql = sql & "   AS InterVisitInterval"
    sql = sql & " FROM Visitation v JOIN Patient p"
    sql = sql & " ON p.PatientID=v.PatientID"
    sql = sql & " WHERE CONVERT(DATE,v.VisitDate) BETWEEN"
    sql = sql & " '" & periodStart & "' AND '" & periodEnd & "'"
    sql = sql & " ),"
    sql = sql & " IntervalRanges AS ("
    sql = sql & "    SELECT"
    sql = sql & "       CASE"
    sql = sql & "           WHEN InterVisitInterval <= 7 THEN '0-7 Days'"
    sql = sql & "          WHEN InterVisitInterval <= 14 THEN '8-14 Days'"
    sql = sql & "          WHEN InterVisitInterval <= 30 THEN '15-30 Days'"
    sql = sql & "           WHEN InterVisitInterval <= 60 THEN '31-60 Days'"
    sql = sql & "          WHEN InterVisitInterval <= 90 THEN '61-90 Days'"
    sql = sql & "          ELSE '91+ Days'"
    sql = sql & "      END AS IntervalRange,"
    sql = sql & "      COUNT(*) AS Frequency"
    sql = sql & "  FROM SelectCTE"
    sql = sql & "   WHERE InterVisitInterval IS NOT NULL"
    sql = sql & "   GROUP BY"
    sql = sql & "       CASE"
    sql = sql & "          WHEN InterVisitInterval <= 7 THEN '0-7 Days'"
    sql = sql & "          WHEN InterVisitInterval <= 14 THEN '8-14 Days'"
    sql = sql & "         WHEN InterVisitInterval <= 30 THEN '15-30 Days'"
    sql = sql & "         WHEN InterVisitInterval <= 60 THEN '31-60 Days'"
    sql = sql & "          WHEN InterVisitInterval <= 90 THEN '61-90 Days'"
    sql = sql & "          ELSE '91+ Days'"
    sql = sql & "      END"
    sql = sql & " )"
    sql = sql & " SELECT"
    sql = sql & "   IntervalRange,"
    sql = sql & "   Frequency,"
    sql = sql & "   FORMAT(Frequency,'N0') AS FrequencyF,"
    sql = sql & "  Frequency * 100.0 / SUM(Frequency) OVER() AS Proportion,"
    sql = sql & "  FORMAT(Frequency * 100.0 / SUM(Frequency) OVER(),'N2') AS ProportionF"
    sql = sql & " FROM IntervalRanges"
    sql = sql & " ORDER BY IntervalRange"

    rst.open sql, conn, 3, 4

    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """IntervalRange"":""" & rst.Fields("IntervalRange").value & ""","
            jsonData = jsonData & """Frequency"":""" & rst.Fields("Frequency").value & ""","
            jsonData = jsonData & """FrequencyF"":""" & rst.Fields("FrequencyF").value & ""","
            jsonData = jsonData & """Proportion"":""" & rst.Fields("Proportion").value & ""","
            jsonData = jsonData & """ProportionF"":""" & rst.Fields("ProportionF").value & """"
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

    ' Creating the data for the donut chart
    response.write "    var intervalRanges = revenueSourcesQuarterly.map(pair => pair.IntervalRange);"
    response.write "    var frequencies = revenueSourcesQuarterly.map(pair => parseFloat(pair.Frequency));"
    response.write "    var proportions = revenueSourcesQuarterly.map(pair => parseFloat(pair.Proportion));"
    response.write "    var traces = [{"
    response.write "        type: 'pie',"
    response.write "        labels: intervalRanges,"
    response.write "        values: frequencies,"
    response.write "        textinfo: 'label+percent',"
    
'    response.write "            text: filteredData.map(pair => 'Frequency: ' + pair.FrequencyF + '<br>Proportion: ' + pair.ProportionF + '%'),"
'    response.write "            hovertemplate: '%{text}<extra></extra>',"
    
    response.write "        insidetextorientation: 'radial',"
    response.write "        hole: 0.4,"
    response.write "        marker: {"
    response.write "            colors: ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F']"
    response.write "        }"
    response.write "    }];"

    ' Layout for donut chart
    response.write "    var donutLayout = {"
    response.write "        title: 'Inter Visit Interval Analysis by Interval Range Between " & FormatDateNew(periodStart) & " And  " & FormatDateNew(periodEnd) & "',"
    response.write "        height: 600, width: window.innerWidth * 1.0,"
'    response.write "        showlegend: true"
     response.write "        legend: { orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    response.write "    };"

    ' Plot the donut chart
    response.write "    Plotly.newPlot('quarterlyChartDiv', traces, donutLayout);"
    response.write "});"
    response.write "</script>"

    ' DataTable Initialization
    response.write "<script>"
    response.write "    new DataTable('#proportionTable', {"
    response.write "        data: dbDataQuarterly.data,"
    response.write "        columns: ["
    response.write "            { data: 'counter' },"
    response.write "            { data: 'IntervalRange' },"
'    response.write "            { data: 'Frequency' },"
    response.write "            { data: 'FrequencyF' },"
'    response.write "            { data: 'Proportion' },"
    response.write "            { data: 'ProportionF' }"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals Proportions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals Proportions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals Proportions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals Proportions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            }"
    response.write "        ]"
    response.write "    });"
    response.write "</script>"
End Sub
'-===================================


'=============================================================
Sub get_inter_visit_intervals()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "WITH SelectCTE AS ("
    sql = sql & "    SELECT"
    sql = sql & "        LAG(v.VisitDate) OVER (PARTITION BY v.PatientID ORDER BY v.VisitDate) AS PreviousVisitDate,"
    sql = sql & "        DATEDIFF(DAY, LAG(v.VisitDate) OVER (PARTITION BY v.PatientID ORDER BY v.VisitDate), v.VisitDate) AS InterVisitInterval"
    sql = sql & "    FROM Visitation v"
    sql = sql & "    JOIN Patient p ON p.PatientID = v.PatientID"
    sql = sql & "    WHERE CONVERT(DATE, v.VisitDate) BETWEEN '2018-01-01' AND '2018-12-31'"
    sql = sql & ")"
    sql = sql & "  SELECT TOP 20"
    sql = sql & "    InterVisitInterval,"
    sql = sql & "    COUNT(*) AS Frequency,"
    sql = sql & "    FORMAT(COUNT(*), 'N0') AS FrequencyF,"
    sql = sql & "    MAX(InterVisitInterval) OVER() AS MaxInterVisitInterval,"
    sql = sql & "    MIN(InterVisitInterval) OVER() AS MinInterVisitInterval,"
    sql = sql & "    AVG(InterVisitInterval) OVER() AS AvgInterVisitInterval,"
    sql = sql & "    PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY InterVisitInterval) OVER() AS Percentile25,"
    sql = sql & "    PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY InterVisitInterval) OVER() AS MedianInterVisitInterval,"
    sql = sql & "    PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY InterVisitInterval) OVER() AS Percentile75"
    sql = sql & "  FROM SelectCTE"
    sql = sql & " WHERE (InterVisitInterval IS NOT NULL)"
    sql = sql & " AND (CONVERT(INT, InterVisitInterval) > 0)"
    sql = sql & " GROUP BY InterVisitInterval"
    sql = sql & " ORDER BY Frequency DESC"

 
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
            jsonData = jsonData & """InterVisitInterval"":""" & CStr(rst.Fields("InterVisitInterval").value) & ""","
            jsonData = jsonData & """Frequency"":""" & rst.Fields("Frequency").value & ""","
            jsonData = jsonData & """FrequencyF"":""" & rst.Fields("FrequencyF").value & ""","
            jsonData = jsonData & """MaxInterVisitInterval"":""" & rst.Fields("MaxInterVisitInterval").value & ""","
            jsonData = jsonData & """MinInterVisitInterval"":""" & rst.Fields("MinInterVisitInterval").value & ""","
            jsonData = jsonData & """AvgInterVisitInterval"":""" & rst.Fields("AvgInterVisitInterval").value & ""","
            jsonData = jsonData & """Percentile25"":""" & rst.Fields("Percentile25").value & ""","
            jsonData = jsonData & """Percentile75"":""" & rst.Fields("Percentile75").value & ""","
            jsonData = jsonData & """MedianInterVisitInterval"":""" & rst.Fields("MedianInterVisitInterval").value & """"
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
    response.write "var dbDataYearlyGender = " & jsonData & ";"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var revenueSourcesYearly = dbDataYearlyGender.data;"
    
    ' Defining a color palette
    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    
    ' Define the chart data with different colors for each bar
    response.write "    var trace = {"
    response.write "        x: revenueSourcesYearly.map(pair => pair.InterVisitInterval+' Days'),"
    response.write "        y: revenueSourcesYearly.map(pair => pair.Frequency),"
    response.write "        type: 'bar',"
    response.write "        text: revenueSourcesYearly.map(pair => 'InterVisitInterval: ' + pair.InterVisitInterval + ' Days<br>Frequency: ' + pair.FrequencyF + '<br>Max Interval: ' + pair.MaxInterVisitInterval + '<br>Min Interval: ' + pair.MinInterVisitInterval + '<br>Average Interval: ' + pair.AvgInterVisitInterval + '<br>25th Percentile: ' + pair.Percentile25 + '<br>75th Percentile: ' + pair.Percentile75 + '<br>Median Interval: ' + pair.MedianInterVisitInterval + ' '),"
    
    response.write "        textposition: 'auto',"
response.write "        texttemplate: '%{y}',"

    response.write "        hovertemplate: '%{text}',"
    response.write "        marker: {"
    response.write "            color: revenueSourcesYearly.map((_, index) => colors[index % colors.length])"
    response.write "        }"
    response.write "    };"

    ' Define the layout for the chart
    response.write "    var barLayout = {"
    response.write "        title: 'Inter Visit Interval Analysis Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " ',"
    response.write "        xaxis: { title: 'Inter Visit Interval ' },"
    response.write "        yaxis: { title: 'Frequency' },"
    response.write "        height: 600, width: window.innerWidth * 0.95,"
    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    response.write "    };"

    ' Create the bar chart
    response.write "    Plotly.newPlot('yearlyChartDivGender', [trace], barLayout);"
    response.write "});"
    response.write "</script>"

    ' DataTable Initialization
    response.write "<script>"
    response.write "    new DataTable('#interVisitTable', {"
    response.write "        data: dbDataYearlyGender.data,"
    response.write "        columns: ["
    response.write "            { data: 'counter' },"
    response.write "            { data: 'InterVisitInterval' },"
    response.write "            { data: 'FrequencyF' },"
    response.write "            { data: 'MaxInterVisitInterval' },"
    response.write "            { data: 'MinInterVisitInterval' },"
    response.write "            { data: 'AvgInterVisitInterval' },"
    response.write "            { data: 'Percentile25' },"
    response.write "            { data: 'Percentile75' },"
    response.write "            { data: 'MedianInterVisitInterval' }"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            }"
    response.write "        ]"
    response.write "    });"
    response.write "</script>"
End Sub
'==========================================================================
Sub get_inter_visit_intervals_by_gender()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "WITH SelectCTE AS ("
    sql = sql & "    SELECT"
    sql = sql & "        CASE p.genderid WHEN 'GEN01' THEN 'Male'"
    sql = sql & "                      WHEN 'GEN02' THEN 'Female' ELSE 'NA' "
    sql = sql & "        END AS Gender,"
    sql = sql & "        LAG(v.VisitDate) OVER (PARTITION BY v.PatientID ORDER BY v.VisitDate) AS PreviousVisitDate,"
    sql = sql & "        DATEDIFF(DAY, LAG(v.VisitDate) OVER (PARTITION BY v.PatientID ORDER BY v.VisitDate), v.VisitDate) AS InterVisitInterval"
    sql = sql & "    FROM Visitation v"
    sql = sql & "    JOIN Patient p ON p.PatientID = v.PatientID"
    sql = sql & "    WHERE CONVERT(DATE, v.VisitDate) BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    sql = sql & ")"
    sql = sql & "  SELECT TOP 20 "
    sql = sql & "    InterVisitInterval,"
    sql = sql & "    Gender,"
    sql = sql & "    COUNT(*) AS Frequency,"
    sql = sql & "    FORMAT(COUNT(*), 'N0') AS FrequencyF,"
    sql = sql & "    MAX(InterVisitInterval) OVER(PARTITION BY Gender) AS MaxInterVisitInterval,"
    sql = sql & "    MIN(InterVisitInterval) OVER(PARTITION BY Gender) AS MinInterVisitInterval,"
    sql = sql & "    AVG(InterVisitInterval) OVER(PARTITION BY Gender) AS AvgInterVisitInterval,"
    sql = sql & "    PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY InterVisitInterval) OVER(PARTITION BY Gender) AS Percentile25,"
    sql = sql & "    PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY InterVisitInterval) OVER(PARTITION BY Gender) AS MedianInterVisitInterval,"
    sql = sql & "    PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY InterVisitInterval) OVER(PARTITION BY Gender) AS Percentile75"
    sql = sql & "  FROM SelectCTE"
    sql = sql & " WHERE (InterVisitInterval IS NOT NULL)"
    sql = sql & " AND (CONVERT(INT, InterVisitInterval) > 0)"
    sql = sql & " GROUP BY InterVisitInterval, Gender"
    sql = sql & " ORDER BY InterVisitInterval, Gender"

    rst.open sql, conn, 3, 4

    ' Generate JSON data
    Dim jsonDatagender, counter
    counter = 1
    jsonDatagender = "{""data"":["
    
    ' Check if the recordset has any records
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            jsonDatagender = jsonDatagender & "{"
            jsonDatagender = jsonDatagender & """counterG"":""" & counter & ""","
            jsonDatagender = jsonDatagender & """InterVisitIntervalG"":""" & CStr(rst.Fields("InterVisitInterval").value) & ""","
            jsonDatagender = jsonDatagender & """GenderG"":""" & rst.Fields("Gender").value & ""","
            jsonDatagender = jsonDatagender & """FrequencyG"":""" & rst.Fields("Frequency").value & ""","
            jsonDatagender = jsonDatagender & """FrequencyFG"":""" & rst.Fields("FrequencyF").value & ""","
            jsonDatagender = jsonDatagender & """MaxInterVisitIntervalG"":""" & rst.Fields("MaxInterVisitInterval").value & ""","
            jsonDatagender = jsonDatagender & """MinInterVisitIntervalG"":""" & rst.Fields("MinInterVisitInterval").value & ""","
            jsonDatagender = jsonDatagender & """AvgInterVisitIntervalG"":""" & rst.Fields("AvgInterVisitInterval").value & ""","
            jsonDatagender = jsonDatagender & """Percentile25G"":""" & rst.Fields("Percentile25").value & ""","
            jsonDatagender = jsonDatagender & """Percentile75G"":""" & rst.Fields("Percentile75").value & ""","
            jsonDatagender = jsonDatagender & """MedianInterVisitIntervalG"":""" & rst.Fields("MedianInterVisitInterval").value & """"
            jsonDatagender = jsonDatagender & "},"
            rst.MoveNext
            counter = counter + 1
        Loop
        jsonDatagender = Left(jsonDatagender, Len(jsonDatagender) - 1) ' Remove the trailing comma
    End If

    jsonDatagender = jsonDatagender & "]}"

    rst.Close
    Set rst = Nothing

    ' Send the data to the client-side
    response.write "<script>"
    response.write "var dbDataYearly = " & jsonDatagender & ";"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var revenueSources = dbDataYearly.data;"
    
    ' Define colors for each gender
    response.write "    var colorMap = { 'Male': '#4682B4', 'Female': '#FF6347' };"
    
    ' Create traces for each gender
    response.write "    var maleData = revenueSources.filter(d => d.GenderG === 'Male');"
    response.write "    var femaleData = revenueSources.filter(d => d.GenderG === 'Female');"
    
    response.write "    var traceMale = {"
    response.write "        x: maleData.map(pair => pair.InterVisitIntervalG+' Days'),"
    response.write "        y: maleData.map(pair => pair.FrequencyG),"
    response.write "        name: 'Male',"
    response.write "        type: 'bar',"
    response.write "        marker: { color: colorMap['Male'] },"
    response.write "        text: maleData.map(pair => 'Gender: ' + pair.GenderG + '<br>Inter Visit Interval: ' + pair.InterVisitIntervalG + ' Days<br>Frequency: ' + pair.FrequencyFG + '<br>Max Interval: ' + pair.MaxInterVisitIntervalG + '<br>Min Interval: ' + pair.MinInterVisitIntervalG + '<br>Average Interval: ' + pair.AvgInterVisitIntervalG + '<br>25th Percentile: ' + pair.Percentile25G + '<br>75th Percentile: ' + pair.Percentile75G + '<br>Median Interval: ' + pair.MedianInterVisitIntervalG + ' '),"
    response.write "        textposition: 'auto',"
    response.write "        texttemplate: '%{y}',"
    response.write "        hovertemplate: '%{text}'"
    response.write "    };"
    
    response.write "    var traceFemale = {"
    response.write "        x: femaleData.map(pair => pair.InterVisitIntervalG+' Days'),"
    response.write "        y: femaleData.map(pair => pair.FrequencyG),"
    response.write "        name: 'Female',"
    response.write "        type: 'bar',"
    response.write "        marker: { color: colorMap['Female'] },"
    response.write "        text: femaleData.map(pair => 'Gender: ' + pair.GenderG + '<br>Inter Visit Interval: ' + pair.InterVisitIntervalG + ' Days<br>Frequency: ' + pair.FrequencyFG + '<br>Max Interval: ' + pair.MaxInterVisitIntervalG + '<br>Min Interval: ' + pair.MinInterVisitIntervalG + '<br>Average Interval: ' + pair.AvgInterVisitIntervalG + '<br>25th Percentile: ' + pair.Percentile25G + '<br>75th Percentile: ' + pair.Percentile75G + '<br>Median Interval: ' + pair.MedianInterVisitIntervalG + ' '),"
    response.write "        textposition: 'auto',"
    response.write "        texttemplate: '%{y}',"
    response.write "        hovertemplate: '%{text}'"
    response.write "    };"
    
    response.write "    var data = [traceMale, traceFemale];"
    
    ' Define the layout for the chart
    response.write "    var barLayout = {"
    response.write "        title: 'Top 20 Inter Visit Interval  Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " ',"
    response.write "        xaxis: { title: 'Inter Visit Interval (Days)' },"
    response.write "        yaxis: { title: 'Frequency' },"
    response.write "        barmode: 'group',"
    response.write "        height: 600, width: window.innerWidth * 0.95,"
    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
'    response.write "        legend: { x: 0.1, y: 1.1, orientation: 'h' },"
     response.write "        legend: { orientation: 'h', y: -0.2, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    response.write "    };"

    ' Create the bar chart
    response.write "    Plotly.newPlot('monthlyChartDiv', data, barLayout);"
    response.write "});"
    response.write "</script>"

    ' DataTable Initialization
    response.write "<script>"
    response.write "    new DataTable('#monthlyTable', {"
    response.write "        data: dbDataYearly.data,"
    response.write "        columns: ["
    response.write "            { data: 'counterG' },"
    response.write "            { data: 'GenderG' },"
    response.write "            { data: 'InterVisitIntervalG' },"
    response.write "            { data: 'FrequencyFG' },"
    response.write "            { data: 'MaxInterVisitIntervalG' },"
    response.write "            { data: 'MinInterVisitIntervalG' },"
    response.write "            { data: 'AvgInterVisitIntervalG' },"
    response.write "            { data: 'Percentile25G' },"
    response.write "            { data: 'Percentile75G' },"
    response.write "            { data: 'MedianInterVisitIntervalG' }"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            }"
    response.write "        ]"
    response.write "    });"
    response.write "</script>"
End Sub



Sub get_inter_visit_intervals_by_age_group()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "WITH SelectCTE AS ("
    sql = sql & "    SELECT"
    sql = sql & "        dbo.GetAgeLabel(v.PatientAge) AS AgeGroup,"
    sql = sql & "        LAG(v.VisitDate) OVER (PARTITION BY v.PatientID ORDER BY v.VisitDate) AS PreviousVisitDate,"
    sql = sql & "        DATEDIFF(DAY, LAG(v.VisitDate) OVER (PARTITION BY v.PatientID ORDER BY v.VisitDate), v.VisitDate) AS InterVisitInterval"
    sql = sql & "    FROM Visitation v"
    sql = sql & "    JOIN Patient p ON p.PatientID = v.PatientID"
    sql = sql & "    WHERE CONVERT(DATE, v.VisitDate) BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    sql = sql & ")"
    sql = sql & "  SELECT TOP 20 "
    sql = sql & "    InterVisitInterval,"
    sql = sql & "    AgeGroup,"
    sql = sql & "    COUNT(*) AS Frequency,"
    sql = sql & "    CONVERT(VARCHAR, COUNT(*)) AS FrequencyF,"
    sql = sql & "    MAX(InterVisitInterval) OVER(PARTITION BY AgeGroup) AS MaxInterVisitInterval,"
    sql = sql & "    MIN(InterVisitInterval) OVER(PARTITION BY AgeGroup) AS MinInterVisitInterval,"
    sql = sql & "    AVG(InterVisitInterval) OVER(PARTITION BY AgeGroup) AS AvgInterVisitInterval,"
    sql = sql & "    PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY InterVisitInterval) OVER(PARTITION BY AgeGroup) AS Percentile25,"
    sql = sql & "    PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY InterVisitInterval) OVER(PARTITION BY AgeGroup) AS MedianInterVisitInterval,"
    sql = sql & "    PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY InterVisitInterval) OVER(PARTITION BY AgeGroup) AS Percentile75"
    sql = sql & "  FROM SelectCTE"
    sql = sql & " WHERE (InterVisitInterval IS NOT NULL) "
    sql = sql & " AND (CONVERT(INT, InterVisitInterval) > 0)"
    sql = sql & " GROUP BY InterVisitInterval, AgeGroup"
    sql = sql & " ORDER BY Frequency DESC, InterVisitInterval "

    rst.open sql, conn, 3, 4

    ' Generate JSON data
    Dim jsonDataAgeGroup, counter
    counter = 1
    jsonDataAgeGroup = "{""dataAgeGroup"":["
    
    ' Check if there are records
    If Not rst.EOF Then
        Do Until rst.EOF
            jsonDataAgeGroup = jsonDataAgeGroup & "{"
            jsonDataAgeGroup = jsonDataAgeGroup & """counter"":""" & CStr(counter) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """AgeGroup"":""" & CStr(rst.Fields("AgeGroup").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """InterVisitInterval"":""" & CStr(rst.Fields("InterVisitInterval").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """Frequency"":""" & CStr(rst.Fields("Frequency").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """FrequencyF"":""" & CStr(rst.Fields("FrequencyF").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """MaxInterVisitInterval"":""" & CStr(rst.Fields("MaxInterVisitInterval").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """MinInterVisitInterval"":""" & CStr(rst.Fields("MinInterVisitInterval").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """AvgInterVisitInterval"":""" & CStr(rst.Fields("AvgInterVisitInterval").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """Percentile25"":""" & CStr(rst.Fields("Percentile25").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """Percentile75"":""" & CStr(rst.Fields("Percentile75").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """MedianInterVisitInterval"":""" & CStr(rst.Fields("MedianInterVisitInterval").value) & """"
            jsonDataAgeGroup = jsonDataAgeGroup & "},"
            rst.MoveNext
            counter = counter + 1
        Loop
        jsonDataAgeGroup = Left(jsonDataAgeGroup, Len(jsonDataAgeGroup) - 1)   ' Remove the trailing comma
    End If

    jsonDataAgeGroup = jsonDataAgeGroup & "]}"

    rst.Close
    Set rst = Nothing

    ' Send the data to the client-side
    response.write "<!-- Spinner HTML -->"
    response.write "<div id='loadingSpinner' style='display: flex; align-items: center; justify-content: center; height: 100vh; width: 100%; position: fixed; top: 0; left: 0; background: rgba(255, 255, 255, 0.7); z-index: 9999;'>"
    response.write "<div style='text-align: center;'>"
    response.write "<div class='spinner' style='border: 8px solid #f3f3f3; border-top: 8px solid #3498db; border-radius: 50%; width: 50px; height: 50px; animation: spin 1s linear infinite;'></div>"
    response.write "<p style='font-size: 18px; color: #3498db; margin-top: 10px;'>Loading...</p>"
    response.write "</div>"
    response.write "</div>"

    response.write "<style>"
    response.write "/* Spinner animation */"
    response.write "@keyframes spin {"
    response.write "    0% { transform: rotate(0deg); }"
    response.write "    100% { transform: rotate(360deg); }"
    response.write "}"
    response.write "</style>"

    response.write "<script>"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var ageGroup = dbAgeGroup.dataAgeGroup;"
    response.write "    var colorPalette = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'];"

    response.write "    var ageGroups = [...new Set(ageGroup.map(d => d.AgeGroup))];"
    
    response.write "    var data = ageGroups.map((group, index) => {"
    response.write "        var filteredData = ageGroup.filter(d => d.AgeGroup === group);"
    response.write "        return {"
    response.write "            x: filteredData.map(d => d.InterVisitInterval + ' Days'),"
    response.write "            y: filteredData.map(d => d.Frequency),"
    response.write "            name: group,"
    response.write "            type: 'bar',"
    response.write "            marker: { color: colorPalette[index % colorPalette.length] },"
    response.write "            text: filteredData.map(d => 'Age Group: ' + d.AgeGroup + '<br>Inter Visit Interval: ' + d.InterVisitInterval + ' Days<br>Frequency: ' + d.Frequency + '<br>Max Interval: ' + d.MaxInterVisitInterval + '<br>Min Interval: ' + d.MinInterVisitInterval + '<br>Average Interval: ' + d.AvgInterVisitInterval + '<br>25th Percentile: ' + d.Percentile25 + '<br>75th Percentile: ' + d.Percentile75 + '<br>Median Interval: ' + d.MedianInterVisitInterval + ' '),"
    response.write "            textposition: 'auto',"
    response.write "            texttemplate: '%{y}',"
    response.write "            hovertemplate: '%{text}'"
    response.write "        };"
    response.write "    });"
    
    response.write "    var barLayout = {"
    response.write "        title: 'Top 20 Visit Interval By Age Group Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & "',"
    response.write "        xaxis: { title: 'Inter Visit Interval (Days)' },"
    response.write "        yaxis: { title: 'Frequency' },"
    response.write "        barmode: 'group',"
    response.write "        height: 600, width: window.innerWidth * 0.95,"
    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    response.write "        legend: { orientation: 'h', y: -0.2, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    response.write "    };"
    
    response.write "    Plotly.newPlot('weeklyVisitsChartDiv', data, barLayout).then(function() {"
    response.write "        document.getElementById('loadingSpinner').style.display = 'none';" ' Hide the spinner after chart is rendered
    response.write "    });"

    response.write "    $('#weeklyTable').DataTable({"
    response.write "        data: dbAgeGroup.dataAgeGroup,"
    response.write "        columns: ["
    response.write "            { data: 'counter', title: 'No' },"
    response.write "            { data: 'AgeGroup', title: 'Age Group' },"
    response.write "            { data: 'InterVisitInterval', title: 'Inter Visit Interval (Days)' },"
    response.write "            { data: 'Frequency', title: 'Frequency' },"
    response.write "            { data: 'FrequencyF', title: 'Frequency (Formatted)' },"
    response.write "            { data: 'MaxInterVisitInterval', title: 'Max Interval (Days)' },"
    response.write "            { data: 'MinInterVisitInterval', title: 'Min Interval (Days)' },"
    response.write "            { data: 'AvgInterVisitInterval', title: 'Average Interval (Days)' },"
    response.write "            { data: 'Percentile25', title: '25th Percentile' },"
    response.write "            { data: 'Percentile75', title: '75th Percentile' },"
    response.write "            { data: 'MedianInterVisitInterval', title: 'Median Interval (Days)' }"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            }"
    response.write "        ]"
    response.write "    });"

    response.write "});"
    response.write "</script>"

End Sub

Sub get_inter_visit_intervals_by_age_group2() ' was the working one
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "WITH SelectCTE AS ("
    sql = sql & "    SELECT"
    sql = sql & "        dbo.GetAgeLabel(v.PatientAge) AS AgeGroup,"
    sql = sql & "        LAG(v.VisitDate) OVER (PARTITION BY v.PatientID ORDER BY v.VisitDate) AS PreviousVisitDate,"
    sql = sql & "        DATEDIFF(DAY, LAG(v.VisitDate) OVER (PARTITION BY v.PatientID ORDER BY v.VisitDate), v.VisitDate) AS InterVisitInterval"
    sql = sql & "    FROM Visitation v"
    sql = sql & "    JOIN Patient p ON p.PatientID = v.PatientID"
    sql = sql & "    WHERE CONVERT(DATE, v.VisitDate) BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    sql = sql & ")"
    sql = sql & "  SELECT TOP 20 "
    sql = sql & "    InterVisitInterval,"
    sql = sql & "    AgeGroup,"
    sql = sql & "    COUNT(*) AS Frequency,"
    sql = sql & "    CONVERT(VARCHAR, COUNT(*)) AS FrequencyF,"
    sql = sql & "    MAX(InterVisitInterval) OVER(PARTITION BY AgeGroup) AS MaxInterVisitInterval,"
    sql = sql & "    MIN(InterVisitInterval) OVER(PARTITION BY AgeGroup) AS MinInterVisitInterval,"
    sql = sql & "    AVG(InterVisitInterval) OVER(PARTITION BY AgeGroup) AS AvgInterVisitInterval,"
    sql = sql & "    PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY InterVisitInterval) OVER(PARTITION BY AgeGroup) AS Percentile25,"
    sql = sql & "    PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY InterVisitInterval) OVER(PARTITION BY AgeGroup) AS MedianInterVisitInterval,"
    sql = sql & "    PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY InterVisitInterval) OVER(PARTITION BY AgeGroup) AS Percentile75"
    sql = sql & "  FROM SelectCTE"
    sql = sql & " WHERE (InterVisitInterval IS NOT NULL) "
    sql = sql & " AND (CONVERT(INT, InterVisitInterval) > 0)"
    sql = sql & " GROUP BY InterVisitInterval, AgeGroup"
    sql = sql & " ORDER BY Frequency DESC, InterVisitInterval "

    rst.open sql, conn, 3, 4

    ' Generate JSON data
    Dim jsonDataAgeGroup, counter
    counter = 1
    jsonDataAgeGroup = "{""dataAgeGroup"":["
    
    ' Check if there are records
    If Not rst.EOF Then
        Do Until rst.EOF
            jsonDataAgeGroup = jsonDataAgeGroup & "{"
            jsonDataAgeGroup = jsonDataAgeGroup & """counter"":""" & CStr(counter) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """AgeGroup"":""" & CStr(rst.Fields("AgeGroup").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """InterVisitInterval"":""" & CStr(rst.Fields("InterVisitInterval").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """Frequency"":""" & CStr(rst.Fields("Frequency").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """FrequencyF"":""" & CStr(rst.Fields("FrequencyF").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """MaxInterVisitInterval"":""" & CStr(rst.Fields("MaxInterVisitInterval").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """MinInterVisitInterval"":""" & CStr(rst.Fields("MinInterVisitInterval").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """AvgInterVisitInterval"":""" & CStr(rst.Fields("AvgInterVisitInterval").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """Percentile25"":""" & CStr(rst.Fields("Percentile25").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """Percentile75"":""" & CStr(rst.Fields("Percentile75").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """MedianInterVisitInterval"":""" & CStr(rst.Fields("MedianInterVisitInterval").value) & """"
            jsonDataAgeGroup = jsonDataAgeGroup & "},"
            rst.MoveNext
            counter = counter + 1
        Loop
        jsonDataAgeGroup = Left(jsonDataAgeGroup, Len(jsonDataAgeGroup) - 1)   ' Remove the trailing comma
    End If

    jsonDataAgeGroup = jsonDataAgeGroup & "]}"

    rst.Close
    Set rst = Nothing

    ' Send the data to the client-side
    response.write "<!-- Spinner HTML -->"
    response.write "<div id='loadingSpinner' style='display: flex; align-items: center; justify-content: center; height: 100vh; width: 100%; position: fixed; top: 0; left: 0; background: rgba(255, 255, 255, 0.7); z-index: 9999;'>"
    response.write "<div style='text-align: center;'>"
    response.write "<div class='spinner' style='border: 8px solid #f3f3f3; border-top: 8px solid #3498db; border-radius: 50%; width: 50px; height: 50px; animation: spin 1s linear infinite;'></div>"
    response.write "<p style='font-size: 18px; color: #3498db; margin-top: 10px;'>Loading...</p>"
    response.write "</div>"
    response.write "</div>"

    response.write "<style>"
    response.write "/* Spinner animation */"
    response.write "@keyframes spin {"
    response.write "    0% { transform: rotate(0deg); }"
    response.write "    100% { transform: rotate(360deg); }"
    response.write "}"
    response.write "</style>"

    response.write "<script>"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var ageGroup = dbAgeGroup.dataAgeGroup;"
    response.write "    var colorPalette = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'];"

    response.write "    var ageGroups = [...new Set(ageGroup.map(d => d.AgeGroup))];"
    
    response.write "    var data = ageGroups.map((group, index) => {"
    response.write "        var filteredData = ageGroup.filter(d => d.AgeGroup === group);"
    response.write "        return {"
    response.write "            x: filteredData.map(d => d.InterVisitInterval + ' Days'),"
    response.write "            y: filteredData.map(d => d.Frequency),"
    response.write "            name: group,"
    response.write "            type: 'bar',"
    response.write "            marker: { color: colorPalette[index % colorPalette.length] },"
    response.write "            text: filteredData.map(d => 'Age Group: ' + d.AgeGroup + '<br>Inter Visit Interval: ' + d.InterVisitInterval + ' Days<br>Frequency: ' + d.Frequency + '<br>Max Interval: ' + d.MaxInterVisitInterval + '<br>Min Interval: ' + d.MinInterVisitInterval + '<br>Average Interval: ' + d.AvgInterVisitInterval + '<br>25th Percentile: ' + d.Percentile25 + '<br>75th Percentile: ' + d.Percentile75 + '<br>Median Interval: ' + d.MedianInterVisitInterval + ' '),"
    response.write "            textposition: 'auto',"
    response.write "            texttemplate: '%{y}',"
    response.write "            hovertemplate: '%{text}'"
    response.write "        };"
    response.write "    });"
    
    response.write "    var barLayout = {"
    response.write "        title: 'Top 20 Visit Interval By Age Group Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & "',"
    response.write "        xaxis: { title: 'Inter Visit Interval (Days)' },"
    response.write "        yaxis: { title: 'Frequency' },"
    response.write "        barmode: 'group',"
    response.write "        height: 600, width: window.innerWidth * 0.95,"
    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    response.write "        legend: { orientation: 'h', y: -0.2, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    response.write "    };"
    
    response.write "    Plotly.newPlot('weeklyVisitsChartDiv', data, barLayout).then(function() {"
    response.write "        document.getElementById('loadingSpinner').style.display = 'none';" ' Hide the spinner after chart is rendered
    response.write "    });"

    response.write "    $('#weeklyTable').DataTable({"
    response.write "        data: dbAgeGroup.dataAgeGroup,"
    response.write "        columns: ["
    response.write "            { data: 'counter', title: 'No' },"
    response.write "            { data: 'AgeGroup', title: 'Age Group' },"
    response.write "            { data: 'InterVisitInterval', title: 'Inter Visit Interval (Days)' },"
    response.write "            { data: 'Frequency', title: 'Frequency' },"
    response.write "            { data: 'FrequencyF', title: 'Frequency (Formatted)' },"
    response.write "            { data: 'MaxInterVisitInterval', title: 'Max Interval (Days)' },"
    response.write "            { data: 'MinInterVisitInterval', title: 'Min Interval (Days)' },"
    response.write "            { data: 'AvgInterVisitInterval', title: 'Average Interval (Days)' },"
    response.write "            { data: 'Percentile25', title: '25th Percentile' },"
    response.write "            { data: 'Percentile75', title: '75th Percentile' },"
    response.write "            { data: 'MedianInterVisitInterval', title: 'Median Interval (Days)' }"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            }"
    response.write "        ]"
    response.write "    });"

    response.write "});"
    response.write "</script>"

End Sub

'==========================================================================

Sub get_inter_visit_intervals_by_age_group1()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "WITH SelectCTE AS ("
    sql = sql & "    SELECT"
    sql = sql & "        dbo.GetAgeLabel(v.PatientAge) AS AgeGroup,"
    sql = sql & "        LAG(v.VisitDate) OVER (PARTITION BY v.PatientID ORDER BY v.VisitDate) AS PreviousVisitDate,"
    sql = sql & "        DATEDIFF(DAY, LAG(v.VisitDate) OVER (PARTITION BY v.PatientID ORDER BY v.VisitDate), v.VisitDate) AS InterVisitInterval"
    sql = sql & "    FROM Visitation v"
    sql = sql & "    JOIN Patient p ON p.PatientID = v.PatientID"
    sql = sql & "    WHERE CONVERT(DATE, v.VisitDate) BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    sql = sql & ")"
    sql = sql & "  SELECT TOP 20 "
    sql = sql & "    InterVisitInterval,"
    sql = sql & "    AgeGroup,"
    sql = sql & "    COUNT(*) AS Frequency,"
    sql = sql & "    CONVERT(VARCHAR, COUNT(*)) AS FrequencyF,"
    sql = sql & "    MAX(InterVisitInterval) OVER(PARTITION BY AgeGroup) AS MaxInterVisitInterval,"
    sql = sql & "    MIN(InterVisitInterval) OVER(PARTITION BY AgeGroup) AS MinInterVisitInterval,"
    sql = sql & "    AVG(InterVisitInterval) OVER(PARTITION BY AgeGroup) AS AvgInterVisitInterval,"
    sql = sql & "    PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY InterVisitInterval) OVER(PARTITION BY AgeGroup) AS Percentile25,"
    sql = sql & "    PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY InterVisitInterval) OVER(PARTITION BY AgeGroup) AS MedianInterVisitInterval,"
    sql = sql & "    PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY InterVisitInterval) OVER(PARTITION BY AgeGroup) AS Percentile75"
    sql = sql & "  FROM SelectCTE"
    sql = sql & " WHERE (InterVisitInterval IS NOT NULL) "
    sql = sql & " AND (CONVERT(INT, InterVisitInterval) > 0)"
    sql = sql & " GROUP BY InterVisitInterval, AgeGroup"
    sql = sql & " ORDER BY Frequency DESC, InterVisitInterval "

    rst.open sql, conn, 3, 4

    ' Generate JSON data
    Dim jsonDataAgeGroup, counter
    counter = 1
    jsonDataAgeGroup = "{""dataAgeGroup"":["
    
    ' Check if the recordset has any records
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            jsonDataAgeGroup = jsonDataAgeGroup & "{"
            jsonDataAgeGroup = jsonDataAgeGroup & """counter"":""" & counter & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """InterVisitInterval"":""" & CStr(rst.Fields("InterVisitInterval").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """AgeGroup"":""" & rst.Fields("AgeGroup").value & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """Frequency"":""" & CStr(rst.Fields("Frequency").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """FrequencyF"":""" & CStr(rst.Fields("FrequencyF").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """MaxInterVisitInterval"":""" & CStr(rst.Fields("MaxInterVisitInterval").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """MinInterVisitInterval"":""" & CStr(rst.Fields("MinInterVisitInterval").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """AvgInterVisitInterval"":""" & CStr(rst.Fields("AvgInterVisitInterval").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """Percentile25"":""" & CStr(rst.Fields("Percentile25").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """Percentile75"":""" & CStr(rst.Fields("Percentile75").value) & ""","
            jsonDataAgeGroup = jsonDataAgeGroup & """MedianInterVisitInterval"":""" & CStr(rst.Fields("MedianInterVisitInterval").value) & """"
            jsonDataAgeGroup = jsonDataAgeGroup & "},"
            rst.MoveNext
            counter = counter + 1
        Loop
        jsonDataAgeGroup = Left(jsonDataAgeGroup, Len(jsonDataAgeGroup) - 1)   ' Remove the trailing comma
    End If

    jsonDataAgeGroup = jsonDataAgeGroup & "]}"

    rst.Close
    Set rst = Nothing

    ' Send the data to the client-side
    response.write "<script>"
    response.write "var dbAgeGroup = " & jsonDataAgeGroup & ";"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var ageGroup = dbAgeGroup.dataAgeGroup;"
    
'    ' Use a predefined color palette from Plotly
'    response.write "    var colorPalette = ['red', '#AADC32', '#FDE725','#35B779', '#482677', '#3E4A8E', '#31688E', '#26828E', '#1F9E89',  '#6DCD59'];"
    
    response.write "    var ageGroups = [...new Set(ageGroup.map(d => d.AgeGroup))];"
    
    ' Create a trace for each age group with colors from the palette
    response.write "    var data = ageGroups.map((group, index) => {"
    response.write "        var filteredData = ageGroup.filter(d => d.AgeGroup === group);"
    response.write "        return {"
    response.write "            x: filteredData.map(d => d.InterVisitInterval + ' Days'),"
    response.write "            y: filteredData.map(d => d.Frequency),"
    response.write "            name: group,"
    response.write "            type: 'bar',"
    response.write "            marker: { color: colorPalette[index % colorPalette.length] },"
    response.write "            text: filteredData.map(d => 'Age Group: ' + d.AgeGroup + '<br>Inter Visit Interval: ' + d.InterVisitInterval + ' Days<br>Frequency: ' + d.Frequency + '<br>Max Interval: ' + d.MaxInterVisitInterval + '<br>Min Interval: ' + d.MinInterVisitInterval + '<br>Average Interval: ' + d.AvgInterVisitInterval + '<br>25th Percentile: ' + d.Percentile25 + '<br>75th Percentile: ' + d.Percentile75 + '<br>Median Interval: ' + d.MedianInterVisitInterval + ' '),"
    response.write "            textposition: 'auto',"
    response.write "            texttemplate: '%{y}',"
    response.write "            hovertemplate: '%{text}'"
    response.write "        };"
    response.write "    });"
    
    ' Define the layout for the chart
    response.write "    var barLayout = {"
    response.write "        title: 'Top 20 Visit Interval By Age Group Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & "',"
    response.write "        xaxis: { title: 'Inter Visit Interval (Days)' },"
    response.write "        yaxis: { title: 'Frequency' },"
    response.write "        barmode: 'group',"
    response.write "        height: 600, width: window.innerWidth * 0.95,"
    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    response.write "        legend: { orientation: 'h', y: -0.2, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    response.write "    };"

    ' Create the bar chart
    response.write "    Plotly.newPlot('weeklyVisitsChartDiv', data, barLayout);"
    response.write "});"
    response.write "</script>"

    ' DataTable Initialization
    response.write "<script>"
    response.write "    $('#weeklyTable').DataTable({"
    response.write "        data: dbAgeGroup.dataAgeGroup,"
    response.write "        columns: ["
    response.write "            { data: 'counter', title: 'No' },"
    response.write "            { data: 'AgeGroup', title: 'Age Group' },"
    response.write "            { data: 'InterVisitInterval', title: 'Inter Visit Interval (Days)' },"
    response.write "            { data: 'Frequency', title: 'Frequency' },"
    response.write "            { data: 'FrequencyF', title: 'Frequency (Formatted)' },"
    response.write "            { data: 'MaxInterVisitInterval', title: 'Max Interval (Days)' },"
    response.write "            { data: 'MinInterVisitInterval', title: 'Min Interval (Days)' },"
    response.write "            { data: 'AvgInterVisitInterval', title: 'Average Interval (Days)' },"
    response.write "            { data: 'Percentile25', title: '25th Percentile' },"
    response.write "            { data: 'Percentile75', title: '75th Percentile' },"
    response.write "            { data: 'MedianInterVisitInterval', title: 'Median Interval (Days)' }"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            }"
    response.write "        ]"
    response.write "    });"
    response.write "</script>"

End Sub





'==========================================================================
' stacked bar
Sub get_inter_visit_intervals_by_gender1()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "WITH SelectCTE AS ("
    sql = sql & "    SELECT"
     sql = sql & " CASE p.genderid WHEN 'GEN01' THEN 'Male'"
     sql = sql & " WHEN 'GEN02' THEN 'Female'"
     sql = sql & " END AS Gender,"
    sql = sql & "        LAG(v.VisitDate) OVER (PARTITION BY v.PatientID ORDER BY v.VisitDate) AS PreviousVisitDate,"
    sql = sql & "        DATEDIFF(DAY, LAG(v.VisitDate) OVER (PARTITION BY v.PatientID ORDER BY v.VisitDate), v.VisitDate) AS InterVisitInterval"
    sql = sql & "    FROM Visitation v"
    sql = sql & "    JOIN Patient p ON p.PatientID = v.PatientID"
    sql = sql & "    WHERE CONVERT(DATE, v.VisitDate) BETWEEN ' " & periodStart & " ' AND ' " & periodEnd & " '"
    sql = sql & ")"
    sql = sql & "  SELECT TOP 20"
    sql = sql & "    InterVisitInterval,"
    sql = sql & "    COUNT(*) AS Frequency,"
    sql = sql & "    Gender,"
    sql = sql & "    FORMAT(COUNT(*), 'N0') AS FrequencyF,"
    sql = sql & "    MAX(InterVisitInterval) OVER() AS MaxInterVisitInterval,"
    sql = sql & "    MIN(InterVisitInterval) OVER() AS MinInterVisitInterval,"
    sql = sql & "    AVG(InterVisitInterval) OVER() AS AvgInterVisitInterval,"
    sql = sql & "    PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY InterVisitInterval) OVER() AS Percentile25,"
    sql = sql & "    PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY InterVisitInterval) OVER() AS MedianInterVisitInterval,"
    sql = sql & "    PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY InterVisitInterval) OVER() AS Percentile75"
    sql = sql & "  FROM SelectCTE"
    sql = sql & " WHERE (InterVisitInterval IS NOT NULL)"
    sql = sql & " AND (CONVERT(INT, InterVisitInterval) > 0)"
    sql = sql & " GROUP BY InterVisitInterval , Gender"
    sql = sql & " ORDER BY Frequency DESC"

 
    rst.open sql, conn, 3, 4

    ' Generate JSON data
    Dim jsonDatagender, counter
    counter = 1
    jsonDatagender = "{""data"":["

    ' Check if the recordset has any records
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            jsonDatagender = jsonDatagender & "{"
            jsonDatagender = jsonDatagender & """counterG"":""" & counter & ""","
            jsonDatagender = jsonDatagender & """InterVisitIntervalG"":""" & CStr(rst.Fields("InterVisitInterval").value) & ""","
            jsonDatagender = jsonDatagender & """GenderG"":""" & rst.Fields("Gender").value & ""","
            jsonDatagender = jsonDatagender & """FrequencyG"":""" & rst.Fields("Frequency").value & ""","
            jsonDatagender = jsonDatagender & """FrequencyFG"":""" & rst.Fields("FrequencyF").value & ""","
            jsonDatagender = jsonDatagender & """MaxInterVisitIntervalG"":""" & rst.Fields("MaxInterVisitInterval").value & ""","
            jsonDatagender = jsonDatagender & """MinInterVisitIntervalG"":""" & rst.Fields("MinInterVisitInterval").value & ""","
            jsonDatagender = jsonDatagender & """AvgInterVisitIntervalG"":""" & rst.Fields("AvgInterVisitInterval").value & ""","
            jsonDatagender = jsonDatagender & """Percentile25G"":""" & rst.Fields("Percentile25").value & ""","
            jsonDatagender = jsonDatagender & """Percentile75G"":""" & rst.Fields("Percentile75").value & ""","
            jsonDatagender = jsonDatagender & """MedianInterVisitIntervalG"":""" & rst.Fields("MedianInterVisitInterval").value & """"
            jsonDatagender = jsonDatagender & "},"
            rst.MoveNext
            counter = counter + 1
        Loop
        jsonDatagender = Left(jsonDatagender, Len(jsonDatagender) - 1) ' Remove the trailing comma
    End If

    jsonDatagender = jsonDatagender & "]}"

  
    rst.Close
    Set rst = Nothing

    ' Send the data to the client-side
    response.write "<script>"
    response.write "var dbDataYearly = " & jsonDatagender & ";"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var revenueSources = dbDataYearly.data;"
    
    ' Defining a color palette
    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    
    ' Define the chart data with different colors for each bar
    response.write "    var trace = {"
    response.write "        x: revenueSources.map(pair => pair.InterVisitIntervalG+' Days'),"
    response.write "        y: revenueSources.map(pair => pair.FrequencyG),"
    response.write "        type: 'bar',"
    response.write "        text: revenueSources.map(pair => ' Gender: ' + pair.GenderG + '  <br>Inter Visit Interval: ' + pair.InterVisitIntervalG + ' Days<br>Frequency: ' + pair.FrequencyFG + '<br>Max Interval: ' + pair.MaxInterVisitIntervalG + '<br>Min Interval: ' + pair.MinInterVisitIntervalG + '<br>Average Interval: ' + pair.AvgInterVisitIntervalG + '<br>25th Percentile: ' + pair.Percentile25G + '<br>75th Percentile: ' + pair.Percentile75G + '<br>Median Interval: ' + pair.MedianInterVisitIntervalG + ' '),"
    
    response.write "        textposition: 'auto',"
response.write "        texttemplate: '%{y}',"

    response.write "        hovertemplate: '%{text}',"
    response.write "        marker: {"
    response.write "            color: revenueSources.map((_, index) => colors[index % colors.length])"
    response.write "        }"
    response.write "    };"

    ' Define the layout for the chart
    response.write "    var barLayout = {"
    response.write "        title: 'Inter Visit Interval Analysis Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " ',"
    response.write "        xaxis: { title: 'Inter Visit Interval ' },"
    response.write "        yaxis: { title: 'Frequency' },"
    response.write "        height: 600, width: window.innerWidth * 0.95,"
    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    response.write "    };"

    ' Create the bar chart
    response.write "    Plotly.newPlot('monthlyChartDiv', [trace], barLayout);"
    response.write "});"
    response.write "</script>"

    ' DataTable Initialization
    response.write "<script>"
    response.write "    new DataTable('#monthlyTable', {"
    response.write "        data: dbDataYearly.data,"
    response.write "        columns: ["
    response.write "            { data: 'counterG' },"
    response.write "            { data: 'GenderG' },"
    response.write "            { data: 'InterVisitIntervalG' },"
    response.write "            { data: 'FrequencyFG' },"
    response.write "            { data: 'MaxInterVisitIntervalG' },"
    response.write "            { data: 'MinInterVisitIntervalG' },"
    response.write "            { data: 'AvgInterVisitIntervalG' },"
    response.write "            { data: 'Percentile25G' },"
    response.write "            { data: 'Percentile75G' },"
    response.write "            { data: 'MedianInterVisitIntervalG' }"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            }"
    response.write "        ]"
    response.write "    });"
    response.write "</script>"
End Sub
'==========================================================================

Sub InitPageScript()
  Dim htStr
  'Client Script
  htStr = ""
  htStr = htStr & "<script id=""scptPrintLayoutExtraScript"" LANGUAGE=""javascript"">" & vbCrLf
  htStr = htStr & vbCrLf
  'RefreshPage()
  htStr = htStr & "function RefreshPage(){" & vbCrLf
  htStr = htStr & "window.location.reload();" & vbCrLf
  htStr = htStr & "}" & vbCrLf


    htStr = htStr & "                    function getQueryParam(param) { "
    htStr = htStr & "                        let urlParams = new URLSearchParams(window.location.search); "
    htStr = htStr & "                        return urlParams.get(param); "
    htStr = htStr & "                    } "

   htStr = htStr & "          function formatDate(dateString) { "
  htStr = htStr & "                   dateString = String(dateString).trim(); "
   htStr = htStr & "                  var date = new Date(dateString); "
  htStr = htStr & "                   if (isNaN(date)) { "
  htStr = htStr & "                       console.error('Invalid date string'); "
   htStr = htStr & "                      return null; "
   htStr = htStr & "                  } "
   htStr = htStr & "                 var year = date.getFullYear(); "
    htStr = htStr & "                var month = String(date.getMonth() + 1).padStart(2, '0'); "
    htStr = htStr & "                var day = String(date.getDate()).padStart(2, '0'); "

   htStr = htStr & "                 return `${year}-${month}-${day}`; "
   htStr = htStr & "             } "

  htStr = htStr & "  document.addEventListener('DOMContentLoaded', (event) => { "

  htStr = htStr & "   var dropdown = document.getElementById('Branchs'); "
     htStr = htStr & "                 var defaultBranch = getQueryParam('branID'); "

htStr = htStr & "             for (var i = 0; i < dropdown.options.length; i++) { "
  htStr = htStr & "               if (dropdown.options[i].value === defaultBranch) { "
  htStr = htStr & "                   dropdown.selectedIndex = i; "
    htStr = htStr & "                 break; "
    htStr = htStr & "             } "
    htStr = htStr & "         } "
   htStr = htStr & "                 var selectedValue = getQueryParam('selectedValue'); "
   htStr = htStr & "                     if (selectedValue === null) {   "
    htStr = htStr & "          var startDateInput = document.getElementById('startDate'); "
    htStr = htStr & "          var today = new Date(); "
    htStr = htStr & "          var day = ('0' + today.getDate()).slice(-2); "
    htStr = htStr & "          var month = ('0' + (today.getMonth() + 1)).slice(-2); "
    htStr = htStr & "          var todayString = today.getFullYear() + '-' + month + '-' + day; "
    htStr = htStr & "          startDateInput.value = todayString; "
   htStr = htStr & "                     } else { "
   htStr = htStr & "                         var formattedDate = formatDate(selectedValue); "
   htStr = htStr & "                         if (formattedDate) { "
   htStr = htStr & "                         console.log('Formatted Date: ' + formattedDate); "
   htStr = htStr & "                          } else { "
   htStr = htStr & "                                    console.error('Invalid date string'); "
   htStr = htStr & "                        } "

      htStr = htStr & "          var startDateInput = document.getElementById('startDate'); "
           htStr = htStr & "          startDateInput.value = formattedDate; "
   htStr = htStr & "                     } "
   htStr = htStr & "                 var selectedValue1 = getQueryParam('selectedValue1'); "
   htStr = htStr & "                     if (selectedValue1 === null) { "
    htStr = htStr & "          var endDateInput = document.getElementById('endDate'); "
    htStr = htStr & "          var today = new Date(); "
    htStr = htStr & "          var day = ('0' + today.getDate()).slice(-2); "
    htStr = htStr & "          var month = ('0' + (today.getMonth() + 1)).slice(-2); "
    htStr = htStr & "          var todayString = today.getFullYear() + '-' + month + '-' + day; "
    htStr = htStr & "          endDateInput.value = todayString; "
   htStr = htStr & "                     } else { "
   htStr = htStr & "                         var formattedDate = formatDate(selectedValue1); "
   htStr = htStr & "                         if (formattedDate) { "
   htStr = htStr & "                         console.log('Formatted Date: ' + formattedDate); "
   htStr = htStr & "                          } else { "
   htStr = htStr & "                                    console.error('Invalid date string'); "
   htStr = htStr & "                        } "

      htStr = htStr & "          var endDateInput = document.getElementById('endDate'); "
           htStr = htStr & "          endDateInput.value = formattedDate; "
   htStr = htStr & "                     } "

  htStr = htStr & "      }); "

htStr = htStr & "var branchID1" & vbCrLf
htStr = htStr & "function MonthOnchange(){" & vbCrLf
htStr = htStr & "var ur, mth, sp, ordByTyp, ms, emr, str;" & vbCrLf
htStr = htStr & "mth =  document.getElementById('NoOfDays').value;" & vbCrLf
htStr = htStr & "dayid = 0;" & vbCrLf
htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=InterVisitIntervalAnalysis&PositionForTableName=WorkingDay';" & vbCrLf
htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&month=' + mth  + ' &yearid=' + dayid ;" & vbCrLf
htStr = htStr & "window.location.href = processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf
'emrMonth
htStr = htStr & "function BranchOnchange(){" & vbCrLf
htStr = htStr & "var ur, sp, ordByTyp, ms, emr, str;" & vbCrLf
htStr = htStr & "branchID1 =  document.getElementById('Branchs').value;" & vbCrLf
htStr = htStr & "dayid = 0;" & vbCrLf

htStr = htStr & "}" & vbCrLf

'emrDay
htStr = htStr & "function YearOnchange(){" & vbCrLf
htStr = htStr & "var ur, mth, sp, ordByTyp, ms, emr, str;" & vbCrLf
htStr = htStr & "dayid = GetEleVal('NoOfDay');" & vbCrLf
htStr = htStr & "mth = 0;" & vbCrLf
htStr = htStr & "emr=GetEleVal('emrdata');" & vbCrLf
htStr = htStr & "ispr= 0;" & vbCrLf
htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=InterVisitIntervalAnalysis&PositionForTableName=WorkingDay';" & vbCrLf
htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&month=' + mth  + ' &yearid=' + dayid ;" & vbCrLf
htStr = htStr & "window.location.href = processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf


  htStr = htStr & "function PeriodOnclick(){ " & vbCrLf
  htStr = htStr & "var branchID1 =  document.getElementById('Branchs').value;" & vbCrLf
  htStr = htStr & "var startDate1 =  document.getElementById('startDate').value;" & vbCrLf
  htStr = htStr & "var endDate1 =  document.getElementById('endDate').value;" & vbCrLf
  htStr = htStr & "startDate1 = startDate1.trimEnd();" & vbCrLf
  htStr = htStr & "endDate1 = endDate1.trimEnd();" & vbCrLf

  htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=InterVisitIntervalAnalysis&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&selectedValue=' + startDate1  + ' &selectedValue1=' + endDate1 + ' &branID=' + branchID1 ;" & vbCrLf


    htStr = htStr & "    window.location = ur;" & vbCrLf

  htStr = htStr & "}" & vbCrLf

  'emrdata
  htStr = htStr & "function OpenDiseaseFilterOnclick(){ " & vbCrLf
  htStr = htStr & "var branchID1 =  document.getElementById('Branchs').value;" & vbCrLf
  htStr = htStr & "var startDate1 =  document.getElementById('startDate').value;" & vbCrLf
  htStr = htStr & "var endDate1 =  document.getElementById('endDate').value;" & vbCrLf
  htStr = htStr & "startDate1 = startDate1.trimEnd();" & vbCrLf
  htStr = htStr & "endDate1 = endDate1.trimEnd();" & vbCrLf
'   htStr = htStr & "var branchID =  'B001';" & vbCrLf
  htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=InterVisitIntervalAnalysis&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&selectedValue=' + startDate1  + ' &selectedValue1=' + endDate1 + ' &branID=' + branchID1 ;" & vbCrLf
  htStr = htStr & "    window.location.href = processurl(ur);" & vbCrLf
'   htStr = htStr & "alert('This Functionality is under Maintenance');" & vbCrf
  htStr = htStr & "}" & vbCrLf
  
   
  
  
  'copy receipt number
    htStr = htStr & "function copyTextToClipboard(icon) {" & vbCrLf
    htStr = htStr & "const textToCopy = icon.parentElement.innerText.trim();" & vbCrLf
    htStr = htStr & "const textField = document.createElement('textarea');" & vbCrLf
    htStr = htStr & "textField.value = textToCopy;" & vbCrLf
    htStr = htStr & "document.body.appendChild(textField);" & vbCrLf
    htStr = htStr & "textField.select();" & vbCrLf
    htStr = htStr & "document.execCommand('copy');" & vbCrLf
    htStr = htStr & "document.body.removeChild(textField);" & vbCrLf
    htStr = htStr & "alert('Text copied to clipboard!');" & vbCrLf
    htStr = htStr & "};" & vbCrLf
'    htStr = htStr & " setTimeout(location.reload(),3000); " & vbCrLf
'    htStr = htStr & " console.log('wossop'); " & vbCrLf
    htStr = htStr & "</script>" & vbCrLf
  
  htStr = htStr & "</script>"
  response.write htStr
  js = js & "<script>" & vbCrLf
  js = js & "  " & vbCrLf
  js = js & "  " & vbCrLf
  js = js & "</script>"
  response.write js
End Sub


Sub filters()
'table for filters
    response.write "<table class = 'table table-bordered'>"
    response.write "        <tr '>"
    response.write "           <td> Facility: </td>   "
    response.write "            <td> "
    SetBranch
    response.write " </td>"

    response.write "           <td>From :</td>   "
    response.write "            <td><input type = 'date' id ='startDate'/>   </td>"
    response.write "           <td> To : </td>   "
    response.write "           <td> <input type = 'date'  id ='endDate'/> </td>    "
    response.write "           <td> <button  class='btn' style='background-color: #007bff; color: #fff' onclick=""PeriodOnclick()"" >Process</button> </td>   "
    
    response.write "    </table>"

End Sub

Sub SetBranch()
        Set rst = CreateObject("ADODB.Recordset")
        dyHt = "<select class='form-select' size=""1"" name=""Branchs"" id=""Branchs"" onchange=""BranchOnchange()"">"
        dyHt = dyHt & "<option value=""""></option>"
        ' dyHt = dyHt & "<option value='B001'></option>"

      
        sql0 = "select BranchID, BranchName from Branch b "
        With rst
          .open qryPro.FltQry(sql0), conn, 3, 4
          If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
              branchIDD = Trim(.Fields("BranchID"))
              branchName11 = Trim(.Fields("BranchName"))

              If UCase(CStr(yearId)) = UCase(branchIDD) Then
                dyHt = dyHt & "<option value=""" & CStr(branchIDD) & """ selected>" & branchName11 & "</option>"
              Else
                 dyHt = dyHt & "<option value=""" & CStr(branchIDD) & """>" & branchName11 & "</option>"
              End If
               rst.MoveNext
                response.flush
            Loop
          End If
          .Close
        End With
        dyHt = dyHt & "</select>"
        response.write dyHt
         Set rst = Nothing
End Sub


Function FormatDate(dateValue)
    FormatDate = year(dateValue) & "-" & Right("0" & month(dateValue), 2) & "-" & Right("0" & day(dateValue), 2)
End Function

'Function FormatDateNew(dateString)
'    Dim dateParts, yearPart, monthPart, dayPart, formatedDate
'    dateParts = Split(dateString, "-")
'    yearPart = dateParts(0)
'    monthPart = dateParts(1)
'    dayPart = dateParts(2)
'    formatedDate = dayPart & "/" & monthPart & "/" & yearPart
'    FormatDateNew = formatedDate
'End Function

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

Sub AddCss()

response.write "<style>"

response.write "/* ==== Data table  = Start ===== */"
response.write ""
response.write ".data_table{"
response.write "   background: #fff;"
response.write "    padding: 15px;"
response.write "    box-shadow: 1px 3px 5px #aaa;"
response.write "    border-radius: 5px;"
response.write "}"
response.write ""
response.write ".data_table .btn{"
response.write "    padding: 5px 10px;"
response.write "    margin: 10px 3px 10px 0;"
response.write "}"


response.write "    #previous, #next {"
response.write "        float: right;"
response.write "    }"

response.write "        #filterButton, #resetButton, #previous, #next {"
response.write "            border-radius: 5px; "
response.write "            background-color: #007bff; "
response.write "            color: #ffffff; "
response.write "            padding: 8px 12px; "
response.write "            margin-right: 10px; "
response.write "            border: none;"
response.write "            cursor: pointer;"
response.write "        }"
response.write "        #filterButton:hover, #resetButton:hover, #previous:hover, #next:hover {"
response.write "            background-color: #0056b3;"
response.write "        }"

response.write ".styled-input {"
response.write "    border: 1px solid #ccc;"
response.write "    border-radius: 8px;"
response.write "    padding: 8px;"
response.write "    margin-right: 10px;"
response.write "    font-size: 14px;"
response.write "    outline: none;"
response.write "}"
response.write ""
response.write ".styled-button {"
response.write "    background-color: #007bff;"
response.write "    color: white;"
response.write "    border: none;"
response.write "    border-radius: 8px;"
response.write "    padding: 8px 16px;"
response.write "    cursor: pointer;"
response.write "    font-size: 14px;"
response.write "}"
response.write ""
response.write ".styled-button:hover {"
response.write "    background-color: #0056b3;"
response.write "}"


response.write "</style>"


End Sub




'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>

