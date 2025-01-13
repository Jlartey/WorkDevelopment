'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

response.Clear
conn.commandTimeOut = 7200
Dim periodStart, periodEnd, brnchID
Dim page_title
page_title = ""

If Len(Trim(Request.QueryString("selectedValue"))) > 1 Then
    periodStart = Trim(Request.QueryString("selectedValue"))
    periodEnd = Trim(Request.QueryString("selectedValue1"))
    brnchID = Trim(Request.QueryString("branID"))

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
response.write " <link href=""https://cdn.datatables.net/v/bs5/jq-3.6.0/jszip-2.5.0/dt-1.13.5/af-2.6.0/b-2.4.0/b-colvis-2.4.0/b-html5-2.4.0/b-print-2.4.0/cr-1.7.0/date-1.5.0/fc-4.3.0/fh-3.4.0/kt-2.10.0/r-2.5.0/rg-1.4.0/rr-1.4.0/sc-2.2.0/sb-1.5.0/sp-2.2.0/sl-1.7.0/sr-1.3.0/datatables.min.css"" rel=""stylesheet""/>"
response.write " <script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js""></script>"
response.write " <script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js""></script>"
response.write " <script src=""https://cdn.datatables.net/v/bs5/jq-3.6.0/jszip-2.5.0/dt-1.13.5/af-2.6.0/b-2.4.0/b-colvis-2.4.0/b-html5-2.4.0/b-print-2.4.0/cr-1.7.0/date-1.5.0/fc-4.3.0/fh-3.4.0/kt-2.10.0/r-2.5.0/rg-1.4.0/rr-1.4.0/sc-2.2.0/sb-1.5.0/sp-2.2.0/sl-1.7.0/sr-1.3.0/datatables.min.js""></script>"

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
'response.write "  <div class='tab-button' onclick='openTab(event, ""yearlySamePeriodTab"")'>Annual Visits</div>"
response.write "  <div class='tab-button active' onclick='openTab(event, ""yearlyTab"")'>Annual Visits</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""quarterlyTab"")'>Quarterly Visits</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""monthlyTab"")'>Monthly Visits</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""weeklyTab"")'>Weekly Visitss</div>"

response.write "</div>"

'calling filters sub
filters

'yearly tab starts here

response.write "<div id='yearlyTab' class='tab-content active'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='yearlyChartDiv' class='chart'></div>"
response.write "  </div>"

' yearly table

  response.write "      <table style=""width:100%"" id=""yearlyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
  response.write "      <thead class=""table-dark"">"
  response.write "        <tr>"
  response.write "             <th>S/No.</th>"
  response.write "             <th>Patient Name</th>"
  response.write "             <th>Visit Type</th>"
  response.write "             <th>Visit Date</th>"
  response.write "             <th>Admission Status </th>"
  response.write "             <th>Admission Date</th>"
  response.write "             <th>Discharge Date</th>"
  response.write "             <th>Ward</th>"
  response.write "             <th>Bed</th>"  
  response.write "        </tr>"
  response.write "       </thead>"
  response.write "    </table>"
response.write "</div>"

'yearly tab end here

'quRTERly tab starts here
response.write "<div id='quarterlyTab' class='tab-content'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='quarterlyChartDiv' class='chart'></div>"
response.write "  </div>"

' quarterly table

   response.write "      <table style=""width:100%"" id=""quarterlyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    response.write "      <thead class=""table-dark"">"
    response.write "              <tr>"
     response.write "                <th>S/No.</th>"
    response.write "                <th>Year</th>"
     response.write "                <th>Quarter</th>"
      response.write "                <th>Age Group</th>"
    response.write "                <th>No. Of Visits</th>"
    response.write "                <th>Prev. No. Of Visits</th>"
    response.write "                <th>Difference</th>"
     response.write "                <th>QoQ % Change</th>"
    response.write "                <th>% Cont. To Age Group</th>"
    response.write "                <th>% To Annual Visits</th>"
     response.write "                <th>Cumulative Visits</th>"
    response.write "                <th>Overall Total</th>"
     response.write "                <th>% To Overall Visits</th>"
    
    response.write "                            </tr>"
    response.write "        </thead>"
    response.write "    </table>"



response.write "</div>"
'qurterly ends here
' monthly tab starts here
response.write "<div id='monthlyTab' class='tab-content'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='monthlyVisitsChartDiv' class='chart'></div>"
response.write "  </div>"
response.write "<br>"
response.write "    <div id='btnMonthDetails' ></div>"


' monthly table

   response.write "      <table style=""width:100%"" id=""monthlyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    response.write "      <thead class=""table-dark"">"
    response.write "              <tr>"
     response.write "                <th>S/No.</th>"
    response.write "                <th>Year</th>"
     response.write "                <th>Month</th>"
    response.write "                <th>Age Group</th>"
    response.write "                <th>No. Of Visits</th>"
    response.write "                <th>Prev No. Of Visits</th>"
    response.write "                <th>Difference</th>"
    response.write "                <th>% Change</th>"
     response.write "                <th>% Cont. To Age Group</th>"
    response.write "                <th>% Cont. To Annual Visits</th>"
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
    response.write "                <th>Year</th>"
     response.write "                <th>Week</th>"
    response.write "                <th>Age Group</th>"
    response.write "                <th>No. Of Visits</th>"
    response.write "                <th>Prev. No. Of Visits</th>"
    response.write "                <th>Difference</th>"
    response.write "                <th>% Change</th>"
    response.write "                <th>% Cont. To Age Group</th>"
     response.write "                <th>% Cont. To Annual Count</th>"
      response.write "                <th>Cumulative Weekly Count</th>"
    response.write "                <th>Overall Count</th>"
   response.write "                <th>% Cont. To Overall Count</th>"
    
    response.write "                            </tr>"
    response.write "        </thead>"
    response.write "    </table>"


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



'get_weekly_visits_analysis
''get_monthly_visits_analysis
''get_quarterly_visits_analysis
get_yearly_visits_analysis


Sub get_weekly_visits_analysis()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "SELECT"
    sql = sql & "   AgeGroup, [Year], [Week] , numOfVisits, PrevWeekCount, [Diff],"
    sql = sql & "   [%Change], [%ContToAgeGroup], [%ToAnnualTotal], [CumulativeWeeklyCounts],"
    sql = sql & "   OverallTotal, [%ToOverallTotal],"
    sql = sql & "   FORMAT(numOfVisits, 'N0') noOfVisitsF,"
    sql = sql & "   FORMAT([%Change], 'N2') [%ChangeF],"
    sql = sql & "   FORMAT([%ContToAgeGroup], 'N2') [%ContToAgeGroupF],"
    sql = sql & "   FORMAT([%ToAnnualTotal], 'N2') [%ToAnnualTotalF],"
    sql = sql & "   FORMAT([CumulativeWeeklyCounts], 'N0') [CumulativeWeeklyCountsF],"
    sql = sql & "   FORMAT(OverallTotal, 'N0') OverallTotalF,"
    sql = sql & "   FORMAT([%ToOverallTotal], 'N5') [%ToOverallTotalF]"
    sql = sql & " FROM [dbo].[fn_get_weekly_age_group_visits_analysis] "
    sql = sql & " ('" & periodStart & "', '" & periodEnd & "') "
    sql = sql & " where agegroup in ('70+','00-to-05','31-to-35') "
    sql = sql & " ORDER BY AgeGroup DESC, [Year] DESC, [Week]"

    rst.open sql, conn, 3, 4

    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """Year"":""" & rst.fields("Year").value & ""","
            jsonData = jsonData & """Week"":""" & rst.fields("Week").value & ""","
            jsonData = jsonData & """AgeGroup"":""" & rst.fields("AgeGroup").value & ""","
            jsonData = jsonData & """numOfVisits"":""" & rst.fields("numOfVisits").value & ""","
            jsonData = jsonData & """noOfVisitsF"":""" & rst.fields("noOfVisitsF").value & ""","
            jsonData = jsonData & """PrevWeekCount"":""" & rst.fields("PrevWeekCount").value & ""","
            jsonData = jsonData & """Diff"":""" & rst.fields("Diff").value & ""","
            jsonData = jsonData & """PercentageChange"":""" & rst.fields("%Change").value & ""","
            jsonData = jsonData & """PercentageChangeF"":""" & rst.fields("%ChangeF").value & ""","
            jsonData = jsonData & """PercentageContToAgeGroup"":""" & rst.fields("%ContToAgeGroup").value & ""","
            jsonData = jsonData & """PercentageContToAgeGroupF"":""" & rst.fields("%ContToAgeGroupF").value & ""","
            jsonData = jsonData & """PercentageToAnnualTotal"":""" & rst.fields("%ToAnnualTotal").value & ""","
            jsonData = jsonData & """PercentageToAnnualTotalF"":""" & rst.fields("%ToAnnualTotalF").value & ""","
            jsonData = jsonData & """CumulativeWeeklyCounts"":""" & rst.fields("CumulativeWeeklyCounts").value & ""","
            jsonData = jsonData & """CumulativeWeeklyCountsF"":""" & rst.fields("CumulativeWeeklyCountsF").value & ""","
            jsonData = jsonData & """OverallTotal"":""" & rst.fields("OverallTotal").value & ""","
            jsonData = jsonData & """OverallTotalF"":""" & rst.fields("OverallTotalF").value & ""","
            jsonData = jsonData & """PercentageToOverallTotal"":""" & rst.fields("%ToOverallTotal").value & ""","
            jsonData = jsonData & """PercentageToOverallTotalF"":""" & rst.fields("%ToOverallTotalF").value & """"
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
    response.write "var dbDataWeekly = " & jsonData & ";"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var visitsData = dbDataWeekly.data;"

    ' Extract unique age groups and prepare data for each age group
    response.write "    var ageGroups = [...new Set(visitsData.map(entry => entry.AgeGroup))];"
    response.write "    var years = [...new Set(visitsData.map(entry => entry.Year))];"
    response.write "    var traces = [];"

    ' Defining a color palette
    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    response.write "    var colorIndex = 0;"

    response.write "    ageGroups.forEach(function(ageGroup) {"
    response.write "        years.forEach(function(year) {"
    response.write "            var filteredData = visitsData.filter(entry => entry.AgeGroup === ageGroup && entry.Year === year);"
    response.write "            var trace = {"
    response.write "                x: filteredData.map(entry => entry.Week),"
    response.write "                y: filteredData.map(entry => parseFloat(entry.numOfVisits)),"
    response.write "                mode: 'lines+markers',"
    response.write "                name: ageGroup + ' - ' + year,"
    response.write "                text: filteredData.map(entry => 'Year: ' + entry.Year + '<br>Week: ' + entry.Week + '<br>Age Group: ' + entry.AgeGroup + '<br>No. of Visits: ' + entry.noOfVisitsF + '<br>Previous Week Count: ' + entry.PrevWeekCount + '<br>Percentage Change: ' + entry.PercentageChangeF + '<br>Contribution to Age Group: ' + entry.PercentageContToAgeGroupF + '<br>Contribution to Annual Total: ' + entry.PercentageToAnnualTotalF + '<br>Cumulative Weekly Counts: ' + entry.CumulativeWeeklyCountsF + '<br>Overall Total: ' + entry.OverallTotalF + '<br>Percentage to Overall Total: ' + entry.PercentageToOverallTotalF),"
    response.write "                hovertemplate: '%{text}',"
    response.write "                line: {"
    response.write "                    color: colors[colorIndex % colors.length]"
    response.write "                }"
    response.write "            };"
    response.write "            traces.push(trace);"
    response.write "            colorIndex++;"
    response.write "        });"
    response.write "    });"

    ' Layout for line chart
    response.write "    var lineLayout = {"
    response.write "        title: 'Weekly Visits Analysis By Age Group and Year',"
    response.write "        xaxis: { title: 'Week' },"
    response.write "        yaxis: { title: 'Number of Visits' },"
    response.write "        legend: { orientation: 'h', y: -0.4, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    response.write "        height: 600, width: window.innerWidth * 1.0"
    response.write "    };"

    ' Plot the line chart
    response.write "    Plotly.newPlot('weeklyVisitsChartDiv', traces, lineLayout);"

    response.write "});"
    response.write "</script>"

'    ' Weekly table starts here
    response.write "<script>"
    response.write "    new DataTable('#weeklyTable', {"
    response.write "        data: dbDataWeekly.data,"
    response.write "        columns: ["
    response.write "            { data: 'counter' },"
    response.write "            { data: 'Year' },"
    response.write "            { data: 'Week' },"
    response.write "            { data: 'AgeGroup' },"
    response.write "            { data: 'noOfVisitsF' },"
    response.write "            { data: 'PrevWeekCount' },"
    response.write "            { data: 'Diff' },"
    response.write "            { data: 'PercentageChangeF' },"
    response.write "            { data: 'PercentageContToAgeGroupF' },"
    response.write "            { data: 'PercentageToAnnualTotalF' },"
    response.write "            { data: 'CumulativeWeeklyCountsF' },"
    response.write "            { data: 'OverallTotalF' },"
    response.write "            { data: 'PercentageToOverallTotalF' }"
    response.write "        ],"


    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, ""All""]],"
    response.write "        dom: 'lBfrtip',"
    response.write "            search: {"
response.write "                smart: true"
response.write "                    },"

response.write "            buttons: ["
response.write "                {"
response.write "                    extend: 'csv',"
response.write "                    text: 'CSV',"
response.write "                    title: '" & brnchName & " Weekly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
response.write "                },"

response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: '" & brnchName & " Weekly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
response.write "                },"



response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: '" & brnchName & " Weekly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
response.write "                },"


response.write "                {"
response.write "                    extend: 'print',"
response.write "                    text: 'PRINT',"
response.write "                    title: '" & brnchName & " Weekly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & " '"
response.write "                }"

response.write "            ]"
    response.write "    });"
'    response.write "</script>"
    
    
    response.write "</script>"

End Sub



Sub get_monthly_visits_analysis()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
   
    sql = "SELECT AgeGroup, [Year], [Month], [MonthNumber], numOfVisits, PrevMonthCount, "
    sql = sql & "Diff, [PercentageChange], [PercentageContToAgeGroup], [PercentageToAnnualTotal], "
    sql = sql & "CumulativeMonthlyCounts, OverallTotal, [PercentageToOverallTotal], "
    sql = sql & "FORMAT(numOfVisits,'N0') AS numOfVisitsF, FORMAT(PrevMonthCount,'N0') AS PrevMonthCountF, "
    sql = sql & "FORMAT(Diff,'N0') AS DiffF, FORMAT([PercentageChange],'N2') AS PercentageChangeF, "
    sql = sql & "FORMAT([PercentageContToAgeGroup],'N2') AS PercentageContToAgeGroupF, "
    sql = sql & "FORMAT([PercentageToAnnualTotal],'N2') AS PercentageToAnnualTotalF, "
    sql = sql & "FORMAT(CumulativeMonthlyCounts,'N0') AS CumulativeMonthlyCountsF, "
    sql = sql & "FORMAT(OverallTotal,'N0') AS OverallTotalF, "
    sql = sql & "FORMAT([PercentageToOverallTotal],'N2') AS PercentageToOverallTotalF "
    sql = sql & "FROM [dbo].[fn_get_monthly_age_group_visits_analysis]('" & periodStart & "', '" & periodEnd & "') "
    sql = sql & " WHERE AgeGroup IN ('70+','00-to-05','31-to-35') "
    sql = sql & " ORDER BY AgeGroup DESC, [Year] DESC, [MonthNumber]"

    rst.open sql, conn, 3, 4

    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """Year"":""" & rst.fields("Year").value & ""","
            jsonData = jsonData & """Month"":""" & rst.fields("Month").value & ""","
            jsonData = jsonData & """MonthNumber"":""" & rst.fields("MonthNumber").value & ""","
            jsonData = jsonData & """AgeGroup"":""" & rst.fields("AgeGroup").value & ""","
            jsonData = jsonData & """numOfVisits"":""" & rst.fields("numOfVisits").value & ""","
            jsonData = jsonData & """numOfVisitsF"":""" & rst.fields("numOfVisitsF").value & ""","
            jsonData = jsonData & """PrevMonthCount"":""" & rst.fields("PrevMonthCount").value & ""","
            jsonData = jsonData & """PrevMonthCountF"":""" & rst.fields("PrevMonthCountF").value & ""","
            jsonData = jsonData & """Diff"":""" & rst.fields("Diff").value & ""","
            jsonData = jsonData & """DiffF"":""" & rst.fields("DiffF").value & ""","
            jsonData = jsonData & """PercentageChange"":""" & rst.fields("PercentageChange").value & ""","
            jsonData = jsonData & """PercentageChangeF"":""" & rst.fields("PercentageChangeF").value & ""","
            jsonData = jsonData & """PercentageContToAgeGroup"":""" & rst.fields("PercentageContToAgeGroup").value & ""","
            jsonData = jsonData & """PercentageContToAgeGroupF"":""" & rst.fields("PercentageContToAgeGroupF").value & ""","
            jsonData = jsonData & """PercentageToAnnualTotal"":""" & rst.fields("PercentageToAnnualTotal").value & ""","
            jsonData = jsonData & """PercentageToAnnualTotalF"":""" & rst.fields("PercentageToAnnualTotalF").value & ""","
            jsonData = jsonData & """CumulativeMonthlyCounts"":""" & rst.fields("CumulativeMonthlyCounts").value & ""","
            jsonData = jsonData & """CumulativeMonthlyCountsF"":""" & rst.fields("CumulativeMonthlyCountsF").value & ""","
            jsonData = jsonData & """OverallTotal"":""" & rst.fields("OverallTotal").value & ""","
            jsonData = jsonData & """OverallTotalF"":""" & rst.fields("OverallTotalF").value & ""","
            jsonData = jsonData & """PercentageToOverallTotal"":""" & rst.fields("PercentageToOverallTotal").value & ""","
            jsonData = jsonData & """PercentageToOverallTotalF"":""" & rst.fields("PercentageToOverallTotalF").value & """"
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
    response.write "var visitsData = dbDataMonthly.data;"

    ' Extract unique age groups and prepare data for each age group
    response.write "var ageGroups = [...new Set(visitsData.map(entry => entry.AgeGroup))];"
    response.write "var years = [...new Set(visitsData.map(entry => entry.Year))];"
    response.write "var traces = [];"

    ' Defining a color palette
    response.write "var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    response.write "var colorIndex = 0;"

    response.write "ageGroups.forEach(function(ageGroup) {"
    response.write "    years.forEach(function(year) {"
    response.write "        var filteredData = visitsData.filter(entry => entry.AgeGroup === ageGroup && entry.Year === year);"
    response.write "        var trace = {"
    response.write "            x: filteredData.map(entry => entry.Month),"
    response.write "            y: filteredData.map(entry => parseFloat(entry.numOfVisits)),"
    response.write "            mode: 'lines+markers',"
    response.write "            name: ageGroup + ' - ' + year,"
    response.write "            text: filteredData.map(entry => 'Year: ' + entry.Year + '<br>Month: ' + entry.Month + '<br>Age Group: ' + entry.AgeGroup + '<br>Number of Visits: ' + entry.numOfVisitsF + '<br>Previous Month Count: ' + entry.PrevMonthCountF + '<br>Percentage Change: ' + entry.PercentageChangeF + '%<br>Contribution to Age Group: ' + entry.PercentageContToAgeGroupF + '%<br>Contribution to Annual Total: ' + entry.PercentageToAnnualTotalF + '%<br>Cumulative Monthly Counts: ' + entry.CumulativeMonthlyCountsF + '<br>Overall Total: ' + entry.OverallTotalF + '<br>Percentage to Overall Total: ' + entry.PercentageToOverallTotalF +'%'),"
    response.write "            hovertemplate: '%{text}',"
    response.write "            line: {"
    response.write "                color: colors[colorIndex % colors.length]"
    response.write "            }"
    response.write "        };"
    response.write "        traces.push(trace);"
    response.write "        colorIndex++;"
    response.write "    });"
    response.write "});"

    ' Layout for line chart
    response.write "var lineLayout = {"
    response.write "    title: 'Monthly Visits Analysis By Age Group and Year',"
    response.write "    xaxis: { title: 'Month' },"
    response.write "    yaxis: { title: 'Number of Visits' },"
    response.write "    legend: { orientation: 'h', y: -0.4, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    response.write "    height: 600, width: window.innerWidth * 1.0"
    response.write "};"

    ' Plot the line chart
    response.write "Plotly.newPlot('monthlyVisitsChartDiv', traces, lineLayout);"

    ' DataTable Initialization
    response.write "new DataTable('#monthlyTable', {"
    response.write "    data: dbDataMonthly.data,"
    response.write "    columns: ["
    response.write "        { data: 'counter' },"
    response.write "        { data: 'Year' },"
    response.write "        { data: 'Month' },"
    response.write "        { data: 'AgeGroup' },"
    response.write "        { data: 'numOfVisitsF' },"
    response.write "        { data: 'PrevMonthCountF' },"
    response.write "        { data: 'DiffF' },"
    response.write "        { data: 'PercentageChangeF' },"
    response.write "        { data: 'PercentageContToAgeGroupF' },"
    response.write "        { data: 'PercentageToAnnualTotalF' },"
    response.write "        { data: 'CumulativeMonthlyCountsF' },"
    response.write "        { data: 'OverallTotalF' },"
    response.write "        { data: 'PercentageToOverallTotalF' }"
    response.write "    ],"
    response.write "    lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "    dom: 'lBfrtip',"
    response.write "    buttons: ["
    response.write "        {"
    response.write "            extend: 'csv',"
    response.write "            text: 'CSV',"
    response.write "            title: '" & brnchName & " Monthly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "        },"
    response.write "        {"
    response.write "            extend: 'excel',"
    response.write "            text: 'EXCEL',"
    response.write "            title: '" & brnchName & " Monthly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "        },"
    response.write "        {"
    response.write "            extend: 'pdf',"
    response.write "            text: 'PDF',"
    response.write "            title: '" & brnchName & " Monthly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "        },"
    response.write "        {"
    response.write "            extend: 'print',"
    response.write "            text: 'PRINT',"
    response.write "            title: '" & brnchName & " Monthly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "        }"
    response.write "    ]"
    response.write "});"
    response.write "</script>"

End Sub



Sub get_quarterly_visits_analysis()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")

   
    sql = "SELECT AgeGroup, [Year], [Quarter], numOfVisits, FORMAT(numOfVisits, 'N0') AS numOfVisitsF, "
    sql = sql & "FORMAT(PrevQuarterCount, 'N0') AS PrevQuarterCountF, FORMAT(Diff, 'N0') AS DiffF, "
    sql = sql & "FORMAT([PercentageChange], 'N2') AS PercentageChangeF, FORMAT([PercentageContToAgeGroup], 'N2') AS PercentageContToAgeGroupF, "
    sql = sql & "FORMAT([PercentageToAnnualTotal], 'N2') AS PercentageToAnnualTotalF, FORMAT(CumulativeQuarterlyCounts, 'N0') AS CumulativeQuarterlyCountsF, "
    sql = sql & "FORMAT(OverallTotal, 'N0') AS OverallTotalF, FORMAT([PercentageToOverallTotal], 'N2') AS PercentageToOverallTotalF, "
    sql = sql & "[Quarter] AS QuarterName, [PercentageToOverallTotal] AS PercentOfOverallTotal "
    sql = sql & "FROM [dbo].[fn_get_quarterly_age_group_visits_analysis]('" & periodStart & "', '" & periodEnd & "') "
    sql = sql & "WHERE AgeGroup IN ('70+', '00-to-05', '31-to-35') "
    sql = sql & "ORDER BY AgeGroup DESC, [Year] DESC, [Quarter]"

    rst.open sql, conn, 3, 4

    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """QuarterName"":""" & rst.fields("QuarterName").value & ""","
            jsonData = jsonData & """AgeGroup"":""" & rst.fields("AgeGroup").value & ""","
            jsonData = jsonData & """Year"":""" & rst.fields("Year").value & ""","
            jsonData = jsonData & """numOfVisits"":""" & rst.fields("numOfVisits").value & ""","
            jsonData = jsonData & """numOfVisitsF"":""" & rst.fields("numOfVisitsF").value & ""","
            jsonData = jsonData & """PrevQuarterCountF"":""" & rst.fields("PrevQuarterCountF").value & ""","
            jsonData = jsonData & """DiffF"":""" & rst.fields("DiffF").value & ""","
            jsonData = jsonData & """PercentageChangeF"":""" & rst.fields("PercentageChangeF").value & ""","
            jsonData = jsonData & """PercentageContToAgeGroupF"":""" & rst.fields("PercentageContToAgeGroupF").value & ""","
            jsonData = jsonData & """PercentageToAnnualTotalF"":""" & rst.fields("PercentageToAnnualTotalF").value & ""","
            jsonData = jsonData & """CumulativeQuarterlyCountsF"":""" & rst.fields("CumulativeQuarterlyCountsF").value & ""","
            jsonData = jsonData & """OverallTotalF"":""" & rst.fields("OverallTotalF").value & ""","
            jsonData = jsonData & """PercentOfOverallTotal"":""" & rst.fields("PercentOfOverallTotal").value & ""","
            jsonData = jsonData & """PercentageToOverallTotalF"":""" & rst.fields("PercentageToOverallTotalF").value & """"
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

    ' Creating a set of unique age groups
    response.write "    var ageGroups = [...new Set(revenueSourcesQuarterly.map(pair => pair.AgeGroup))];"
    response.write "    var traces = [];"

    ' Defining a color palette
    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    response.write "    var colorIndex = 0;"

    response.write "    ageGroups.forEach(function(ageGroup) {"
    response.write "        var filteredData = revenueSourcesQuarterly.filter(pair => pair.AgeGroup == ageGroup);"
    response.write "        var trace = {"
    response.write "            x: filteredData.map(pair => pair.QuarterName),"
    response.write "            y: filteredData.map(pair => parseFloat(pair.numOfVisits)),"
    response.write "            type: 'bar',"
    response.write "            name: ageGroup,"
    response.write "            text: filteredData.map(pair => 'Year: ' + pair.Year + '<br>Quarter: ' + pair.QuarterName + '<br>Age Group: ' + pair.AgeGroup + '<br>Number of Visits: ' + pair.numOfVisitsF + '<br>Previous Quarter Count: ' + pair.PrevQuarterCountF + '<br>Difference: ' + pair.DiffF + '<br>Percentage Change: ' + pair.PercentageChangeF + '%<br>Contribution to Age Group: ' + pair.PercentageContToAgeGroupF + '%<br>Contribution to Annual Total: ' + pair.PercentageToAnnualTotalF + '%<br>Cumulative Quarterly Counts: ' + pair.CumulativeQuarterlyCountsF + '<br>Overall Total: ' + pair.OverallTotalF + '<br>Percentage Of Overall Total: ' + pair.PercentageToOverallTotalF  + '%'),"
    response.write "            hovertemplate: '%{text}',"
    response.write "            marker: {"
    response.write "                color: colors[colorIndex % colors.length]"
    response.write "            }"
    response.write "        };"
    response.write "        traces.push(trace);"
    response.write "        colorIndex++;"
    response.write "    });"

    ' Layout for bar chart
    response.write "    var barLayout = {"
    response.write "        title: 'Quarterly Visits Analysis By Quarter and Age Group',"
    response.write "        xaxis: { title: 'Quarter' },"
    response.write "        yaxis: { title: 'Number of Visits' },"
    response.write "        legend: { orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    response.write "        height: 600, width: window.innerWidth * 1.0"
    response.write "    };"

    ' Plot the bar chart
    response.write "    Plotly.newPlot('quarterlyChartDiv', traces, barLayout);"
    response.write "});"
    response.write "</script>"

    ' DataTable Initialization
    response.write "<script>"
    response.write "    new DataTable('#quarterlyTable', {"
    response.write "        data: dbDataQuarterly.data,"
    response.write "        columns: ["
    response.write "            { data: 'counter' },"
    response.write "            { data: 'Year' },"
    response.write "            { data: 'QuarterName' },"
    response.write "            { data: 'AgeGroup' },"
    response.write "            { data: 'numOfVisitsF' },"
    response.write "            { data: 'PrevQuarterCountF' },"
    response.write "            { data: 'DiffF' },"
    response.write "            { data: 'PercentageChangeF' },"
    response.write "            { data: 'PercentageContToAgeGroupF' },"
    response.write "            { data: 'PercentageToAnnualTotalF' },"
    response.write "            { data: 'CumulativeQuarterlyCountsF' },"
    response.write "            { data: 'OverallTotalF' },"
    response.write "            { data: 'PercentageToOverallTotalF' }"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Quarterly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Quarterly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Quarterly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Quarterly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "            }"
    response.write "        ]"
    response.write "    });"
    response.write "</script>"
End Sub


Sub get_yearly_visits_analysis()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    

   
'    sql = "SELECT AgeGroup, [Year], numOfVisits, PrevYearCount, "
'    sql = sql & "Diff, [PercentageChange], [PercentageContToAgeGroup], [PercentageToAnnualTotal], "
'    sql = sql & "CumulativeYearlyCounts, OverallTotal, [PercentageToOverallTotal], "
'    sql = sql & "FORMAT(numOfVisits, 'N0') AS numOfVisitsF, "
'    sql = sql & "FORMAT(PrevYearCount, 'N0') AS PrevYearCountF, "
'    sql = sql & "FORMAT(Diff, 'N0') AS DiffF, FORMAT([PercentageChange], 'N2') AS PercentageChangeF, "
'    sql = sql & "FORMAT([PercentageContToAgeGroup], 'N2') AS PercentageContToAgeGroupF, "
'    sql = sql & "FORMAT([PercentageToAnnualTotal], 'N2') AS PercentageToAnnualTotalF, "
'    sql = sql & "FORMAT(CumulativeYearlyCounts, 'N0') AS CumulativeYearlyCountsF, "
'    sql = sql & "FORMAT(OverallTotal, 'N0') AS OverallTotalF, "
'    sql = sql & "FORMAT([PercentageToOverallTotal], 'N5') AS PercentageToOverallTotalF "
'    sql = sql & "FROM [dbo].[fn_get_yearly_age_group_visits_analysis]('" & periodStart & "','" & periodEnd & "') "
'    sql = sql & "WHERE AgeGroup IN ('70+', '00-to-05', '31-to-35') "
'    sql = sql & "ORDER BY AgeGroup DESC, [Year]"

   ' sql for populating table
    
    sql = "WITH PatientVisitCTE AS ( "
    sql = sql & "SELECT Patient.PatientID, PatientName, VisitationID, VisitType.VisitTypeName, CONVERT(DATE, Visitation.VisitDate) VisitDate "
    sql = sql & "FROM Patient JOIN Visitation ON Patient.PatientID = Visitation.PatientID JOIN VisitType ON Visitation.VisitTypeID = VisitType.VisitTypeID "

    
    ' Add date range filter
    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "WHERE CONVERT(DATE, Visitation.VisitDate) BETWEEN '2018-01-01' AND '2018-01-31' "
    Else
        sql = sql & "WHERE CONVERT(DATE, Visitation.VisitDate) BETWEEN '2018-01-01' AND '2018-01-31' "
    End If
'    sql = sql & "GROUP BY CONVERT(DATE, ConsultReviewDate)), "
    
    sql = sql & "), AdmissionCTE AS( "
    sql = sql & "SELECT AdmissionStatusName, CONVERT(DATETIME, AdmissionDate) AdmissionDate, CONVERT(DATETIME, DischargeDate) DischargeDate, PatientID, WardID, BedID  "
    sql = sql & "FROM Admission JOIN AdmissionStatus ON Admission.AdmissionStatusID = AdmissionStatus.AdmissionStatusID "
    
    ' Add date range filter for AdmissionCTE
    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "WHERE CONVERT(DATE, AdmissionDate) BETWEEN '2018-01-01' AND '2018-01-31' "
    Else
        sql = sql & "WHERE CONVERT(DATE, AdmissionDate) BETWEEN '2018-01-01' AND '2018-01-31' "
    End If
'    sql = sql & "GROUP BY CONVERT(DATE, DischargeDate)) "
    
    sql = sql & ") SELECT PatientVisitCTE.PatientID, PatientVisitCTE.PatientName, PatientVisitCTE.VisitationID, PatientVisitCTE.VisitTypeName, "
    sql = sql & "CONVERT(VARCHAR(20), PatientVisitCTE.VisitDate, 103) VisitDate, AdmissionCTE.AdmissionStatusName,  "
    sql = sql & "CONVERT(VARCHAR(20), AdmissionCTE.AdmissionDate, 103) AdmissionDate, CONVERT(VARCHAR(20), AdmissionCTE.DischargeDate, 103) DischargeDate,  "
    sql = sql & "Ward.WardName [Ward], Bed.BedName [Bed] FROM PatientVisitCTE JOIN AdmissionCTE ON PatientVisitCTE.PatientID = AdmissionCTE.PatientID "
    sql = sql & "JOIN Ward ON Ward.WardID = AdmissionCTE.WardID JOIN Bed ON Bed.BedID = AdmissionCTE.BedID "
    
    ' Incorporate selected treat types dynamically
'    If selectedWardIDs <> "" Then
'        sql = sql & "WHERE Ward.WardID IN (" & formattedIDs & ") "
'    Else
'        sql = sql & "WHERE Ward.WardName LIKE '%MATERNITY%' "
'    End If
'    sql = sql & "WHERE Ward.WardName LIKE '%MATERNITY%' "
    sql = sql & "ORDER BY VisitDate DESC, AdmissionDate DESC "
    
    ' sql for populating multiselect field
    sql2 = "SELECT WardID, WardName FROM Ward"
    If selectedWardIDs <> "" Then
        sql2 = sql2 & " WHERE WardID IN (" & formattedIDs & ")"
    End If

  
    rst.open sql, conn, 3, 4

    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """PatientName"":""" & rst.fields("PatientName").value & ""","
            jsonData = jsonData & """VisitTypeName"":""" & rst.fields("VisitTypeName").value & ""","
            jsonData = jsonData & """VisitDate"":""" & rst.fields("VisitDate").value & ""","
            jsonData = jsonData & """AdmissionStatusName"":""" & rst.fields("AdmissionStatusName").value & ""","
            jsonData = jsonData & """AdmissionDate"":""" & rst.fields("AdmissionDate").value & ""","
            jsonData = jsonData & """DischargeDate"":""" & rst.fields("DischargeDate").value & ""","
            jsonData = jsonData & """Ward"":""" & rst.fields("Ward").value & ""","
            jsonData = jsonData & """Bed"":""" & rst.fields("Bed").value & """"
'            jsonData = jsonData & """PercentageToAnnualTotalF"":""" & rst.fields("PercentageToAnnualTotalF").value & ""","
'            jsonData = jsonData & """CumulativeYearlyCountsF"":""" & rst.fields("CumulativeYearlyCountsF").value & ""","
'            jsonData = jsonData & """OverallTotalF"":""" & rst.fields("OverallTotalF").value & ""","
'            jsonData = jsonData & """PercentageToOverallTotalF"":""" & rst.fields("PercentageToOverallTotalF").value & """"
            jsonData = jsonData & "},"
            rst.MoveNext
            counter = counter + 1
            
        
            
            
            
        Loop
        jsonData = Left(jsonData, Len(jsonData) - 1) ' Remove the trailing comma
    End If

    jsonData = jsonData & "]}"

    rst.Close
    Set rst = Nothing
' ==========================================================================================================================
'    ' Send the data to the client-side
'    response.write "<script>"
'    response.write "var dbDataYearly = " & jsonData & ";"
'    response.write "document.addEventListener('DOMContentLoaded', function() {"
'    response.write "    var revenueSourcesYearly = dbDataYearly.data;"
'
'    ' Creating a set of unique age groups
'    response.write "    var ageGroups = [...new Set(revenueSourcesYearly.map(pair => pair.AgeGroup))];"
'    response.write "    var traces = [];"
'
'    ' Defining a color palette
'    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
'    response.write "    var colorIndex = 0;"
'
'    response.write "    ageGroups.forEach(function(ageGroup) {"
'    response.write "        var filteredData = revenueSourcesYearly.filter(pair => pair.AgeGroup == ageGroup);"
'    response.write "        var trace = {"
'    response.write "            x: filteredData.map(pair => pair.Year),"
'    response.write "            y: filteredData.map(pair => parseFloat(pair.numOfVisits)),"
'    response.write "            type: 'bar',"
'    response.write "            name: ageGroup,"
'    response.write "            text: filteredData.map(pair => 'Year: ' + pair.Year + '<br>Age Group: ' + pair.AgeGroup + '<br>Number of Visits: ' + pair.numOfVisitsF + '<br>Previous Year Count: ' + pair.PrevYearCountF + '<br>Difference: ' + pair.DiffF + '<br>Percentage Change: ' + pair.PercentageChangeF + '%<br>Contribution to Age Group: ' + pair.PercentageContToAgeGroupF + '%<br>Contribution to Annual Total: ' + pair.PercentageToAnnualTotalF + '%<br>Cumulative Yearly Counts: ' + pair.CumulativeYearlyCountsF + '<br>Overall Total: ' + pair.OverallTotalF + '<br>Percentage Of Overall Total: ' + pair.PercentageToOverallTotalF  + '%'),"
'    response.write "            hovertemplate: '%{text}',"
'    response.write "            marker: {"
'    response.write "                color: colors[colorIndex % colors.length]"
'    response.write "            }"
'    response.write "        };"
'    response.write "        traces.push(trace);"
'    response.write "        colorIndex++;"
'    response.write "    });"
'
'    ' Layout for bar chart
'    response.write "    var barLayout = {"
'    response.write "        title: 'Yearly Visits Analysis By Year and Age Group',"
'    response.write "        xaxis: { title: 'Year' },"
'    response.write "        yaxis: { title: 'Number of Visits' },"
'    response.write "        legend: { orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center', yanchor: 'top' },"
'    response.write "        height: 600, width: window.innerWidth * 1.0"
'    response.write "    };"
'
'    ' Plot the bar chart
'    response.write "    Plotly.newPlot('yearlyChartDiv', traces, barLayout);"
'    response.write "});"
'    response.write "</script>"

    ' DataTable Initialization
    response.write "<script>"
    response.write "var dbDataYearly = " & jsonData & ";"
    response.write "    new DataTable('#yearlyTable', {"
    response.write "        data: dbDataYearly.data,"
    response.write "        columns: ["
    response.write "            { data: 'counter' },"
    response.write "            { data: 'PatientName' },"
    response.write "            { data: 'VisitTypeName' },"
    response.write "            { data: 'VisitDate' },"
    response.write "            { data: 'AdmissionStatusName' },"
    response.write "            { data: 'AdmissionDate' },"
    response.write "            { data: 'DischargeDate' },"
    response.write "            { data: 'Ward' },"
    response.write "            { data: 'Bed' }"
    response.write "        ],"
    
    
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Yearly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Yearly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Yearly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Yearly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "            }"
    response.write "        ]"
    response.write "    });"
    response.write "</script>"
End Sub



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
htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=visitFrequencyAnalysis&PositionForTableName=WorkingDay';" & vbCrLf
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
htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitFrequencyAnalysis&PositionForTableName=WorkingDay';" & vbCrLf
htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&month=' + mth  + ' &yearid=' + dayid ;" & vbCrLf
htStr = htStr & "window.location.href = processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf


  htStr = htStr & "function PeriodOnclick(){ " & vbCrLf
  htStr = htStr & "var branchID1 =  document.getElementById('Branchs').value;" & vbCrLf
  htStr = htStr & "var startDate1 =  document.getElementById('startDate').value;" & vbCrLf
  htStr = htStr & "var endDate1 =  document.getElementById('endDate').value;" & vbCrLf
  htStr = htStr & "startDate1 = startDate1.trimEnd();" & vbCrLf
  htStr = htStr & "endDate1 = endDate1.trimEnd();" & vbCrLf

  htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=visitFrequencyAnalysis&PositionForTableName=WorkingDay';" & vbCrLf
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
  htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=visitFrequencyAnalysis&PositionForTableName=WorkingDay';" & vbCrLf
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
              branchIDD = Trim(.fields("BranchID"))
              branchName11 = Trim(.fields("BranchName"))

              If UCase(CStr(yearID)) = UCase(branchIDD) Then
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

Function FormatDateNew(dateString)
    Dim dateParts, yearPart, monthPart, dayPart, formatedDate
    dateParts = Split(dateString, "-")
    yearPart = dateParts(0)
    monthPart = dateParts(1)
    dayPart = dateParts(2)
    formatedDate = dayPart & "/" & monthPart & "/" & yearPart
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

