'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Response.Clear
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


Response.Write "<!DOCTYPE html>"
Response.Write "<html lang='en'>"
Response.Write "<head>"
Response.Write "<meta charset='UTF-8'>"
Response.Write "<meta name='viewport' content='width=device-width, initial-scale=1.0'>"
Response.Write "<title>Visitation  Analysis</title>"

Response.Write "<script src='https://cdn.plot.ly/plotly-latest.min.js'></script>"

Response.Write "    <link href=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css"" rel=""stylesheet"""
Response.Write "        integrity=""sha384-9ndCyUaIbzAi2FUVXJi0CjmCapSmO7SnpJef0486qhLnuZ2cdeRhO02iuK6FUUVM"" crossorigin=""anonymous"">"
Response.Write "    <script src=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"""
Response.Write "        integrity=""sha384-geWF76RCwLtnZ8qwWowPQNguL3RmwHVBC9FhGdlKrxdiJJigb/j/68SIy3Te4Bkz"""
Response.Write "        crossorigin=""anonymous""></script>"
Response.Write " <link href=""https://cdn.datatables.net/v/bs5/jq-3.6.0/jszip-2.5.0/dt-1.13.5/af-2.6.0/b-2.4.0/b-colvis-2.4.0/b-html5-2.4.0/b-print-2.4.0/cr-1.7.0/date-1.5.0/fc-4.3.0/fh-3.4.0/kt-2.10.0/r-2.5.0/rg-1.4.0/rr-1.4.0/sc-2.2.0/sb-1.5.0/sp-2.2.0/sl-1.7.0/sr-1.3.0/datatables.min.css"" rel=""stylesheet""/>"
Response.Write " <script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js""></script>"
Response.Write " <script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js""></script>"
Response.Write " <script src=""https://cdn.datatables.net/v/bs5/jq-3.6.0/jszip-2.5.0/dt-1.13.5/af-2.6.0/b-2.4.0/b-colvis-2.4.0/b-html5-2.4.0/b-print-2.4.0/cr-1.7.0/date-1.5.0/fc-4.3.0/fh-3.4.0/kt-2.10.0/r-2.5.0/rg-1.4.0/rr-1.4.0/sc-2.2.0/sb-1.5.0/sp-2.2.0/sl-1.7.0/sr-1.3.0/datatables.min.js""></script>"

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


'calling InitPageScript sub
InitPageScript

Response.Write "<script>"
Response.Write "function openTab(event, tabId) {"
Response.Write "  var i, tabcontent, tabbuttons;"
Response.Write "  tabcontent = document.getElementsByClassName('tab-content');"
Response.Write "  for (i = 0; i < tabcontent.length; i++) {"
Response.Write "    tabcontent[i].style.display = 'none';"
Response.Write "  }"
Response.Write "  tabbuttons = document.getElementsByClassName('tab-button');"
Response.Write "  for (i = 0; i < tabbuttons.length; i++) {"
Response.Write "    tabbuttons[i].className = tabbuttons[i].className.replace(' active', '');"
Response.Write "  }"
Response.Write "  document.getElementById(tabId).style.display = 'block';"
Response.Write "  event.currentTarget.className += ' active';"
Response.Write "}"
Response.Write "</script>"

Response.Write "<div class='tab-header'>"
'response.write "  <div class='tab-button' onclick='openTab(event, ""yearlySamePeriodTab"")'>Annual Visits</div>"
Response.Write "  <div class='tab-button active' onclick='openTab(event, ""yearlyTab"")'>Annual Visits</div>"
Response.Write "  <div class='tab-button' onclick='openTab(event, ""quarterlyTab"")'>Quarterly Visits</div>"
Response.Write "  <div class='tab-button' onclick='openTab(event, ""monthlyTab"")'>Monthly Visits</div>"
Response.Write "  <div class='tab-button' onclick='openTab(event, ""weeklyTab"")'>Weekly Visitss</div>"

Response.Write "</div>"

'calling filters sub
filters

'yearly tab starts here

Response.Write "<div id='yearlyTab' class='tab-content active'>"
Response.Write "  <div class='chart-container'>"
Response.Write "    <div id='yearlyChartDiv' class='chart'></div>"
Response.Write "  </div>"

' yearly table

   Response.Write "      <table style=""width:100%"" id=""yearlyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    Response.Write "      <thead class=""table-dark"">"
    Response.Write "              <tr>"
     Response.Write "                <th>S/No.</th>"
    Response.Write "                <th>Year</th>"
      Response.Write "                <th>Age Group</th>"
    Response.Write "                <th>No. Of Visits</th>"
    Response.Write "                <th>Prev. No. Of Visits</th>"
    Response.Write "                <th>Difference</th>"
     Response.Write "                <th>YoY % Change</th>"
    Response.Write "                <th>% Cont. To Age Group</th>"
    Response.Write "                <th>% To Annual Visits</th>"
     Response.Write "                <th>Cumulative Visits</th>"
    Response.Write "                <th>Overall Total</th>"
     Response.Write "                <th>% To Overall Visits</th>"
    
    Response.Write "                            </tr>"
    Response.Write "        </thead>"
    Response.Write "    </table>"
Response.Write "</div>"

'yearly tab end here

'quRTERly tab starts here
Response.Write "<div id='quarterlyTab' class='tab-content'>"
Response.Write "  <div class='chart-container'>"
Response.Write "    <div id='quarterlyChartDiv' class='chart'></div>"
Response.Write "  </div>"

' quarterly table

   Response.Write "      <table style=""width:100%"" id=""quarterlyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    Response.Write "      <thead class=""table-dark"">"
    Response.Write "              <tr>"
     Response.Write "                <th>S/No.</th>"
    Response.Write "                <th>Year</th>"
     Response.Write "                <th>Quarter</th>"
      Response.Write "                <th>Age Group</th>"
    Response.Write "                <th>No. Of Visits</th>"
    Response.Write "                <th>Prev. No. Of Visits</th>"
    Response.Write "                <th>Difference</th>"
     Response.Write "                <th>QoQ % Change</th>"
    Response.Write "                <th>% Cont. To Age Group</th>"
    Response.Write "                <th>% To Annual Visits</th>"
     Response.Write "                <th>Cumulative Visits</th>"
    Response.Write "                <th>Overall Total</th>"
     Response.Write "                <th>% To Overall Visits</th>"
    
    Response.Write "                            </tr>"
    Response.Write "        </thead>"
    Response.Write "    </table>"



Response.Write "</div>"
'qurterly ends here
' monthly tab starts here
Response.Write "<div id='monthlyTab' class='tab-content'>"
Response.Write "  <div class='chart-container'>"
Response.Write "    <div id='monthlyVisitsChartDiv' class='chart'></div>"
Response.Write "  </div>"
Response.Write "<br>"
Response.Write "    <div id='btnMonthDetails' ></div>"


' monthly table

   Response.Write "      <table style=""width:100%"" id=""monthlyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    Response.Write "     <thead class=""table-dark"">"
    Response.Write "              <tr>"
     Response.Write "                <th>S/No.</th>"
    Response.Write "                <th>Year</th>"
     Response.Write "                <th>Month</th>"
    Response.Write "                <th>Age Group</th>"
    Response.Write "                <th>No. Of Visits</th>"
    Response.Write "                <th>Prev No. Of Visits</th>"
    Response.Write "                <th>Difference</th>"
    Response.Write "                <th>% Change</th>"
     Response.Write "                <th>% Cont. To Age Group</th>"
    Response.Write "                <th>% Cont. To Annual Visits</th>"
     Response.Write "                <th>Cumulative Monthly Visits </th>"
      Response.Write "                <th>Overall Visits</th>"
    Response.Write "                <th>% To Overall Visits</th>"
    
    Response.Write "     </tr>"
    Response.Write "       </thead>"
    Response.Write "    </table>"


Response.Write "</div>"
'monthly tab ends here

'weekly tab starts here
Response.Write "<div id='weeklyTab' class='tab-content'>"
Response.Write "  <div class='chart-container'>"
Response.Write "    <div id='weeklyVisitsChartDiv' class='chart'></div>"
Response.Write "  </div>"

' weekly table

   Response.Write "      <table style=""width:100%"" id=""weeklyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    Response.Write "      <thead class=""table-dark"">"
    Response.Write "              <tr>"
     Response.Write "                <th>S/No.</th>"
    Response.Write "                <th>Year</th>"
     Response.Write "                <th>Week</th>"
    Response.Write "                <th>Age Group</th>"
    Response.Write "                <th>No. Of Visits</th>"
    Response.Write "                <th>Prev. No. Of Visits</th>"
    Response.Write "                <th>Difference</th>"
    Response.Write "                <th>% Change</th>"
    Response.Write "                <th>% Cont. To Age Group</th>"
     Response.Write "                <th>% Cont. To Annual Count</th>"
      Response.Write "                <th>Cumulative Weekly Count</th>"
    Response.Write "                <th>Overall Count</th>"
   Response.Write "                <th>% Cont. To Overall Count</th>"
    
    Response.Write "                            </tr>"
    Response.Write "        </thead>"
    Response.Write "    </table>"


Response.Write "</div>"
'weekly tab ends here

'same period revenue-yearly
'response.write "<div id='yearlySamePeriodTab' class='tab-content'>"
'response.write "  <div class='chart-container'>"
'response.write "    <div id='yearlySamePeriodChartDiv' class='chart'></div>"
'response.write "  </div>"
'
'response.write "</div>"

Response.Write "</body>"
Response.Write "</html>"



get_weekly_visits_analysis
get_monthly_visits_analysis
get_quarterly_visits_analysis
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
            jsonData = jsonData & """Year"":""" & rst.Fields("Year").value & ""","
            jsonData = jsonData & """Week"":""" & rst.Fields("Week").value & ""","
            jsonData = jsonData & """AgeGroup"":""" & rst.Fields("AgeGroup").value & ""","
            jsonData = jsonData & """numOfVisits"":""" & rst.Fields("numOfVisits").value & ""","
            jsonData = jsonData & """noOfVisitsF"":""" & rst.Fields("noOfVisitsF").value & ""","
            jsonData = jsonData & """PrevWeekCount"":""" & rst.Fields("PrevWeekCount").value & ""","
            jsonData = jsonData & """Diff"":""" & rst.Fields("Diff").value & ""","
            jsonData = jsonData & """PercentageChange"":""" & rst.Fields("%Change").value & ""","
            jsonData = jsonData & """PercentageChangeF"":""" & rst.Fields("%ChangeF").value & ""","
            jsonData = jsonData & """PercentageContToAgeGroup"":""" & rst.Fields("%ContToAgeGroup").value & ""","
            jsonData = jsonData & """PercentageContToAgeGroupF"":""" & rst.Fields("%ContToAgeGroupF").value & ""","
            jsonData = jsonData & """PercentageToAnnualTotal"":""" & rst.Fields("%ToAnnualTotal").value & ""","
            jsonData = jsonData & """PercentageToAnnualTotalF"":""" & rst.Fields("%ToAnnualTotalF").value & ""","
            jsonData = jsonData & """CumulativeWeeklyCounts"":""" & rst.Fields("CumulativeWeeklyCounts").value & ""","
            jsonData = jsonData & """CumulativeWeeklyCountsF"":""" & rst.Fields("CumulativeWeeklyCountsF").value & ""","
            jsonData = jsonData & """OverallTotal"":""" & rst.Fields("OverallTotal").value & ""","
            jsonData = jsonData & """OverallTotalF"":""" & rst.Fields("OverallTotalF").value & ""","
            jsonData = jsonData & """PercentageToOverallTotal"":""" & rst.Fields("%ToOverallTotal").value & ""","
            jsonData = jsonData & """PercentageToOverallTotalF"":""" & rst.Fields("%ToOverallTotalF").value & """"
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
    Response.Write "<script>"
    Response.Write "var dbDataWeekly = " & jsonData & ";"
    Response.Write "document.addEventListener('DOMContentLoaded', function() {"
    Response.Write "    var visitsData = dbDataWeekly.data;"

    ' Extract unique age groups and prepare data for each age group
    Response.Write "    var ageGroups = [...new Set(visitsData.map(entry => entry.AgeGroup))];"
    Response.Write "    var years = [...new Set(visitsData.map(entry => entry.Year))];"
    Response.Write "    var traces = [];"

    ' Defining a color palette
    Response.Write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    Response.Write "    var colorIndex = 0;"

    Response.Write "    ageGroups.forEach(function(ageGroup) {"
    Response.Write "        years.forEach(function(year) {"
    Response.Write "            var filteredData = visitsData.filter(entry => entry.AgeGroup === ageGroup && entry.Year === year);"
    Response.Write "            var trace = {"
    Response.Write "                x: filteredData.map(entry => entry.Week),"
    Response.Write "                y: filteredData.map(entry => parseFloat(entry.numOfVisits)),"
    Response.Write "                mode: 'lines+markers',"
    Response.Write "                name: ageGroup + ' - ' + year,"
    Response.Write "                text: filteredData.map(entry => 'Year: ' + entry.Year + '<br>Week: ' + entry.Week + '<br>Age Group: ' + entry.AgeGroup + '<br>No. of Visits: ' + entry.noOfVisitsF + '<br>Previous Week Count: ' + entry.PrevWeekCount + '<br>Percentage Change: ' + entry.PercentageChangeF + '<br>Contribution to Age Group: ' + entry.PercentageContToAgeGroupF + '<br>Contribution to Annual Total: ' + entry.PercentageToAnnualTotalF + '<br>Cumulative Weekly Counts: ' + entry.CumulativeWeeklyCountsF + '<br>Overall Total: ' + entry.OverallTotalF + '<br>Percentage to Overall Total: ' + entry.PercentageToOverallTotalF),"
    Response.Write "                hovertemplate: '%{text}',"
    Response.Write "                line: {"
    Response.Write "                    color: colors[colorIndex % colors.length]"
    Response.Write "                }"
    Response.Write "            };"
    Response.Write "            traces.push(trace);"
    Response.Write "            colorIndex++;"
    Response.Write "        });"
    Response.Write "    });"

    ' Layout for line chart
    Response.Write "    var lineLayout = {"
    Response.Write "        title: 'Weekly Visits Analysis By Age Group and Year',"
    Response.Write "        xaxis: { title: 'Week' },"
    Response.Write "        yaxis: { title: 'Number of Visits' },"
    Response.Write "        legend: { orientation: 'h', y: -0.4, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    Response.Write "        height: 600, width: window.innerWidth * 1.0"
    Response.Write "    };"

    ' Plot the line chart
    Response.Write "    Plotly.newPlot('weeklyVisitsChartDiv', traces, lineLayout);"

    Response.Write "});"
    Response.Write "</script>"

'    ' Weekly table starts here
    Response.Write "<script>"
    Response.Write "    new DataTable('#weeklyTable', {"
    Response.Write "        data: dbDataWeekly.data,"
    Response.Write "        columns: ["
    Response.Write "            { data: 'counter' },"
    Response.Write "            { data: 'Year' },"
    Response.Write "            { data: 'Week' },"
    Response.Write "            { data: 'AgeGroup' },"
    Response.Write "            { data: 'noOfVisitsF' },"
    Response.Write "            { data: 'PrevWeekCount' },"
    Response.Write "            { data: 'Diff' },"
    Response.Write "            { data: 'PercentageChangeF' },"
    Response.Write "            { data: 'PercentageContToAgeGroupF' },"
    Response.Write "            { data: 'PercentageToAnnualTotalF' },"
    Response.Write "            { data: 'CumulativeWeeklyCountsF' },"
    Response.Write "            { data: 'OverallTotalF' },"
    Response.Write "            { data: 'PercentageToOverallTotalF' }"
    Response.Write "        ],"


    Response.Write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, ""All""]],"
    Response.Write "        dom: 'lBfrtip',"
    Response.Write "            search: {"
Response.Write "                smart: true"
Response.Write "                    },"

Response.Write "            buttons: ["
Response.Write "                {"
Response.Write "                    extend: 'csv',"
Response.Write "                    text: 'CSV',"
Response.Write "                    title: '" & brnchName & " Weekly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
Response.Write "                },"

Response.Write "                {"
Response.Write "                    extend: 'excel',"
Response.Write "                    text: 'EXCEL',"
Response.Write "                    title: '" & brnchName & " Weekly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
Response.Write "                },"



Response.Write "                {"
Response.Write "                    extend: 'pdf',"
Response.Write "                    text: 'PDF',"
Response.Write "                    title: '" & brnchName & " Weekly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
Response.Write "                },"


Response.Write "                {"
Response.Write "                    extend: 'print',"
Response.Write "                    text: 'PRINT',"
Response.Write "                    title: '" & brnchName & " Weekly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & " '"
Response.Write "                }"

Response.Write "            ]"
    Response.Write "    });"
'    response.write "</script>"
    
    
    Response.Write "</script>"

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
            jsonData = jsonData & """Year"":""" & rst.Fields("Year").value & ""","
            jsonData = jsonData & """Month"":""" & rst.Fields("Month").value & ""","
            jsonData = jsonData & """MonthNumber"":""" & rst.Fields("MonthNumber").value & ""","
            jsonData = jsonData & """AgeGroup"":""" & rst.Fields("AgeGroup").value & ""","
            jsonData = jsonData & """numOfVisits"":""" & rst.Fields("numOfVisits").value & ""","
            jsonData = jsonData & """numOfVisitsF"":""" & rst.Fields("numOfVisitsF").value & ""","
            jsonData = jsonData & """PrevMonthCount"":""" & rst.Fields("PrevMonthCount").value & ""","
            jsonData = jsonData & """PrevMonthCountF"":""" & rst.Fields("PrevMonthCountF").value & ""","
            jsonData = jsonData & """Diff"":""" & rst.Fields("Diff").value & ""","
            jsonData = jsonData & """DiffF"":""" & rst.Fields("DiffF").value & ""","
            jsonData = jsonData & """PercentageChange"":""" & rst.Fields("PercentageChange").value & ""","
            jsonData = jsonData & """PercentageChangeF"":""" & rst.Fields("PercentageChangeF").value & ""","
            jsonData = jsonData & """PercentageContToAgeGroup"":""" & rst.Fields("PercentageContToAgeGroup").value & ""","
            jsonData = jsonData & """PercentageContToAgeGroupF"":""" & rst.Fields("PercentageContToAgeGroupF").value & ""","
            jsonData = jsonData & """PercentageToAnnualTotal"":""" & rst.Fields("PercentageToAnnualTotal").value & ""","
            jsonData = jsonData & """PercentageToAnnualTotalF"":""" & rst.Fields("PercentageToAnnualTotalF").value & ""","
            jsonData = jsonData & """CumulativeMonthlyCounts"":""" & rst.Fields("CumulativeMonthlyCounts").value & ""","
            jsonData = jsonData & """CumulativeMonthlyCountsF"":""" & rst.Fields("CumulativeMonthlyCountsF").value & ""","
            jsonData = jsonData & """OverallTotal"":""" & rst.Fields("OverallTotal").value & ""","
            jsonData = jsonData & """OverallTotalF"":""" & rst.Fields("OverallTotalF").value & ""","
            jsonData = jsonData & """PercentageToOverallTotal"":""" & rst.Fields("PercentageToOverallTotal").value & ""","
            jsonData = jsonData & """PercentageToOverallTotalF"":""" & rst.Fields("PercentageToOverallTotalF").value & """"
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
    Response.Write "<script>"
    Response.Write "var dbDataMonthly = " & jsonData & ";"
    Response.Write "var visitsData = dbDataMonthly.data;"

    ' Extract unique age groups and prepare data for each age group
    Response.Write "var ageGroups = [...new Set(visitsData.map(entry => entry.AgeGroup))];"
    Response.Write "var years = [...new Set(visitsData.map(entry => entry.Year))];"
    Response.Write "var traces = [];"

    ' Defining a color palette
    Response.Write "var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    Response.Write "var colorIndex = 0;"

    Response.Write "ageGroups.forEach(function(ageGroup) {"
    Response.Write "    years.forEach(function(year) {"
    Response.Write "        var filteredData = visitsData.filter(entry => entry.AgeGroup === ageGroup && entry.Year === year);"
    Response.Write "        var trace = {"
    Response.Write "            x: filteredData.map(entry => entry.Month),"
    Response.Write "            y: filteredData.map(entry => parseFloat(entry.numOfVisits)),"
    Response.Write "            mode: 'lines+markers',"
    Response.Write "            name: ageGroup + ' - ' + year,"
    Response.Write "            text: filteredData.map(entry => 'Year: ' + entry.Year + '<br>Month: ' + entry.Month + '<br>Age Group: ' + entry.AgeGroup + '<br>Number of Visits: ' + entry.numOfVisitsF + '<br>Previous Month Count: ' + entry.PrevMonthCountF + '<br>Percentage Change: ' + entry.PercentageChangeF + '%<br>Contribution to Age Group: ' + entry.PercentageContToAgeGroupF + '%<br>Contribution to Annual Total: ' + entry.PercentageToAnnualTotalF + '%<br>Cumulative Monthly Counts: ' + entry.CumulativeMonthlyCountsF + '<br>Overall Total: ' + entry.OverallTotalF + '<br>Percentage to Overall Total: ' + entry.PercentageToOverallTotalF +'%'),"
    Response.Write "            hovertemplate: '%{text}',"
    Response.Write "            line: {"
    Response.Write "                color: colors[colorIndex % colors.length]"
    Response.Write "            }"
    Response.Write "        };"
    Response.Write "        traces.push(trace);"
    Response.Write "        colorIndex++;"
    Response.Write "    });"
    Response.Write "});"

    ' Layout for line chart
    Response.Write "var lineLayout = {"
    Response.Write "    title: 'Monthly Visits Analysis By Age Group and Year',"
    Response.Write "    xaxis: { title: 'Month' },"
    Response.Write "    yaxis: { title: 'Number of Visits' },"
    Response.Write "    legend: { orientation: 'h', y: -0.4, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    Response.Write "    height: 600, width: window.innerWidth * 1.0"
    Response.Write "};"

    ' Plot the line chart
    Response.Write "Plotly.newPlot('monthlyVisitsChartDiv', traces, lineLayout);"

    ' DataTable Initialization
    Response.Write "new DataTable('#monthlyTable', {"
    Response.Write "    data: dbDataMonthly.data,"
    Response.Write "    columns: ["
    Response.Write "        { data: 'counter' },"
    Response.Write "        { data: 'Year' },"
    Response.Write "        { data: 'Month' },"
    Response.Write "        { data: 'AgeGroup' },"
    Response.Write "        { data: 'numOfVisitsF' },"
    Response.Write "        { data: 'PrevMonthCountF' },"
    Response.Write "        { data: 'DiffF' },"
    Response.Write "        { data: 'PercentageChangeF' },"
    Response.Write "        { data: 'PercentageContToAgeGroupF' },"
    Response.Write "        { data: 'PercentageToAnnualTotalF' },"
    Response.Write "        { data: 'CumulativeMonthlyCountsF' },"
    Response.Write "        { data: 'OverallTotalF' },"
    Response.Write "        { data: 'PercentageToOverallTotalF' }"
    Response.Write "    ],"
    Response.Write "    lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    Response.Write "    dom: 'lBfrtip',"
    Response.Write "    buttons: ["
    Response.Write "        {"
    Response.Write "            extend: 'csv',"
    Response.Write "            text: 'CSV',"
    Response.Write "            title: '" & brnchName & " Monthly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "        },"
    Response.Write "        {"
    Response.Write "            extend: 'excel',"
    Response.Write "            text: 'EXCEL',"
    Response.Write "            title: '" & brnchName & " Monthly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "        },"
    Response.Write "        {"
    Response.Write "            extend: 'pdf',"
    Response.Write "            text: 'PDF',"
    Response.Write "            title: '" & brnchName & " Monthly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "        },"
    Response.Write "        {"
    Response.Write "            extend: 'print',"
    Response.Write "            text: 'PRINT',"
    Response.Write "            title: '" & brnchName & " Monthly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "        }"
    Response.Write "    ]"
    Response.Write "});"
    Response.Write "</script>"

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
            jsonData = jsonData & """QuarterName"":""" & rst.Fields("QuarterName").value & ""","
            jsonData = jsonData & """AgeGroup"":""" & rst.Fields("AgeGroup").value & ""","
            jsonData = jsonData & """Year"":""" & rst.Fields("Year").value & ""","
            jsonData = jsonData & """numOfVisits"":""" & rst.Fields("numOfVisits").value & ""","
            jsonData = jsonData & """numOfVisitsF"":""" & rst.Fields("numOfVisitsF").value & ""","
            jsonData = jsonData & """PrevQuarterCountF"":""" & rst.Fields("PrevQuarterCountF").value & ""","
            jsonData = jsonData & """DiffF"":""" & rst.Fields("DiffF").value & ""","
            jsonData = jsonData & """PercentageChangeF"":""" & rst.Fields("PercentageChangeF").value & ""","
            jsonData = jsonData & """PercentageContToAgeGroupF"":""" & rst.Fields("PercentageContToAgeGroupF").value & ""","
            jsonData = jsonData & """PercentageToAnnualTotalF"":""" & rst.Fields("PercentageToAnnualTotalF").value & ""","
            jsonData = jsonData & """CumulativeQuarterlyCountsF"":""" & rst.Fields("CumulativeQuarterlyCountsF").value & ""","
            jsonData = jsonData & """OverallTotalF"":""" & rst.Fields("OverallTotalF").value & ""","
            jsonData = jsonData & """PercentOfOverallTotal"":""" & rst.Fields("PercentOfOverallTotal").value & ""","
            jsonData = jsonData & """PercentageToOverallTotalF"":""" & rst.Fields("PercentageToOverallTotalF").value & """"
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
    Response.Write "<script>"
    Response.Write "var dbDataQuarterly = " & jsonData & ";"
    Response.Write "document.addEventListener('DOMContentLoaded', function() {"
    Response.Write "    var revenueSourcesQuarterly = dbDataQuarterly.data;"

    ' Creating a set of unique age groups
    Response.Write "    var ageGroups = [...new Set(revenueSourcesQuarterly.map(pair => pair.AgeGroup))];"
    Response.Write "    var traces = [];"

    ' Defining a color palette
    Response.Write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    Response.Write "    var colorIndex = 0;"

    Response.Write "    ageGroups.forEach(function(ageGroup) {"
    Response.Write "        var filteredData = revenueSourcesQuarterly.filter(pair => pair.AgeGroup == ageGroup);"
    Response.Write "        var trace = {"
    Response.Write "            x: filteredData.map(pair => pair.QuarterName),"
    Response.Write "            y: filteredData.map(pair => parseFloat(pair.numOfVisits)),"
    Response.Write "            type: 'bar',"
    Response.Write "            name: ageGroup,"
    Response.Write "            text: filteredData.map(pair => 'Year: ' + pair.Year + '<br>Quarter: ' + pair.QuarterName + '<br>Age Group: ' + pair.AgeGroup + '<br>Number of Visits: ' + pair.numOfVisitsF + '<br>Previous Quarter Count: ' + pair.PrevQuarterCountF + '<br>Difference: ' + pair.DiffF + '<br>Percentage Change: ' + pair.PercentageChangeF + '%<br>Contribution to Age Group: ' + pair.PercentageContToAgeGroupF + '%<br>Contribution to Annual Total: ' + pair.PercentageToAnnualTotalF + '%<br>Cumulative Quarterly Counts: ' + pair.CumulativeQuarterlyCountsF + '<br>Overall Total: ' + pair.OverallTotalF + '<br>Percentage Of Overall Total: ' + pair.PercentageToOverallTotalF  + '%'),"
    Response.Write "            hovertemplate: '%{text}',"
    Response.Write "            marker: {"
    Response.Write "                color: colors[colorIndex % colors.length]"
    Response.Write "            }"
    Response.Write "        };"
    Response.Write "        traces.push(trace);"
    Response.Write "        colorIndex++;"
    Response.Write "    });"

    ' Layout for bar chart
    Response.Write "    var barLayout = {"
    Response.Write "        title: 'Quarterly Visits Analysis By Quarter and Age Group',"
    Response.Write "        xaxis: { title: 'Quarter' },"
    Response.Write "        yaxis: { title: 'Number of Visits' },"
    Response.Write "        legend: { orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    Response.Write "        height: 600, width: window.innerWidth * 1.0"
    Response.Write "    };"

    ' Plot the bar chart
    Response.Write "    Plotly.newPlot('quarterlyChartDiv', traces, barLayout);"
    Response.Write "});"
    Response.Write "</script>"

    ' DataTable Initialization
    Response.Write "<script>"
    Response.Write "    new DataTable('#quarterlyTable', {"
    Response.Write "        data: dbDataQuarterly.data,"
    Response.Write "        columns: ["
    Response.Write "            { data: 'counter' },"
    Response.Write "            { data: 'Year' },"
    Response.Write "            { data: 'QuarterName' },"
    Response.Write "            { data: 'AgeGroup' },"
    Response.Write "            { data: 'numOfVisitsF' },"
    Response.Write "            { data: 'PrevQuarterCountF' },"
    Response.Write "            { data: 'DiffF' },"
    Response.Write "            { data: 'PercentageChangeF' },"
    Response.Write "            { data: 'PercentageContToAgeGroupF' },"
    Response.Write "            { data: 'PercentageToAnnualTotalF' },"
    Response.Write "            { data: 'CumulativeQuarterlyCountsF' },"
    Response.Write "            { data: 'OverallTotalF' },"
    Response.Write "            { data: 'PercentageToOverallTotalF' }"
    Response.Write "        ],"
    Response.Write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    Response.Write "        dom: 'lBfrtip',"
    Response.Write "        buttons: ["
    Response.Write "            {"
    Response.Write "                extend: 'csv',"
    Response.Write "                text: 'CSV',"
    Response.Write "                title: '" & brnchName & " Quarterly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'excel',"
    Response.Write "                text: 'EXCEL',"
    Response.Write "                title: '" & brnchName & " Quarterly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'pdf',"
    Response.Write "                text: 'PDF',"
    Response.Write "                title: '" & brnchName & " Quarterly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'print',"
    Response.Write "                text: 'PRINT',"
    Response.Write "                title: '" & brnchName & " Quarterly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            }"
    Response.Write "        ]"
    Response.Write "    });"
    Response.Write "</script>"
End Sub


Sub get_yearly_visits_analysis()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")

   
    sql = "SELECT AgeGroup, [Year], numOfVisits, PrevYearCount, "
    sql = sql & "Diff, [PercentageChange], [PercentageContToAgeGroup], [PercentageToAnnualTotal], "
    sql = sql & "CumulativeYearlyCounts, OverallTotal, [PercentageToOverallTotal], "
    sql = sql & "FORMAT(numOfVisits, 'N0') AS numOfVisitsF, "
    sql = sql & "FORMAT(PrevYearCount, 'N0') AS PrevYearCountF, "
    sql = sql & "FORMAT(Diff, 'N0') AS DiffF, FORMAT([PercentageChange], 'N2') AS PercentageChangeF, "
    sql = sql & "FORMAT([PercentageContToAgeGroup], 'N2') AS PercentageContToAgeGroupF, "
    sql = sql & "FORMAT([PercentageToAnnualTotal], 'N2') AS PercentageToAnnualTotalF, "
    sql = sql & "FORMAT(CumulativeYearlyCounts, 'N0') AS CumulativeYearlyCountsF, "
    sql = sql & "FORMAT(OverallTotal, 'N0') AS OverallTotalF, "
    sql = sql & "FORMAT([PercentageToOverallTotal], 'N5') AS PercentageToOverallTotalF "
    sql = sql & "FROM [dbo].[fn_get_yearly_age_group_visits_analysis]('" & periodStart & "','" & periodEnd & "') "
    sql = sql & "WHERE AgeGroup IN ('70+', '00-to-05', '31-to-35') "
    sql = sql & "ORDER BY AgeGroup DESC, [Year]"

  
    rst.open sql, conn, 3, 4

    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """AgeGroup"":""" & rst.Fields("AgeGroup").value & ""","
            jsonData = jsonData & """Year"":""" & rst.Fields("Year").value & ""","
            jsonData = jsonData & """numOfVisits"":""" & rst.Fields("numOfVisits").value & ""","
            jsonData = jsonData & """numOfVisitsF"":""" & rst.Fields("numOfVisitsF").value & ""","
            jsonData = jsonData & """PrevYearCountF"":""" & rst.Fields("PrevYearCountF").value & ""","
            jsonData = jsonData & """DiffF"":""" & rst.Fields("DiffF").value & ""","
            jsonData = jsonData & """PercentageChangeF"":""" & rst.Fields("PercentageChangeF").value & ""","
            jsonData = jsonData & """PercentageContToAgeGroupF"":""" & rst.Fields("PercentageContToAgeGroupF").value & ""","
            jsonData = jsonData & """PercentageToAnnualTotalF"":""" & rst.Fields("PercentageToAnnualTotalF").value & ""","
            jsonData = jsonData & """CumulativeYearlyCountsF"":""" & rst.Fields("CumulativeYearlyCountsF").value & ""","
            jsonData = jsonData & """OverallTotalF"":""" & rst.Fields("OverallTotalF").value & ""","
            jsonData = jsonData & """PercentageToOverallTotalF"":""" & rst.Fields("PercentageToOverallTotalF").value & """"
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
    Response.Write "<script>"
    Response.Write "var dbDataYearly = " & jsonData & ";"
    Response.Write "document.addEventListener('DOMContentLoaded', function() {"
    Response.Write "    var revenueSourcesYearly = dbDataYearly.data;"

    ' Creating a set of unique age groups
    Response.Write "    var ageGroups = [...new Set(revenueSourcesYearly.map(pair => pair.AgeGroup))];"
    Response.Write "    var traces = [];"

    ' Defining a color palette
    Response.Write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    Response.Write "    var colorIndex = 0;"

    Response.Write "    ageGroups.forEach(function(ageGroup) {"
    Response.Write "        var filteredData = revenueSourcesYearly.filter(pair => pair.AgeGroup == ageGroup);"
    Response.Write "        var trace = {"
    Response.Write "            x: filteredData.map(pair => pair.Year),"
    Response.Write "            y: filteredData.map(pair => parseFloat(pair.numOfVisits)),"
    Response.Write "            type: 'bar',"
    Response.Write "            name: ageGroup,"
    Response.Write "            text: filteredData.map(pair => 'Year: ' + pair.Year + '<br>Age Group: ' + pair.AgeGroup + '<br>Number of Visits: ' + pair.numOfVisitsF + '<br>Previous Year Count: ' + pair.PrevYearCountF + '<br>Difference: ' + pair.DiffF + '<br>Percentage Change: ' + pair.PercentageChangeF + '%<br>Contribution to Age Group: ' + pair.PercentageContToAgeGroupF + '%<br>Contribution to Annual Total: ' + pair.PercentageToAnnualTotalF + '%<br>Cumulative Yearly Counts: ' + pair.CumulativeYearlyCountsF + '<br>Overall Total: ' + pair.OverallTotalF + '<br>Percentage Of Overall Total: ' + pair.PercentageToOverallTotalF  + '%'),"
    Response.Write "            hovertemplate: '%{text}',"
    Response.Write "            marker: {"
    Response.Write "                color: colors[colorIndex % colors.length]"
    Response.Write "            }"
    Response.Write "        };"
    Response.Write "        traces.push(trace);"
    Response.Write "        colorIndex++;"
    Response.Write "    });"

    ' Layout for bar chart
    Response.Write "    var barLayout = {"
    Response.Write "        title: 'Yearly Visits Analysis By Year and Age Group',"
    Response.Write "        xaxis: { title: 'Year' },"
    Response.Write "        yaxis: { title: 'Number of Visits' },"
    Response.Write "        legend: { orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    Response.Write "        height: 600, width: window.innerWidth * 1.0"
    Response.Write "    };"

    ' Plot the bar chart
    Response.Write "    Plotly.newPlot('yearlyChartDiv', traces, barLayout);"
    Response.Write "});"
    Response.Write "</script>"

    ' DataTable Initialization
    Response.Write "<script>"
    Response.Write "    new DataTable('#yearlyTable', {"
    Response.Write "        data: dbDataYearly.data,"
    Response.Write "        columns: ["
    Response.Write "            { data: 'counter' },"
    Response.Write "            { data: 'Year' },"
    Response.Write "            { data: 'AgeGroup' },"
    Response.Write "            { data: 'numOfVisitsF' },"
    Response.Write "            { data: 'PrevYearCountF' },"
    Response.Write "            { data: 'DiffF' },"
    Response.Write "            { data: 'PercentageChangeF' },"
    Response.Write "            { data: 'PercentageContToAgeGroupF' },"
    Response.Write "            { data: 'PercentageToAnnualTotalF' },"
    Response.Write "            { data: 'CumulativeYearlyCountsF' },"
    Response.Write "            { data: 'OverallTotalF' },"
    Response.Write "            { data: 'PercentageToOverallTotalF' }"
    Response.Write "        ],"
    Response.Write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    Response.Write "        dom: 'lBfrtip',"
    Response.Write "        buttons: ["
    Response.Write "            {"
    Response.Write "                extend: 'csv',"
    Response.Write "                text: 'CSV',"
    Response.Write "                title: '" & brnchName & " Yearly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'excel',"
    Response.Write "                text: 'EXCEL',"
    Response.Write "                title: '" & brnchName & " Yearly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'pdf',"
    Response.Write "                text: 'PDF',"
    Response.Write "                title: '" & brnchName & " Yearly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'print',"
    Response.Write "                text: 'PRINT',"
    Response.Write "                title: '" & brnchName & " Yearly Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            }"
    Response.Write "        ]"
    Response.Write "    });"
    Response.Write "</script>"
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
  Response.Write htStr
  js = js & "<script>" & vbCrLf
  js = js & "  " & vbCrLf
  js = js & "  " & vbCrLf
  js = js & "</script>"
  Response.Write js
End Sub


Sub filters()
'table for filters
    Response.Write "<table class = 'table table-bordered'>"
    Response.Write "        <tr '>"
    Response.Write "           <td> Facility: </td>   "
    Response.Write "            <td> "
    SetBranch
    Response.Write " </td>"

    Response.Write "           <td>From :</td>   "
    Response.Write "            <td><input type = 'date' id ='startDate'/>   </td>"
    Response.Write "           <td> To : </td>   "
    Response.Write "           <td> <input type = 'date'  id ='endDate'/> </td>    "
    Response.Write "           <td> <button  class='btn' style='background-color: #007bff; color: #fff' onclick=""PeriodOnclick()"" >Process</button> </td>   "
    
    Response.Write "    </table>"

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
                Response.flush
            Loop
          End If
          .Close
        End With
        dyHt = dyHt & "</select>"
        Response.Write dyHt
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

Response.Write "<style>"

Response.Write "/* ==== Data table  = Start ===== */"
Response.Write ""
Response.Write ".data_table{"
Response.Write "   background: #fff;"
Response.Write "    padding: 15px;"
Response.Write "    box-shadow: 1px 3px 5px #aaa;"
Response.Write "    border-radius: 5px;"
Response.Write "}"
Response.Write ""
Response.Write ".data_table .btn{"
Response.Write "    padding: 5px 10px;"
Response.Write "    margin: 10px 3px 10px 0;"
Response.Write "}"


Response.Write "    #previous, #next {"
Response.Write "        float: right;"
Response.Write "    }"

Response.Write "        #filterButton, #resetButton, #previous, #next {"
Response.Write "            border-radius: 5px; "
Response.Write "            background-color: #007bff; "
Response.Write "            color: #ffffff; "
Response.Write "            padding: 8px 12px; "
Response.Write "            margin-right: 10px; "
Response.Write "            border: none;"
Response.Write "            cursor: pointer;"
Response.Write "        }"
Response.Write "        #filterButton:hover, #resetButton:hover, #previous:hover, #next:hover {"
Response.Write "            background-color: #0056b3;"
Response.Write "        }"

Response.Write ".styled-input {"
Response.Write "    border: 1px solid #ccc;"
Response.Write "    border-radius: 8px;"
Response.Write "    padding: 8px;"
Response.Write "    margin-right: 10px;"
Response.Write "    font-size: 14px;"
Response.Write "    outline: none;"
Response.Write "}"
Response.Write ""
Response.Write ".styled-button {"
Response.Write "    background-color: #007bff;"
Response.Write "    color: white;"
Response.Write "    border: none;"
Response.Write "    border-radius: 8px;"
Response.Write "    padding: 8px 16px;"
Response.Write "    cursor: pointer;"
Response.Write "    font-size: 14px;"
Response.Write "}"
Response.Write ""
Response.Write ".styled-button:hover {"
Response.Write "    background-color: #0056b3;"
Response.Write "}"


Response.Write "</style>"


End Sub




'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
