'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

response.Clear

Dim periodStart, periodEnd, brnchID, diseaseID
Dim page_title
page_title = ""

If Len(Trim(Request.QueryString("selectedValue"))) > 1 Then
    periodStart = Trim(Request.QueryString("selectedValue"))
    periodEnd = Trim(Request.QueryString("selectedValue1"))
    brnchID = Trim(Request.QueryString("branID"))
    diseaseID = Trim(Request.QueryString("disid"))

    periodStart = FormatDate(periodStart)
    periodEnd = FormatDate(periodEnd)
    brnchID = Trim(brnchID)
Else
    periodStart = FormatDate(Now - 1)
    periodEnd = FormatDate(Now)
    brnchID = "B001"
End If

page_title = "Admission Analysis Between " & periodStart & " And " & periodEnd & " :" & GetComboName("Branch", brnchID) & " "



AddCss

response.write "<!DOCTYPE html>"
response.write "<html lang='en'>"
response.write "<head>"
response.write "<meta charset='UTF-8'>"
response.write "<meta name='viewport' content='width=device-width, initial-scale=1.0'>"
response.write "<title>Admission Time Series Analysis</title>"
response.write "<script src='https://cdn.plot.ly/plotly-latest.min.js'></script>"
'Response.Write "    <link rel=""stylesheet"" href=""https://cdn.datatables.net/2.0.8/css/dataTables.dataTables.css"">"
'Response.Write "    <script type=""text/javascript"" src=""https://cdn.datatables.net/2.0.8/js/dataTables.js""></script>"

'Response.Write "    <script src=""https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js""></script>"
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

'======================================= spinner
'response.write "<!-- Spinner HTML -->"
    response.write "<div id='loadingSpinner' style='display: flex; align-items: center; justify-content: center; height: 100vh; width: 100%; position: fixed; top: 0; left: 0; background: rgba(255, 255, 255, 0.7); z-index: 9999;'>"
    response.write "<div style='text-align: center;'>"
    response.write "<div class='spinner' style='border: 10px solid #f3f3f3; border-top: 10px solid #3498db; border-radius: 50%; width: 120px; height: 120px; animation: spin 1s linear infinite;'></div>"
'    response.write "<p style='font-size: 18px; color: #3498db; margin-top: 10px;'>Loading data...</p>"
    response.write "</div>"
    response.write "</div>"

    response.write "<style>"
    response.write "/* Spinner animation */"
    response.write "@keyframes spin {"
    response.write "    0% { transform: rotate(0deg); }"
    response.write "    100% { transform: rotate(360deg); }"
    response.write "}"
    response.write "</style>"
    
'======================================= spinner ends
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
response.write "  <div class='tab-button active' onclick='openTab(event, ""yearlyTab"")'>Yearly Admissions</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""quarterlyTab"")'>Quarterly Admissions</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""monthlyTab"")'>Monthly Admissions</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""weeklyTab"")'>Weekly Admissions</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""diseasePairsTab"")'>Multi Morbid Admissions</div>"
response.write "</div>"


filters

'yearly tab starts here
response.write "<div id='yearlyTab' class='tab-content active'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='yearlyChartDiv' class='chart'></div>"
response.write "  </div>"

  ' table
response.write "      <table style=""width:100%"" id=""yearlyAdmissionTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    response.write "      <thead class=""table-dark"">"
    response.write "              <tr>"
    response.write "                <th>Disease</th>"
    response.write "                <th>Year</th>"
'    response.write "                <th>Week</th>"
    response.write "                <th>Frequency</th>"
    response.write "                <th>Previous Frequency</th>"
    response.write "                <th>Difference</th>"
    response.write "                <th>% Change</th>"
    
    response.write "                <th>Running Total</th>"
    response.write "                <th>Rolling Average</th>"
       response.write "                <th>Annual Total</th>"
       response.write "                <th>% Of Annual Total</th>"
          response.write "                <th>Overall Total</th>"
             response.write "                <th>% Of Overall Total</th>"
                response.write "                <th>% YearOnYear Growth</th>"

    response.write "                            </tr>"
    response.write "        </thead>"
    response.write "    </table>"
    

response.write "</div>"
'yearly tab end here

'quRTERly tab starts here
response.write "<div id='quarterlyTab' class='tab-content'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='quarterlyChartDiv' class='chart'></div>"
response.write "  </div>"

  ' table
response.write "      <table style=""width:100%"" id=""quarterlyAdmissionTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    response.write "      <thead class=""table-dark"">"
    response.write "              <tr>"
    response.write "                <th>Disease</th>"
    response.write "                <th>Year</th>"
    response.write "                <th>Quarter</th>"
    response.write "                <th>Frequency</th>"
    response.write "                <th>Previous Frequency</th>"
    response.write "                <th>Difference</th>"
    response.write "                <th>% Change</th>"
    response.write "                <th>CumulativeCount</th>"
    response.write "                <th>Running Total</th>"
    response.write "                <th>Rolling Average</th>"
    response.write "                <th>Annual Total</th>"
    response.write "                <th>% Of Annual Total</th>"
    response.write "             </tr>"
    response.write "        </thead>"
    response.write "    </table>"
    

response.write "</div>"
'qurterly ends here
' monthly tab starts here
response.write "<div id='monthlyTab' class='tab-content'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='monthlyChartDiv' class='chart'></div>"
response.write "  </div>"
'monthly table
    ' table
response.write "      <table style=""width:100%"" id=""monthlyAdmissionTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    response.write "      <thead class=""table-dark"">"
    response.write "              <tr>"
    response.write "                <th>Disease</th>"
    response.write "                <th>Year</th>"
    response.write "                <th>Week</th>"
    response.write "                <th>Frequency</th>"
    response.write "                <th>Previous Frequency</th>"
    response.write "                <th>Difference</th>"
    response.write "                <th>% Change</th>"

    response.write "                            </tr>"
    response.write "        </thead>"
    response.write "    </table>"
    

response.write "</div>"
'monthly tab ends here

'weekly tab starts here
response.write "<div id='weeklyTab' class='tab-content'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='weeklyChartDiv' class='chart'></div>"

response.write "  </div>"
response.write "  </br>"
' table
response.write "      <table style=""width:100%"" id=""weeklyAdmissionTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    response.write "      <thead class=""table-dark"">"
    response.write "              <tr>"
    response.write "                <th>Year</th>"
    response.write "                <th>Week</th>"
    response.write "                <th>Frequency</th>"
    response.write "                <th>Previous Frequency</th>"
    response.write "                <th>Difference</th>"
    response.write "                <th>% Change</th>"

    response.write "                            </tr>"
    response.write "        </thead>"
    response.write "    </table>"
    
response.write "  </div>"

response.write "</div>"


response.write "</div>"
'weekly tab ends here

'multi morbid tab starts here
response.write "<div id='diseasePairsTab' class='tab-content'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='diseasePairsChartDiv' class='chart'></div>"
response.write "  </div>"

  ' table
response.write "      <table style=""width:100%"" id=""multiMorbidAdmissionTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
    response.write "      <thead class=""table-dark"">"
    response.write "              <tr>"
   response.write "                <th>S/No.</th>"
    response.write "                <th>Year</th>"
     response.write "                <th>Diseases</th>"
    
    response.write "                <th>Frequency</th>"
'    response.write "                <th>Previous Frequency</th>"
'    response.write "                <th>Difference</th>"
'    response.write "                <th>% Change</th>"
    
    response.write "                            </tr>"
    response.write "        </thead>"
    response.write "    </table>"
    

response.write "</div>"

response.write "</body>"
response.write "</html>"


'calling subroutines starts here
get_weekly_admission_trends
get_monthly_admission_trends
get_quarterly_admission_trends
get_yearly_admission_trends
get_admission_disease_pairs
LoadingSpinner

Sub get_weekly_admission_trends()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")

    sql = "With WeeklyAdmissionsCTE AS ("
    sql = sql & " SELECT DATEPART(WEEK, Admission.AdmissionDate) AS [Week],"
    sql = sql & " YEAR(Admission.AdmissionDate) AS [Year],"
    sql = sql & " Disease.DiseaseName,"
    sql = sql & " COUNT(*) AS [Count]"
    sql = sql & " FROM Admission"
    sql = sql & " JOIN Diagnosis ON Admission.VisitationID = Diagnosis.VisitationID"
    sql = sql & " JOIN Disease ON Disease.DiseaseID = Diagnosis.DiseaseID"
    sql = sql & " WHERE (CONVERT(DATE,Admission.AdmissionDate) BETWEEN ' " & periodStart & " ' AND ' " & periodEnd & " ')"
    sql = sql & " AND (CONVERT(DATE,Diagnosis.ConsultReviewDate) BETWEEN ' " & periodStart & " ' AND ' " & periodEnd & " ')"
    sql = sql & " AND Disease.DiseaseName IN (" & diseaseID & ")"
    sql = sql & " GROUP BY DATEPART(WEEK, Admission.AdmissionDate),"
    sql = sql & " YEAR(Admission.AdmissionDate),"
    sql = sql & " Disease.DiseaseName"
    sql = sql & " ), AnalysisCTE AS ("
    sql = sql & " SELECT CONCAT(DiseaseName, ' ') AS DiseaseName, [Year], dbo.format_week_num([Week]) AS [Week], [Count],"
    sql = sql & " LAG([Count]) OVER(PARTITION BY [Year], DiseaseName ORDER BY DiseaseName, [Year], [Week]) AS PrevCount,"
    sql = sql & " [Count] - LAG([Count]) OVER(PARTITION BY [Year], DiseaseName ORDER BY DiseaseName, [Year], [Week]) AS [Difference],"
    sql = sql & " ([Count] - LAG([Count]) OVER(PARTITION BY [Year], DiseaseName ORDER BY DiseaseName, [Year], [Week])) * 100.0 / [Count] AS [%Change]"
    sql = sql & " FROM WeeklyAdmissionsCTE"
    sql = sql & " )"
    sql = sql & " SELECT DiseaseName, [Year], [Week],"
    sql = sql & " FORMAT([Count], 'N0') AS [Count],"
    sql = sql & " FORMAT(PrevCount, 'N0') AS PrevCount,"
    sql = sql & " FORMAT([Difference], 'N0') AS [Difference],"
    sql = sql & " CONVERT(NUMERIC(10, 2), [%Change]) AS [%Change]"
    sql = sql & " FROM AnalysisCTE"

    rst.open sql, conn, 3, 4

    Dim jsonData
    jsonData = "{""data"":["

    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """DiseaseName"":""" & rst.Fields("DiseaseName").Value & ""","
            jsonData = jsonData & """Year"":""" & rst.Fields("Year").Value & ""","
            jsonData = jsonData & """Week"":""" & rst.Fields("Week").Value & ""","
            jsonData = jsonData & """Count"":""" & rst.Fields("Count").Value & ""","
            jsonData = jsonData & """PrevCount"":""" & rst.Fields("PrevCount").Value & ""","
            jsonData = jsonData & """Difference"":""" & rst.Fields("Difference").Value & ""","
            jsonData = jsonData & """%Change"":""" & rst.Fields("%Change").Value & """"
            jsonData = jsonData & "},"
            rst.MoveNext
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
    response.write "    var admissionPairsWeekly = dbDataWeekly.data;"

    ' Extract unique disease-year combinations and prepare data for each combination
    response.write "    var diseaseYearCombos = [...new Set(admissionPairsWeekly.map(pair => pair.DiseaseName + pair.Year))];"
    response.write "    var traces = [];"

    ' Define a color palette
    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    response.write "    var colorIndex = 0;"

    response.write "    diseaseYearCombos.forEach(function(combo) {"
    response.write "        var filteredData = admissionPairsWeekly.filter(pair => (pair.DiseaseName + pair.Year) == combo);"
    response.write "        var disease = filteredData[0].DiseaseName.split(' ')[0];"
    response.write "        var year = filteredData[0].Year;"
    response.write "        var trace = {"
    response.write "            x: filteredData.map(pair => pair.Week),"
    response.write "            y: filteredData.map(pair => parseInt(pair.Count)),"
    response.write "            mode: 'lines+markers',"
    response.write "            name: disease + ' ' + year,"
    response.write "            text: filteredData.map(pair => 'Disease: ' + pair.DiseaseName + '<br>Week: ' + pair.Week + '<br>Count: ' + pair.Count + '<br>Previous Count: ' + pair.PrevCount + '<br>Difference: ' + pair.Difference + '<br>%Change: ' + pair['%Change']),"
    response.write "            hovertemplate: '%{text}',"
    response.write "            line: {"
    response.write "                color: colors[colorIndex % colors.length]"
    response.write "            }"
    response.write "        };"
    response.write "        traces.push(trace);"
    response.write "        colorIndex++;"
    response.write "    });"

    ' Layout for line chart
    response.write "    var lineLayout = {"
    ' response.write "        title: 'Weekly Admission Trends',"
    ' response.write "        title: 'Weekly Admission Trends',"
    response.write "        title: 'Weekly Admission Trends from " & FormatDateNew(periodStart) & " to " & FormatDateNew(periodEnd) & "',"

    response.write "        xaxis: { title: 'Week' },"
    response.write "        yaxis: { title: 'Frequency' },"
    response.write "        legend: { orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    response.write "        height: 600, width: window.innerWidth * 1.0"
    response.write "    };"

    ' Plot the line chart
    response.write "    Plotly.newPlot('weeklyChartDiv', traces, lineLayout);"

    response.write "});"
    response.write "</script>"
    
    
    
    
    response.write "</script>"
    
    ' weekly data table  script starts here
     response.write "<script>"
     response.write "    new DataTable('#weeklyAdmissionTable', {"
    response.write "        data: dbDataWeekly.data,"
    response.write "        columns: ["
    response.write "            { data: 'Year' },"
    response.write "            { data: 'Week' },"
'    response.write "            { data: 'RevenueSourceAmount' },"
    response.write "            { data: 'Count' },"
    response.write "            { data: 'PrevCount' },"
    response.write "            { data: 'Difference' },"
    response.write "            { data: '%Change' }"
  

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
response.write "                    title: '" & brnchName & " Weekly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"

response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: '" & brnchName & " Weekly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"



response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: '" & brnchName & " Weekly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"


response.write "                {"
response.write "                    extend: 'print',"
response.write "                    text: 'PRINT',"
response.write "                    title: '" & brnchName & " Weekly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & " '"
response.write "                }"

response.write "            ]"
    response.write "    });"
    response.write "</script>"
    
    
End Sub

Sub get_monthly_admission_trends()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")

    sql = " With MonthlyAdmissionsCTE AS ("
    sql = sql & " SELECT DATEPART(MONTH, Admission.AdmissionDate) AS [MonthNum],"
    sql = sql & " DATENAME(MONTH, Admission.AdmissionDate) AS [Month],"
    sql = sql & " YEAR(Admission.AdmissionDate) AS [Year],"
    sql = sql & " Disease.DiseaseName,"
    sql = sql & " COUNT(*) AS [Count]"
    sql = sql & " FROM Admission"
    sql = sql & " JOIN Diagnosis ON Admission.VisitationID = Diagnosis.VisitationID"
    sql = sql & " JOIN Disease ON Disease.DiseaseID = Diagnosis.DiseaseID"
    sql = sql & " WHERE (Admission.AdmissionDate BETWEEN ' " & periodStart & " ' AND ' " & periodEnd & " ')"
    sql = sql & " AND (Diagnosis.ConsultReviewDate BETWEEN ' " & periodStart & " ' AND ' " & periodEnd & " ')"
    sql = sql & " AND Disease.DiseaseName IN ('MALARIA', 'DIABETES MELLITUS')"
    sql = sql & " GROUP BY DATEPART(MONTH, Admission.AdmissionDate),"
    sql = sql & " DATENAME(MONTH, Admission.AdmissionDate),"
    sql = sql & " YEAR(Admission.AdmissionDate),"
    sql = sql & " Disease.DiseaseName"
    sql = sql & " ),"
    sql = sql & " AnalysisCTE AS ("
    sql = sql & " SELECT CONCAT(DiseaseName, ' ') AS DiseaseName, [Year], [Month], [MonthNum], [Count],"
    sql = sql & " LAG([Count]) OVER(PARTITION BY [Year], DiseaseName ORDER BY DiseaseName, [Year], [MonthNum]) AS PrevCount,"
    sql = sql & " [Count] - LAG([Count]) OVER(PARTITION BY [Year], DiseaseName ORDER BY DiseaseName, [Year], [MonthNum]) AS [Difference],"
    sql = sql & " ([Count] - LAG([Count]) OVER(PARTITION BY [Year], DiseaseName ORDER BY DiseaseName, [Year], [MonthNum])) * 100.0 / [Count] AS [%Change]"
    sql = sql & " FROM MonthlyAdmissionsCTE"
    sql = sql & " )"
    sql = sql & " SELECT DiseaseName, [Year], [Month], [MonthNum],"
    sql = sql & " FORMAT([Count], 'N0') AS [Count],"
    sql = sql & " FORMAT(PrevCount, 'N0') AS PrevCount,"
    sql = sql & " FORMAT([Difference], 'N0') AS [Difference],"
    sql = sql & " CONVERT(NUMERIC(10, 2), [%Change]) AS [%Change]"
    sql = sql & " FROM AnalysisCTE"

    rst.open sql, conn, 3, 4

    Dim jsonData
    jsonData = "{""data"":["

    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """DiseaseName"":""" & rst.Fields("DiseaseName").Value & ""","
            jsonData = jsonData & """Year"":""" & rst.Fields("Year").Value & ""","
            jsonData = jsonData & """Month"":""" & rst.Fields("Month").Value & ""","
            jsonData = jsonData & """MonthNum"":""" & rst.Fields("MonthNum").Value & ""","
            jsonData = jsonData & """Count"":""" & rst.Fields("Count").Value & ""","
            jsonData = jsonData & """PrevCount"":""" & rst.Fields("PrevCount").Value & ""","
            jsonData = jsonData & """Difference"":""" & rst.Fields("Difference").Value & ""","
            jsonData = jsonData & """%Change"":""" & rst.Fields("%Change").Value & """"
            jsonData = jsonData & "},"
            rst.MoveNext
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
    response.write "    var admissionPairsMonthly = dbDataMonthly.data;"

    ' Extract unique disease-year combinations and prepare data for each combination
    response.write "    var diseaseYearCombos = [...new Set(admissionPairsMonthly.map(pair => pair.DiseaseName + pair.Year))];"
    response.write "    var traces = [];"

    ' Define a color palette
    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    response.write "    var colorIndex = 0;"

    response.write "    diseaseYearCombos.forEach(function(combo) {"
    response.write "        var filteredData = admissionPairsMonthly.filter(pair => (pair.DiseaseName + pair.Year) == combo);"
    response.write "        var disease = filteredData[0].DiseaseName.split(' ')[0];"
    response.write "        var year = filteredData[0].Year;"
    response.write "        var trace = {"
    response.write "            x: filteredData.map(pair => pair.Month),"
    response.write "            y: filteredData.map(pair => parseInt(pair.Count)),"
    response.write "            mode: 'lines+markers',"
    response.write "            name: disease + ' ' + year,"
    response.write "            text: filteredData.map(pair => 'Disease: ' + pair.DiseaseName + '<br>Month: ' + pair.Month + '<br>Count: ' + pair.Count + '<br>Previous Count: ' + pair.PrevCount + '<br>Difference: ' + pair.Difference + '<br>%Change: ' + pair['%Change']),"
    response.write "            hovertemplate: '%{text}',"
    response.write "            line: {"
    response.write "                color: colors[colorIndex % colors.length]"
    response.write "            }"
    response.write "        };"
    response.write "        traces.push(trace);"
    response.write "        colorIndex++;"
    response.write "    });"

    ' Layout for line chart
    response.write "    var lineLayout = {"
    ' response.write "        title: 'Monthly Admission Trends',"
    response.write "        title: 'Monthly Admission Trends from " & FormatDateNew(periodStart) & " to " & FormatDateNew(periodEnd) & "',"
        ' response.write "        title: 'Quarterly Admission Trends from " & periodStart & " to " & periodEnd & "',"

    response.write "        xaxis: { title: 'Month' },"
    response.write "        yaxis: { title: 'Count' },"
    response.write "        legend: { orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    response.write "        height: 600, width: window.innerWidth * 1.0"
    response.write "    };"

    ' Plot the line chart
    response.write "    Plotly.newPlot('monthlyChartDiv', traces, lineLayout);"

    response.write "});"
    response.write "</script>"
    
    
    ' monthly admission table starts here
    response.write "<script>"
     response.write "    new DataTable('#monthlyAdmissionTable', {"
    response.write "        data: dbDataMonthly.data,"
    response.write "        columns: ["
    response.write "            { data: 'DiseaseName' },"
    response.write "            { data: 'Year' },"
    response.write "            { data: 'Month' },"
'    response.write "            { data: 'RevenueSourceAmount' },"
    response.write "            { data: 'Count' },"
    response.write "            { data: 'PrevCount' },"
    response.write "            { data: 'Difference' },"
    response.write "            { data: '%Change' }"

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
response.write "                    title: '" & brnchName & " Monthly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"

response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: '" & brnchName & " Monthly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"



response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: '" & brnchName & " Monthly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"


response.write "                {"
response.write "                    extend: 'print',"
response.write "                    text: 'PRINT',"
response.write "                    title: '" & brnchName & " Monthly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & " '"
response.write "                }"

response.write "            ]"
    response.write "    });"
    response.write "</script>"
    
    
End Sub


Sub get_quarterly_admission_trends()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")

    sql = " WITH QuarterlyAdmissionsCTE AS ("
    sql = sql & "   SELECT"
    sql = sql & "  DATEPART(QUARTER, Admission.AdmissionDate) AS [QuarterNum],"
    sql = sql & " CONCAT(DATENAME(YEAR, Admission.AdmissionDate), 'Q', DATENAME(QUARTER, Admission.AdmissionDate)) AS [Quarter],"
    sql = sql & " YEAR(Admission.AdmissionDate) AS [Year],"
    sql = sql & " Disease.DiseaseName,"
    sql = sql & " COUNT(*) AS [Count]"
    sql = sql & " From Admission"
    sql = sql & " JOIN Diagnosis ON Admission.VisitationID = Diagnosis.VisitationID"
    sql = sql & " JOIN Disease ON Disease.DiseaseID = Diagnosis.DiseaseID"
    sql = sql & " Where"
    sql = sql & "  Admission.AdmissionDate Between ' " & periodStart & " ' AND ' " & periodEnd & " ' AND"
    sql = sql & " Diagnosis.ConsultReviewDate Between ' " & periodStart & " ' AND ' " & periodEnd & " ' AND"
    sql = sql & " Disease.DiseaseName IN ('MALARIA', 'DIABETES MELLITUS')"
    sql = sql & " Group By"
    sql = sql & " CONCAT(DATENAME(YEAR, Admission.AdmissionDate), 'Q', DATENAME(QUARTER, Admission.AdmissionDate)),"
    sql = sql & " YEAR(Admission.AdmissionDate),"
    sql = sql & " DATEPART(QUARTER, Admission.AdmissionDate),"
    sql = sql & " Disease.DiseaseName"
    sql = sql & " ),"
    sql = sql & " AnalysisCTE AS ("
    sql = sql & " SELECT"
    sql = sql & " CONCAT(DiseaseName, ' ') AS DiseaseName,"
    sql = sql & " [Year],"
    sql = sql & " [Quarter],"
    sql = sql & " [QuarterNum],"
    sql = sql & " [Count],"
    sql = sql & " LAG([Count]) OVER(PARTITION BY [Year], DiseaseName ORDER BY DiseaseName, [Year], [QuarterNum]) AS PrevCount,"
    sql = sql & " [Count] - LAG([Count]) OVER(PARTITION BY [Year], DiseaseName ORDER BY DiseaseName, [Year], [QuarterNum]) AS [Difference],"
    sql = sql & " ([Count] - LAG([Count]) OVER(PARTITION BY [Year], DiseaseName ORDER BY DiseaseName, [Year], [QuarterNum])) * 100.0 / [Count] AS [%QuarterOnChange],"
    sql = sql & " SUM([Count]) OVER(PARTITION BY DiseaseName, [Year] ORDER BY [Year], [QuarterNum]) AS CumulativeCount,"
    sql = sql & " SUM([Count]) OVER(PARTITION BY DiseaseName ORDER BY [Year], [QuarterNum]) AS RunningTotal,"
    sql = sql & " AVG([Count]) OVER(PARTITION BY DiseaseName ORDER BY [Year], [QuarterNum] ROWS BETWEEN 3 PRECEDING AND CURRENT ROW) AS RollingAverage,"
    sql = sql & " SUM([Count]) OVER(PARTITION BY [Year]) AS YearlyTotal,"
    sql = sql & " [Count] * 100.0 / SUM([Count]) OVER(PARTITION BY [Year]) AS PercentageOfYearlyTotal,"
    sql = sql & " SUM([Count]) OVER() AS OverallTotal,"
    sql = sql & " [Count] * 100.0 / SUM([Count]) OVER() AS PercentageOfOverallTotal,"
    sql = sql & " CASE"
    sql = sql & " WHEN LAG([Count]) OVER(PARTITION BY DiseaseName, [QuarterNum] ORDER BY [Year], [QuarterNum]) IS NOT NULL"
    sql = sql & " THEN ([Count] - LAG([Count]) OVER(PARTITION BY DiseaseName, [QuarterNum] ORDER BY [Year], [QuarterNum])) * 100.0 /"
    sql = sql & " LAG([Count]) OVER(PARTITION BY DiseaseName, [QuarterNum] ORDER BY [Year], [QuarterNum])"
    sql = sql & " ELSE NULL"
    sql = sql & "  END As [%YearOnYearGrowth]"
    sql = sql & " From QuarterlyAdmissionsCTE"
    sql = sql & " )"
    sql = sql & " SELECT"
    sql = sql & " DiseaseName,"
    sql = sql & " [Year],"
    sql = sql & " [Quarter],"
    sql = sql & " [QuarterNum],"
    sql = sql & " FORMAT([Count], 'N0') AS [Count],"
    sql = sql & " FORMAT(PrevCount, 'N0') AS PrevCount,"
    sql = sql & " FORMAT([Difference], 'N0') AS [Difference],"
    sql = sql & " CONVERT(NUMERIC(10, 2), [%QuarterOnChange]) AS [%QuarterOnChange],"
    sql = sql & " FORMAT(CumulativeCount, 'N0') AS CumulativeCount,"
    sql = sql & " FORMAT(RunningTotal, 'N0') AS RunningTotal,"
    sql = sql & " CONVERT(NUMERIC(10, 2), RollingAverage) AS RollingAverage,"
    sql = sql & " FORMAT(YearlyTotal, 'N0') AS YearlyTotal,"
    sql = sql & " CONVERT(NUMERIC(10, 2), PercentageOfYearlyTotal) AS PercentageOfYearlyTotal,"
    sql = sql & " FORMAT(OverallTotal, 'N0') AS OverallTotal,"
    sql = sql & " CONVERT(NUMERIC(10, 2), PercentageOfOverallTotal) AS PercentageOfOverallTotal,"
    sql = sql & " CONVERT(Numeric(10, 2), [%YearOnYearGrowth]) As [%YearOnYearGrowth]"
    sql = sql & " From AnalysisCTE ORDER BY DiseaseName, [Year], [QuarterNum];"

    rst.open sql, conn, 3, 4

    Dim jsonData
    jsonData = "{""data"":["

    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """DiseaseName"":""" & rst.Fields("DiseaseName").Value & ""","
            jsonData = jsonData & """Year"":""" & rst.Fields("Year").Value & ""","
            jsonData = jsonData & """Quarter"":""" & rst.Fields("Quarter").Value & ""","
            jsonData = jsonData
            jsonData = jsonData & """QuarterNum"":""" & rst.Fields("QuarterNum").Value & ""","
            jsonData = jsonData & """Count"":""" & rst.Fields("Count").Value & ""","
            jsonData = jsonData & """PrevCount"":""" & rst.Fields("PrevCount").Value & ""","
            jsonData = jsonData & """Difference"":""" & rst.Fields("Difference").Value & ""","
            jsonData = jsonData & """%QuarterOnChange"":""" & rst.Fields("%QuarterOnChange").Value & ""","
            jsonData = jsonData & """CumulativeCount"":""" & rst.Fields("CumulativeCount").Value & ""","
            jsonData = jsonData & """RunningTotal"":""" & rst.Fields("RunningTotal").Value & ""","
            jsonData = jsonData & """RollingAverage"":""" & rst.Fields("RollingAverage").Value & ""","
            jsonData = jsonData & """YearlyTotal"":""" & rst.Fields("YearlyTotal").Value & ""","
            jsonData = jsonData & """PercentageOfYearlyTotal"":""" & rst.Fields("PercentageOfYearlyTotal").Value & ""","
            jsonData = jsonData & """OverallTotal"":""" & rst.Fields("OverallTotal").Value & ""","
            jsonData = jsonData & """PercentageOfOverallTotal"":""" & rst.Fields("PercentageOfOverallTotal").Value & ""","
            jsonData = jsonData & """%YearOnYearGrowth"":""" & rst.Fields("%YearOnYearGrowth").Value & """"
            jsonData = jsonData & "},"
            rst.MoveNext
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
    response.write "    var admissionPairsQuarterly = dbDataQuarterly.data;"

    ' Extract unique disease-year combinations and prepare data for each combination
    response.write "    var diseaseQuarterCombos = [...new Set(admissionPairsQuarterly.map(pair => pair.DiseaseName + pair.Year))];"
    response.write "    var traces = [];"

    ' Define a color palette
    response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
    response.write "    var colorIndex = 0;"

    response.write "    diseaseQuarterCombos.forEach(function(combo) {"
    response.write "        var filteredData = admissionPairsQuarterly.filter(pair => (pair.DiseaseName + pair.Year) == combo);"
    response.write "        var disease = filteredData[0].DiseaseName.split(' ')[0];"
    response.write "        var year = filteredData[0].Year;"
    response.write "        var trace = {"
    response.write "            x: filteredData.map(pair => pair.Quarter),"
    response.write "            y: filteredData.map(pair => parseInt(pair.Count)),"
    response.write "            mode: 'lines+markers',"
    response.write "            name: disease + ' ' + year,"
    response.write "            text: filteredData.map(pair => 'Disease: ' + pair.DiseaseName + '<br>Quarter: ' + pair.Quarter + '<br>Count: ' + pair.Count + '<br>Previous Count: ' + pair.PrevCount + '<br>Difference: ' + pair.Difference + '<br>%Change: ' + pair['%QuarterOnChange'] + '<br>Cumulative Count: ' + pair.CumulativeCount + '<br>Running Total: ' + pair.RunningTotal + '<br>Rolling Average: ' + pair.RollingAverage + '<br>Yearly Total: ' + pair.YearlyTotal + '<br>% of Yearly Total: ' + pair.PercentageOfYearlyTotal + '<br>Overall Total: ' + pair.OverallTotal + '<br>% of Overall Total: ' + pair.PercentageOfOverallTotal + '<br>% Year-on-Year Growth: ' + pair['%YearOnYearGrowth']),"
    response.write "            hovertemplate: '%{text}',"
    response.write "            line: {"
    response.write "                color: colors[colorIndex % colors.length]"
    response.write "            }"
    response.write "        };"
    response.write "        traces.push(trace);"
    response.write "        colorIndex++;"
    response.write "    });"

    ' Layout for line chart
    response.write "    var lineLayout = {"
    ' response.write "        title: 'Quarterly Admission Trends',"
    response.write "        title: 'Quarterly Admission Trends from " & FormatDateNew(periodStart) & " to " & FormatDateNew(periodEnd) & "',"
    response.write "        xaxis: { title: 'Quarter' },"
    response.write "        yaxis: { title: 'Count' },"
    response.write "        legend: { orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center', yanchor: 'top' },"
    response.write "        height: 600, width: window.innerWidth * 1.0"
    response.write "    };"

    ' Plot the line chart
    response.write "    Plotly.newPlot('quarterlyChartDiv', traces, lineLayout);"

    response.write "});"
    response.write "</script>"
    
    
       ' quarterly admission table starts here
    response.write "<script>"
     response.write "    new DataTable('#quarterlyAdmissionTable', {"
    response.write "        data: dbDataQuarterly.data,"
    response.write "        columns: ["
    response.write "            { data: 'DiseaseName' },"
    response.write "            { data: 'Year' },"
    response.write "            { data: 'Quarter' },"
'    response.write "            { data: 'RevenueSourceAmount' },"
    response.write "            { data: 'Count' },"
    response.write "            { data: 'PrevCount' },"
    response.write "            { data: 'Difference' },"
    response.write "            { data: '%QuarterOnChange' },"
    
'response.write "            { data: 'CumulativeCount' },"
response.write "            { data: 'RunningTotal' },"
response.write "            { data: 'RollingAverage' },"

response.write "            { data: 'YearlyTotal' },"
response.write "            { data: 'PercentageOfYearlyTotal' }"

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
response.write "                    title: '" & brnchName & " Quarterly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"

response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: '" & brnchName & " Quarterly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"



response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: '" & brnchName & " Quarterly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"


response.write "                {"
response.write "                    extend: 'print',"
response.write "                    text: 'PRINT',"
response.write "                    title: '" & brnchName & " Quarterly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & " '"
response.write "                }"

response.write "            ]"
    response.write "    });"
    response.write "</script>"
    
    
    
End Sub


Sub get_yearly_admission_trends()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = " WITH YearlyAdmissionsCTE AS ("
    sql = sql & " SELECT"
    sql = sql & "   YEAR(Admission.AdmissionDate) AS [Year],"
    sql = sql & "    Disease.DiseaseName,"
    sql = sql & "    COUNT(*) AS [Count]"
    sql = sql & " From Admission"
    sql = sql & " JOIN Diagnosis ON Admission.VisitationID = Diagnosis.VisitationID"
    sql = sql & "  JOIN Disease ON Disease.DiseaseID = Diagnosis.DiseaseID"
    sql = sql & " Where"
    sql = sql & "     CONVERT(DATE,Admission.AdmissionDate) Between ' " & periodStart & " ' AND ' " & periodEnd & " ' AND "
    sql = sql & "     CONVERT(DATE,Diagnosis.ConsultReviewDate) Between ' " & periodStart & " ' AND ' " & periodEnd & " ' AND "
    sql = sql & "     Disease.DiseaseName IN ('MALARIA', 'DIABETES MELLITUS')"
    sql = sql & " Group By"
    sql = sql & "    YEAR(Admission.AdmissionDate),"
    sql = sql & "    Disease.DiseaseName"
    sql = sql & "),"
    sql = sql & " AnalysisCTE AS ("
    sql = sql & " SELECT"
    sql = sql & "    DiseaseName,"
    sql = sql & "   [Year],"
    sql = sql & "   [Count],"
    sql = sql & "   LAG([Count]) OVER(PARTITION BY DiseaseName ORDER BY [Year]) AS PrevCount,"
    sql = sql & "   COALESCE([Count] - LAG([Count]) OVER(PARTITION BY DiseaseName ORDER BY [Year]), 0) AS [Difference],"
    sql = sql & "   CASE"
    sql = sql & "      WHEN LAG([Count]) OVER(PARTITION BY DiseaseName ORDER BY [Year]) <> 0"
    sql = sql & "      THEN ([Count] - LAG([Count]) OVER(PARTITION BY DiseaseName ORDER BY [Year])) * 100.0 / LAG([Count]) OVER(PARTITION BY DiseaseName ORDER BY [Year])"
    sql = sql & "     ELSE NULL"
    sql = sql & "  END AS [%YearOnChange],"
    sql = sql & "  SUM([Count]) OVER(PARTITION BY DiseaseName, [Year] ORDER BY [Year]) AS CumulativeCount,"
    sql = sql & "  SUM([Count]) OVER(PARTITION BY DiseaseName ORDER BY [Year]) AS RunningTotal,"
    sql = sql & "   AVG([Count]) OVER(PARTITION BY DiseaseName ORDER BY [Year] ROWS BETWEEN 3 PRECEDING AND CURRENT ROW) AS RollingAverage,"
    sql = sql & "  SUM([Count]) OVER(PARTITION BY [Year]) AS YearlyTotal,"
    sql = sql & "   [Count] * 100.0 / SUM([Count]) OVER(PARTITION BY [Year]) AS PercentageOfYearlyTotal,"
    sql = sql & "   SUM([Count]) OVER() AS OverallTotal,"
    sql = sql & "    [Count] * 100.0 / SUM([Count]) OVER() AS PercentageOfOverallTotal,"
    sql = sql & "   CASE"
    sql = sql & "       WHEN LAG([Count]) OVER(PARTITION BY DiseaseName ORDER BY [Year]) <> 0"
    sql = sql & "       THEN ([Count] - LAG([Count]) OVER(PARTITION BY DiseaseName ORDER BY [Year])) * 100.0 / LAG([Count]) OVER(PARTITION BY DiseaseName ORDER BY [Year])"
    sql = sql & "       ELSE NULL"
    sql = sql & "   END As [%YearOnYearGrowth]"
    sql = sql & " From YearlyAdmissionsCTE"
    sql = sql & " )"
    sql = sql & " SELECT"
    sql = sql & "   DiseaseName,"
    sql = sql & "   [Year],"
    sql = sql & "   FORMAT([Count], 'N0') AS [Count],"
    sql = sql & "   FORMAT(COALESCE(PrevCount, 0), 'N0') AS PrevCount,"
    sql = sql & "   FORMAT([Difference], 'N0') AS [Difference],"
    sql = sql & "   CONVERT(NUMERIC(10, 2), COALESCE([%YearOnChange], 0)) AS [%YearOnChange],"
    sql = sql & "  FORMAT(CumulativeCount, 'N0') AS CumulativeCount,"
    sql = sql & "   FORMAT(RunningTotal, 'N0') AS RunningTotal,"
    sql = sql & "   CONVERT(NUMERIC(10, 2), RollingAverage) AS RollingAverage,"
    sql = sql & "  FORMAT(YearlyTotal, 'N0') AS YearlyTotal,"
    sql = sql & "   CONVERT(NUMERIC(10, 2), PercentageOfYearlyTotal) AS PercentageOfYearlyTotal,"
    sql = sql & "  FORMAT(OverallTotal, 'N0') AS OverallTotal,"
    sql = sql & "  CONVERT(NUMERIC(10, 2), PercentageOfOverallTotal) AS PercentageOfOverallTotal,"
    sql = sql & "  CONVERT(Numeric(10, 2), COALESCE([%YearOnYearGrowth], 0)) As [%YearOnYearGrowth]"
    sql = sql & " From AnalysisCTE"
    sql = sql & " ORDER BY DiseaseName, [Year];"


    rst.open sql, conn, 3, 4

    Dim jsonData
    jsonData = "{""data"":["

    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """DiseaseName"":""" & rst.Fields("DiseaseName").Value & ""","
            jsonData = jsonData & """Year"":""" & rst.Fields("Year").Value & ""","
            jsonData = jsonData & """Count"":""" & rst.Fields("Count").Value & ""","
            jsonData = jsonData & """PrevCount"":""" & rst.Fields("PrevCount").Value & ""","
            jsonData = jsonData & """Difference"":""" & rst.Fields("Difference").Value & ""","
            jsonData = jsonData & """%YearOnChange"":""" & rst.Fields("%YearOnChange").Value & ""","
            jsonData = jsonData & """CumulativeCount"":""" & rst.Fields("CumulativeCount").Value & ""","
            jsonData = jsonData & """RunningTotal"":""" & rst.Fields("RunningTotal").Value & ""","
            jsonData = jsonData & """RollingAverage"":""" & rst.Fields("RollingAverage").Value & ""","
            jsonData = jsonData & """YearlyTotal"":""" & rst.Fields("YearlyTotal").Value & ""","
            jsonData = jsonData & """PercentageOfYearlyTotal"":""" & rst.Fields("PercentageOfYearlyTotal").Value & ""","
            jsonData = jsonData & """OverallTotal"":""" & rst.Fields("OverallTotal").Value & ""","
            jsonData = jsonData & """PercentageOfOverallTotal"":""" & rst.Fields("PercentageOfOverallTotal").Value & ""","
            jsonData = jsonData & """%YearOnYearGrowth"":""" & rst.Fields("%YearOnYearGrowth").Value & """"
            jsonData = jsonData & "},"
            rst.MoveNext
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
    response.write "    var admissionPairsYearly = dbDataYearly.data;"

    ' Extract unique years and diseases for chart
    response.write "    var uniqueYears = [...new Set(admissionPairsYearly.map(pair => pair.Year))];"
    response.write "    var uniqueDiseases = [...new Set(admissionPairsYearly.map(pair => pair.DiseaseName))];"

    ' Prepare bar chart data
    response.write "    var traces = [];"

    response.write "    admissionPairsYearly.forEach(function(pair) {"
    response.write "        var trace = {"
    response.write "            x: [pair.Year],"
    response.write "            y: [parseInt(pair.Count)],"
    response.write "            type: 'bar',"
    response.write "            name: pair.DiseaseName + '-' + pair.Year,"
    response.write "            text: 'Disease: ' + pair.DiseaseName + '<br>Year: ' + pair.Year + '<br>Count: ' + pair.Count + '<br>Previous Count: ' + pair.PrevCount + '<br>Difference: ' + pair.Difference + '<br>% Change: ' + pair['%YearOnChange'] + '<br>Cumulative Count: ' + pair.CumulativeCount + '<br>Running Total: ' + pair.RunningTotal + '<br>Rolling Average: ' + pair.RollingAverage + '<br>Yearly Total: ' + pair.YearlyTotal + '<br>% of Yearly Total: ' + pair.PercentageOfYearlyTotal + '<br>Overall Total: ' + pair.OverallTotal + '<br>% of Overall Total: ' + pair.PercentageOfOverallTotal + '<br>% Year on Year Growth: ' + pair['%YearOnYearGrowth'],"
'    Response.Write "            hovertemplate: '%{text}',"
    response.write "            marker: { color: '#' + Math.floor(Math.random()*16777215).toString(16) }" ' Random color for each bar
    response.write "        };"
    response.write "        traces.push(trace);"
    response.write "    });"

    ' Chart layout
    response.write "    var layout = {"
    ' response.write "        title: 'Yearly Admissions by Disease from 2018 to 2020',"
    response.write "        title: 'Yearly Admissions by Disease from " & FormatDateNew(periodStart) & " to " & FormatDateNew(periodEnd) & "',"

    response.write "        xaxis: { title: 'Year', type: 'category' },"
    response.write "        yaxis: { title: 'Count' },"
    response.write "        barmode: 'group',"
    response.write "        legend: { orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center', yanchor: 'top' }"
    response.write "    };"

    ' Render the chart
    response.write "    Plotly.newPlot('yearlyChartDiv', traces, layout);"
    response.write "});"
    response.write "</script>"
    
    
    
       ' yearly admission table starts here
    response.write "<script>"
     response.write "    new DataTable('#yearlyAdmissionTable', {"
    response.write "        data: dbDataYearly.data,"
    response.write "        columns: ["
    response.write "            { data: 'DiseaseName' },"
    response.write "            { data: 'Year' },"
'    response.write "            { data: 'Quarter' },"
'    response.write "            { data: 'RevenueSourceAmount' },"
    response.write "            { data: 'Count' },"
    response.write "            { data: 'PrevCount' },"
    response.write "            { data: 'Difference' },"
    response.write "            { data: '%YearOnChange' },"
    
'response.write "            { data: 'CumulativeCount' },"
response.write "            { data: 'RunningTotal' },"
response.write "            { data: 'RollingAverage' },"

response.write "            { data: 'YearlyTotal' },"
response.write "            { data: 'PercentageOfYearlyTotal' },"
response.write "            { data: 'OverallTotal' },"
response.write "            { data: 'PercentageOfOverallTotal' },"
response.write "            { data: '%YearOnYearGrowth' }"

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
response.write "                    title: '" & brnchName & " Quarterly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"

response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: '" & brnchName & " Quarterly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"



response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: '" & brnchName & " Quarterly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"


response.write "                {"
response.write "                    extend: 'print',"
response.write "                    text: 'PRINT',"
response.write "                    title: '" & brnchName & " Quarterly Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & " '"
response.write "                }"

response.write "            ]"
    response.write "    });"
    response.write "</script>"
    
    
    
    
End Sub


Sub get_admission_disease_pairs()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT replace(DiseasePair,'|++|',' , ')DiseasePair, CONVERT(VARCHAR(20), AdmissionYear) AS [Year], PairCount AS [Count]"
    sql = sql & " FROM [dbo].[fn_get_admission_disease_pairs]('B001', ' " & periodStart & " ', ' " & periodEnd & " ')"

    rst.open sql, conn, 3, 4

    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """DiseasePair"":""" & rst.Fields("DiseasePair").Value & ""","
            jsonData = jsonData & """Year"":""" & rst.Fields("Year").Value & ""","
            jsonData = jsonData & """Count"":""" & rst.Fields("Count").Value & """"
            jsonData = jsonData & "},"
            counter = counter + 1
            rst.MoveNext
        Loop
        jsonData = Left(jsonData, Len(jsonData) - 1) ' Remove the trailing comma
    End If

    jsonData = jsonData & "]}"

    rst.Close
    Set rst = Nothing

    ' Send the data to the client-side
    response.write "<script>"
    response.write "var dbDataDiseasePairs = " & jsonData & ";"
    response.write "document.addEventListener('DOMContentLoaded', function() {"
    response.write "    var admissionDiseasePairs = dbDataDiseasePairs.data;"

    ' Extract unique years and disease pairs for chart
    response.write "    var uniqueYears = [...new Set(admissionDiseasePairs.map(pair => pair.Year))];"
    response.write "    var uniqueDiseasePairs = [...new Set(admissionDiseasePairs.map(pair => pair.DiseasePair))];"

    ' Prepare bar chart data
    response.write "    var traces = [];"

    response.write "    admissionDiseasePairs.forEach(function(pair) {"
    response.write "        var trace = {"
    response.write "            x: [pair.Year],"
    response.write "            y: [parseInt(pair.Count)],"
    response.write "            type: 'bar',"
    response.write "            name: pair.DiseasePair + '-' + pair.Year,"
'    response.write "            text: 'Disease Pair: ' + pair.DiseasePair + '<br>Year: ' + pair.Year + '<br>Count: ' + pair.Count,"
'    response.write "            hovertemplate: '%{text}',"
    response.write "            marker: { color: '#' + Math.floor(Math.random()*16777215).toString(16) }" ' Random color for each bar
    response.write "        };"
    response.write "        traces.push(trace);"
    response.write "    });"

    ' Chart layout
    response.write "    var layout = {"
    response.write "        title: 'Top 20 Annual Multimorbid Admissions from " & FormatDateNew(periodStart) & " to " & FormatDateNew(periodEnd) & "',"
    response.write "        xaxis: { title: 'Year'},"
'     response.write "        xaxis: { title: 'Year', type: 'category' },"
    response.write "        yaxis: { title: 'Count' },"
    response.write "        barmode: 'group',"
    response.write "        height: 800, width: window.innerWidth * 1.0 ,"
    response.write "        legend: { orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center', yanchor: 'top' }"
    response.write "    };"

    ' Render the chart
    response.write "    Plotly.newPlot('diseasePairsChartDiv', traces, layout);"
    response.write "});"
    response.write "</script>"
    
    
    ' multimorbid data table  script starts here
     response.write "<script>"
     response.write "    new DataTable('#multiMorbidAdmissionTable', {"
    response.write "        data: dbDataDiseasePairs.data,"
    response.write "        columns: ["
      response.write "            { data: 'counter' },"
    response.write "            { data: 'Year' },"
    response.write "            { data: 'DiseasePair' },"
'    response.write "            { data: 'DiseasePair', render: function(data, type, row) { return data.replace(/\|++\|/g, ','); } },"
    response.write "            { data: 'Count' }"

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
response.write "                    title: '" & brnchName & " Top 20 Annual Multimorbid  Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"

response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: '" & brnchName & " Top 20 Annual Multimorbid  Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"



response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: '" & brnchName & " Top 20 Annual Multimorbid  Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
response.write "                },"


response.write "                {"
response.write "                    extend: 'print',"
response.write "                    text: 'PRINT',"
response.write "                    title: '" & brnchName & " Top 20 Annual Multimorbid Admissions From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & " '"
response.write "                }"

response.write "            ]"
    response.write "    });"
   
    
End Sub

Sub LoadingSpinner()
 response.write "</script>"
    
     ' Function to hide spinner
     response.write "<script>"
    response.write "function hideSpinner() {"
    response.write "    document.getElementById('loadingSpinner').style.display = 'none';"
    response.write "}"
    response.write "hideSpinner()"
    response.write "</script>"
End Sub

Sub InitPageScript()
  Dim htStr
  'Client Script
  htStr = ""
  htStr = htStr & "<script id=""scptPrintLayoutExtraScript"" LANGUAGE=""javascript"">" & vbCrLf
  htStr = htStr & vbCrLf

  ' Function to process URL
  htStr = htStr & "function processurl(url) {" & vbCrLf
  htStr = htStr & "  // Add any URL processing logic here if needed" & vbCrLf
  htStr = htStr & "  return url;" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' RefreshPage()
  htStr = htStr & "function RefreshPage(){" & vbCrLf
  htStr = htStr & "  window.location.reload();" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Helper function to get query parameter value by name
  htStr = htStr & "function getQueryParam(param) { " & vbCrLf
  htStr = htStr & "  let urlParams = new URLSearchParams(window.location.search); " & vbCrLf
  htStr = htStr & "  return urlParams.get(param); " & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Helper function to format date
  htStr = htStr & "function formatDate(dateString) { " & vbCrLf
  htStr = htStr & "  dateString = String(dateString).trim(); " & vbCrLf
  htStr = htStr & "  var date = new Date(dateString); " & vbCrLf
  htStr = htStr & "  if (isNaN(date)) { " & vbCrLf
  htStr = htStr & "    console.error('Invalid date string'); " & vbCrLf
  htStr = htStr & "    return null; " & vbCrLf
  htStr = htStr & "  } " & vbCrLf
  htStr = htStr & "  var year = date.getFullYear(); " & vbCrLf
  htStr = htStr & "  var month = String(date.getMonth() + 1).padStart(2, '0'); " & vbCrLf
  htStr = htStr & "  var day = String(date.getDate()).padStart(2, '0'); " & vbCrLf
  htStr = htStr & "  return `${year}-${month}-${day}`; " & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Define MonthOnchange function
  htStr = htStr & "function MonthOnchange(){" & vbCrLf
  htStr = htStr & "  var mth = document.getElementById('NoOfDays').value;" & vbCrLf
  htStr = htStr & "  var dayid = 0;" & vbCrLf
  htStr = htStr & "  var ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=dispAdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "  ur = ur + '&WorkingDayID=DAY20160401&month=' + mth + '&yearid=' + dayid;" & vbCrLf
  htStr = htStr & "  window.location.href = processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Define BranchOnchange function
  htStr = htStr & "function BranchOnchange(){" & vbCrLf
  htStr = htStr & "  branchID1 = document.getElementById('Branchs').value;" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Define DiseaseOnchange function
  htStr = htStr & "function DiseaseOnchange(){" & vbCrLf
  htStr = htStr & "  var disid = document.getElementById('Disease').value;" & vbCrLf
  htStr = htStr & "  var ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=dispAdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "  ur = ur + '&WorkingDayID=DAY20160401&diseaseID=' + disid;" & vbCrLf
  htStr = htStr & "  window.location.href = processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Define PeriodOnclick function
  htStr = htStr & "function PeriodOnclick(){" & vbCrLf
  htStr = htStr & "  var branchID1 = document.getElementById('Branchs').value;" & vbCrLf
  htStr = htStr & "  var startDate1 = document.getElementById('startDate').value;" & vbCrLf
  htStr = htStr & "  var endDate1 = document.getElementById('endDate').value;" & vbCrLf
  htStr = htStr & "  startDate1 = startDate1.split('-').join('');" & vbCrLf
  htStr = htStr & "  endDate1 = endDate1.split('-').join('');" & vbCrLf
  htStr = htStr & "  var ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=dispAdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "  ur = ur + '&startDate=' + startDate1 + '&endDate=' + endDate1 + '&branID=' + branchID1;" & vbCrLf
  htStr = htStr & "  window.location.href = processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Set default values on page load
  htStr = htStr & "document.addEventListener('DOMContentLoaded', (event) => { " & vbCrLf
  htStr = htStr & "  var dropdown = document.getElementById('Branchs'); " & vbCrLf
  htStr = htStr & "  var defaultBranch = getQueryParam('branID'); " & vbCrLf
  htStr = htStr & "  for (var i = 0; i < dropdown.options.length; i++) { " & vbCrLf
  htStr = htStr & "    if (dropdown.options[i].value === defaultBranch) { " & vbCrLf
  htStr = htStr & "      dropdown.selectedIndex = i; " & vbCrLf
  htStr = htStr & "      break; " & vbCrLf
  htStr = htStr & "    } " & vbCrLf
  htStr = htStr & "  } " & vbCrLf

  ' Set start date
  htStr = htStr & "  var selectedValue = getQueryParam('selectedValue'); " & vbCrLf
  htStr = htStr & "  if (selectedValue === null) { " & vbCrLf
  htStr = htStr & "    var startDateInput = document.getElementById('startDate'); " & vbCrLf
  htStr = htStr & "    var today = new Date(); " & vbCrLf
  htStr = htStr & "    var day = ('0' + today.getDate()).slice(-2); " & vbCrLf
  htStr = htStr & "    var month = ('0' + (today.getMonth() + 1)).slice(-2); " & vbCrLf
  htStr = htStr & "    var todayString = today.getFullYear() + '-' + month + '-' + day; " & vbCrLf
  htStr = htStr & "    startDateInput.value = todayString; " & vbCrLf
  htStr = htStr & "  } else { " & vbCrLf
  htStr = htStr & "    var formattedDate = formatDate(selectedValue); " & vbCrLf
  htStr = htStr & "    if (formattedDate) { " & vbCrLf
  htStr = htStr & "      console.log('Formatted Date: ' + formattedDate); " & vbCrLf
  htStr = htStr & "    } else { " & vbCrLf
  htStr = htStr & "      console.error('Invalid date string'); " & vbCrLf
  htStr = htStr & "    } " & vbCrLf
  htStr = htStr & "    var startDateInput = document.getElementById('startDate'); " & vbCrLf
  htStr = htStr & "    startDateInput.value = formattedDate; " & vbCrLf
  htStr = htStr & "  } " & vbCrLf

  htStr = htStr & "});" & vbCrLf

  ' Closing script tag
  htStr = htStr & "</script>" & vbCrLf

  ' Add script to the page
  response.write htStr
End Sub


Sub InitPageScript03()
  Dim htStr
  'Client Script
  htStr = ""
  htStr = htStr & "<script id=""scptPrintLayoutExtraScript"" LANGUAGE=""javascript"">" & vbCrLf
  htStr = htStr & vbCrLf

  'RefreshPage()
  htStr = htStr & "function RefreshPage(){" & vbCrLf
  htStr = htStr & "window.location.reload();" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Helper function to get query parameter value by name
  htStr = htStr & "function getQueryParam(param) { " & vbCrLf
  htStr = htStr & "let urlParams = new URLSearchParams(window.location.search); " & vbCrLf
  htStr = htStr & "return urlParams.get(param); " & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Helper function to format date
  htStr = htStr & "function formatDate(dateString) { " & vbCrLf
  htStr = htStr & "dateString = String(dateString).trim(); " & vbCrLf
  htStr = htStr & "var date = new Date(dateString); " & vbCrLf
  htStr = htStr & "if (isNaN(date)) { " & vbCrLf
  htStr = htStr & "console.error('Invalid date string'); " & vbCrLf
  htStr = htStr & "return null; " & vbCrLf
  htStr = htStr & "} " & vbCrLf
  htStr = htStr & "var year = date.getFullYear(); " & vbCrLf
  htStr = htStr & "var month = String(date.getMonth() + 1).padStart(2, '0'); " & vbCrLf
  htStr = htStr & "var day = String(date.getDate()).padStart(2, '0'); " & vbCrLf
  htStr = htStr & "return `${year}-${month}-${day}`; " & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Set default values on page load
  htStr = htStr & "document.addEventListener('DOMContentLoaded', (event) => { " & vbCrLf
  htStr = htStr & "var dropdown = document.getElementById('Branchs'); " & vbCrLf
  htStr = htStr & "var defaultBranch = getQueryParam('branID'); " & vbCrLf
  htStr = htStr & "for (var i = 0; i < dropdown.options.length; i++) { " & vbCrLf
  htStr = htStr & "if (dropdown.options[i].value === defaultBranch) { " & vbCrLf
  htStr = htStr & "dropdown.selectedIndex = i; " & vbCrLf
  htStr = htStr & "break; " & vbCrLf
  htStr = htStr & "} " & vbCrLf
  htStr = htStr & "} " & vbCrLf

  ' Set start date
  htStr = htStr & "var selectedValue = getQueryParam('selectedValue'); " & vbCrLf
  htStr = htStr & "if (selectedValue === null) { " & vbCrLf
  htStr = htStr & "var startDateInput = document.getElementById('startDate'); " & vbCrLf
  htStr = htStr & "var today = new Date(); " & vbCrLf
  htStr = htStr & "var day = ('0' + today.getDate()).slice(-2); " & vbCrLf
  htStr = htStr & "var month = ('0' + (today.getMonth() + 1)).slice(-2); " & vbCrLf
  htStr = htStr & "var todayString = today.getFullYear() + '-' + month + '-' + day; " & vbCrLf
  htStr = htStr & "startDateInput.value = todayString; " & vbCrLf
  htStr = htStr & "} else { " & vbCrLf
  htStr = htStr & "var formattedDate = formatDate(selectedValue); " & vbCrLf
  htStr = htStr & "if (formattedDate) { " & vbCrLf
  htStr = htStr & "console.log('Formatted Date: ' + formattedDate); " & vbCrLf
  htStr = htStr & "} else { " & vbCrLf
  htStr = htStr & "console.error('Invalid date string'); " & vbCrLf
  htStr = htStr & "} " & vbCrLf
  htStr = htStr & "var startDateInput = document.getElementById('startDate'); " & vbCrLf
  htStr = htStr & "startDateInput.value = formattedDate; " & vbCrLf
  htStr = htStr & "} " & vbCrLf

  ' Set end date
  htStr = htStr & "var selectedValue1 = getQueryParam('selectedValue1'); " & vbCrLf
  htStr = htStr & "if (selectedValue1 === null) { " & vbCrLf
  htStr = htStr & "var endDateInput = document.getElementById('endDate'); " & vbCrLf
  htStr = htStr & "var today = new Date(); " & vbCrLf
  htStr = htStr & "var day = ('0' + today.getDate()).slice(-2); " & vbCrLf
  htStr = htStr & "var month = ('0' + (today.getMonth() + 1)).slice(-2); " & vbCrLf
  htStr = htStr & "var todayString = today.getFullYear() + '-' + month + '-' + day; " & vbCrLf
  htStr = htStr & "endDateInput.value = todayString; " & vbCrLf
  htStr = htStr & "} else { " & vbCrLf
  htStr = htStr & "var formattedDate = formatDate(selectedValue1); " & vbCrLf
  htStr = htStr & "if (formattedDate) { " & vbCrLf
  htStr = htStr & "console.log('Formatted Date: ' + formattedDate); " & vbCrLf
  htStr = htStr & "} else { " & vbCrLf
  htStr = htStr & "console.error('Invalid date string'); " & vbCrLf
  htStr = htStr & "} " & vbCrLf
  htStr = htStr & "var endDateInput = document.getElementById('endDate'); " & vbCrLf
  htStr = htStr & "endDateInput.value = formattedDate; " & vbCrLf
  htStr = htStr & "} " & vbCrLf
  htStr = htStr & "}); " & vbCrLf

  ' Define MonthOnchange function
  htStr = htStr & "function MonthOnchange(){" & vbCrLf
  htStr = htStr & "var mth = document.getElementById('NoOfDays').value;" & vbCrLf
  htStr = htStr & "var dayid = 0;" & vbCrLf
  htStr = htStr & "var ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=dispAdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&month=' + mth + '&yearid=' + dayid;" & vbCrLf
  htStr = htStr & "window.location.href = processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Define BranchOnchange function
  htStr = htStr & "function BranchOnchange(){" & vbCrLf
  htStr = htStr & "branchID1 = document.getElementById('Branchs').value;" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Define DiseaseOnchange function
  htStr = htStr & "function DiseaseOnchange(){" & vbCrLf
  htStr = htStr & "var diseaseID1 = document.getElementById('Diseases').value;" & vbCrLf
  htStr = htStr & "console.log('Selected Disease ID: ' + diseaseID1);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Define YearOnchange function
  htStr = htStr & "function YearOnchange(){" & vbCrLf
  htStr = htStr & "var dayid = GetEleVal('NoOfDay');" & vbCrLf
  htStr = htStr & "var mth = 0;" & vbCrLf
  htStr = htStr & "var emr = GetEleVal('emrdata');" & vbCrLf
  htStr = htStr & "var ispr = 0;" & vbCrLf
  htStr = htStr & "var ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=dispAdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&month=' + mth + '&yearid=' + dayid;" & vbCrLf
  htStr = htStr & "window.location.href = processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Define PeriodOnclick function
  htStr = htStr & "function PeriodOnclick(){" & vbCrLf
  htStr = htStr & "var branchID1 = document.getElementById('Branchs').value;" & vbCrLf
  htStr = htStr & "var startDate1 = document.getElementById('startDate').value;" & vbCrLf
  htStr = htStr & "var endDate1 = document.getElementById('endDate').value;" & vbCrLf
  htStr = htStr & "startDate1 = startDate1.split('-').join('');" & vbCrLf
  htStr = htStr & "endDate1 = endDate1.split('-').join('');" & vbCrLf
  htStr = htStr & "var ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=dispAdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "ur = ur + '&startDate=' + startDate1 + '&endDate=' + endDate1 + '&branID=' + branchID1;" & vbCrLf
  htStr = htStr & "window.location.href = processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Closing script tag
  htStr = htStr & "</script>" & vbCrLf

  ' Add script to the page
  response.write htStr
End Sub

Sub InitPageScript04()
  Dim htStr
  'Client Script
  htStr = ""
  htStr = htStr & "<script id=""scptPrintLayoutExtraScript"" LANGUAGE=""javascript"">" & vbCrLf
  htStr = htStr & vbCrLf

  ' Function to process URL
  htStr = htStr & "function processurl(url) {" & vbCrLf
  htStr = htStr & "  // Add any URL processing logic here if needed" & vbCrLf
  htStr = htStr & "  return url;" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' RefreshPage()
  htStr = htStr & "function RefreshPage(){" & vbCrLf
  htStr = htStr & "  window.location.reload();" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Helper function to get query parameter value by name
  htStr = htStr & "function getQueryParam(param) { " & vbCrLf
  htStr = htStr & "  let urlParams = new URLSearchParams(window.location.search); " & vbCrLf
  htStr = htStr & "  return urlParams.get(param); " & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Helper function to format date
  htStr = htStr & "function formatDate(dateString) { " & vbCrLf
  htStr = htStr & "  dateString = String(dateString).trim(); " & vbCrLf
  htStr = htStr & "  var date = new Date(dateString); " & vbCrLf
  htStr = htStr & "  if (isNaN(date)) { " & vbCrLf
  htStr = htStr & "    console.error('Invalid date string'); " & vbCrLf
  htStr = htStr & "    return null; " & vbCrLf
  htStr = htStr & "  } " & vbCrLf
  htStr = htStr & "  var year = date.getFullYear(); " & vbCrLf
  htStr = htStr & "  var month = String(date.getMonth() + 1).padStart(2, '0'); " & vbCrLf
  htStr = htStr & "  var day = String(date.getDate()).padStart(2, '0'); " & vbCrLf
  htStr = htStr & "  return `${year}-${month}-${day}`; " & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Set default values on page load
  htStr = htStr & "document.addEventListener('DOMContentLoaded', (event) => { " & vbCrLf
  htStr = htStr & "  var dropdown = document.getElementById('Branchs'); " & vbCrLf
  htStr = htStr & "  var defaultBranch = getQueryParam('branID'); " & vbCrLf
  htStr = htStr & "  for (var i = 0; i < dropdown.options.length; i++) { " & vbCrLf
  htStr = htStr & "    if (dropdown.options[i].value === defaultBranch) { " & vbCrLf
  htStr = htStr & "      dropdown.selectedIndex = i; " & vbCrLf
  htStr = htStr & "      break; " & vbCrLf
  htStr = htStr & "    } " & vbCrLf
  htStr = htStr & "  } " & vbCrLf

  ' Set start date
  htStr = htStr & "  var selectedValue = getQueryParam('selectedValue'); " & vbCrLf
  htStr = htStr & "  if (selectedValue === null) { " & vbCrLf
  htStr = htStr & "    var startDateInput = document.getElementById('startDate'); " & vbCrLf
  htStr = htStr & "    var today = new Date(); " & vbCrLf
  htStr = htStr & "    var day = ('0' + today.getDate()).slice(-2); " & vbCrLf
  htStr = htStr & "    var month = ('0' + (today.getMonth() + 1)).slice(-2); " & vbCrLf
  htStr = htStr & "    var todayString = today.getFullYear() + '-' + month + '-' + day; " & vbCrLf
  htStr = htStr & "    startDateInput.value = todayString; " & vbCrLf
  htStr = htStr & "  } else { " & vbCrLf
  htStr = htStr & "    var formattedDate = formatDate(selectedValue); " & vbCrLf
  htStr = htStr & "    if (formattedDate) { " & vbCrLf
  htStr = htStr & "      console.log('Formatted Date: ' + formattedDate); " & vbCrLf
  htStr = htStr & "    } else { " & vbCrLf
  htStr = htStr & "      console.error('Invalid date string'); " & vbCrLf
  htStr = htStr & "    } " & vbCrLf
  htStr = htStr & "    var startDateInput = document.getElementById('startDate'); " & vbCrLf
  htStr = htStr & "    startDateInput.value = formattedDate; " & vbCrLf
  htStr = htStr & "  } " & vbCrLf

  ' Set end date
  htStr = htStr & "  var selectedValue1 = getQueryParam('selectedValue1'); " & vbCrLf
  htStr = htStr & "  if (selectedValue1 === null) { " & vbCrLf
  htStr = htStr & "    var endDateInput = document.getElementById('endDate'); " & vbCrLf
  htStr = htStr & "    var today = new Date(); " & vbCrLf
  htStr = htStr & "    var day = ('0' + today.getDate()).slice(-2); " & vbCrLf
  htStr = htStr & "    var month = ('0' + (today.getMonth() + 1)).slice(-2); " & vbCrLf
  htStr = htStr & "    var todayString = today.getFullYear() + '-' + month + '-' + day; " & vbCrLf
  htStr = htStr & "    endDateInput.value = todayString; " & vbCrLf
  htStr = htStr & "  } else { " & vbCrLf
  htStr = htStr & "    var formattedDate = formatDate(selectedValue1); " & vbCrLf
  htStr = htStr & "    if (formattedDate) { " & vbCrLf
  htStr = htStr & "      console.log('Formatted Date: ' + formattedDate); " & vbCrLf
  htStr = htStr & "    } else { " & vbCrLf
  htStr = htStr & "      console.error('Invalid date string'); " & vbCrLf
  htStr = htStr & "    } " & vbCrLf
  htStr = htStr & "    var endDateInput = document.getElementById('endDate'); " & vbCrLf
  htStr = htStr & "    endDateInput.value = formattedDate; " & vbCrLf
  htStr = htStr & "  } " & vbCrLf
  htStr = htStr & "}); " & vbCrLf

  ' Define MonthOnchange function
  htStr = htStr & "function MonthOnchange(){" & vbCrLf
  htStr = htStr & "  var mth = document.getElementById('NoOfDays').value;" & vbCrLf
  htStr = htStr & "  var dayid = 0;" & vbCrLf
  htStr = htStr & "  var ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=dispAdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "  ur = ur + '&WorkingDayID=DAY20160401&month=' + mth + '&yearid=' + dayid;" & vbCrLf
  htStr = htStr & "  window.location.href = processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Define BranchOnchange function
  htStr = htStr & "function BranchOnchange(){" & vbCrLf
  htStr = htStr & "  branchID1 = document.getElementById('Branchs').value;" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Define DiseaseOnchange function
  htStr = htStr & "function DiseaseOnchange(){" & vbCrLf
  htStr = htStr & "  var disid = document.getElementById('Disease').value;" & vbCrLf
  htStr = htStr & "  var ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=dispAdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "  ur = ur + '&WorkingDayID=DAY20160401&diseaseID=' + disid;" & vbCrLf
  htStr = htStr & "  window.location.href = processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Define PeriodOnclick function
  htStr = htStr & "function PeriodOnclick(){" & vbCrLf
  htStr = htStr & "  var branchID1 = document.getElementById('Branchs').value;" & vbCrLf
  htStr = htStr & "  var startDate1 = document.getElementById('startDate').value;" & vbCrLf
  htStr = htStr & "  var endDate1 = document.getElementById('endDate').value;" & vbCrLf
  htStr = htStr & "  startDate1 = startDate1.split('-').join('');" & vbCrLf
  htStr = htStr & "  endDate1 = endDate1.split('-').join('');" & vbCrLf
  htStr = htStr & "  var ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=dispAdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "  ur = ur + '&startDate=' + startDate1 + '&endDate=' + endDate1 + '&branID=' + branchID1;" & vbCrLf
  htStr = htStr & "  window.location.href = processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' Closing script tag
  htStr = htStr & "</script>" & vbCrLf

  ' Add script to the page
  response.write htStr
End Sub


Sub InitPageScript02()
  Dim htStr
  'Client Script
  htStr = ""
  htStr = htStr & "<script id=""scptPrintLayoutExtraScript"" LANGUAGE=""javascript"">" & vbCrLf
  htStr = htStr & vbCrLf
  'RefreshPage()
  htStr = htStr & "function RefreshPage(){" & vbCrLf
  htStr = htStr & "window.location.reload();" & vbCrLf
  htStr = htStr & "}" & vbCrLf

'   htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&selectedValue=' + startDate1  + ' &selectedValue1=' + endDate1 + ' &branID=' + branchID1 ;" & vbCrLf


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
htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=dispAdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
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
htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=AdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&month=' + mth  + ' &yearid=' + dayid ;" & vbCrLf
htStr = htStr & "window.location.href = processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf

  'emrdata
  htStr = htStr & "function PeriodOnclick(){ " & vbCrLf
  htStr = htStr & "var branchID1 =  document.getElementById('Branchs').value;" & vbCrLf
  htStr = htStr & "var startDate1 =  document.getElementById('startDate').value;" & vbCrLf
  htStr = htStr & "var endDate1 =  document.getElementById('endDate').value;" & vbCrLf
  htStr = htStr & "startDate1 = startDate1.trimEnd();" & vbCrLf
  htStr = htStr & "endDate1 = endDate1.trimEnd();" & vbCrLf
'   htStr = htStr & "var branchID =  'B001';" & vbCrLf
  htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=AdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&selectedValue=' + startDate1  + ' &selectedValue1=' + endDate1 + ' &branID=' + branchID1 ;" & vbCrLf


' htStr = htStr & "  var link = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=GetPharmacyDispenseChart&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=' "& fromDate &" '&selectedValue1=' "& fromDate &" '&branID=' "& fromDate &";" & vbCrLf
    htStr = htStr & "    window.location = ur;" & vbCrLf
'   htStr = htStr & "alert('This Functionality is under Maintenance');" & vbCrf
  htStr = htStr & "}" & vbCrLf

  'emrdata
  htStr = htStr & "function OpenDiseaseFilterOnclick(){ " & vbCrLf
  htStr = htStr & "var branchID1 =  document.getElementById('Branchs').value;" & vbCrLf
  htStr = htStr & "var startDate1 =  document.getElementById('startDate').value;" & vbCrLf
  htStr = htStr & "var endDate1 =  document.getElementById('endDate').value;" & vbCrLf
  htStr = htStr & "startDate1 = startDate1.trimEnd();" & vbCrLf
  htStr = htStr & "endDate1 = endDate1.trimEnd();" & vbCrLf
'   htStr = htStr & "var branchID =  'B001';" & vbCrLf
  htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=AdmissionTimeSeries&PositionForTableName=WorkingDay';" & vbCrLf
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
    response.write "<table class = 'table table-bordered'>"
    response.write "        <tr >"
     response.write "           <td> Diseases: </td>   "
    response.write "            <td style=""width: 100%;""> "
    SetDiseases
    response.write " </td>"
    response.write "        </tr>"
    
    
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
            .movefirst
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

Sub SetDiseases()
    Set rst = CreateObject("ADODB.Recordset")
    dyHt = "<select class='form-select' size=""1"" name=""Diseases"" id=""Diseases"" onchange=""DiseaseOnchange()"">"
    dyHt = dyHt & "<option value=""""></option>"
    ' dyHt = dyHt & "<option value='D001'></option>"

    sql1 = "SELECT Distinct Diagnosis.DiseaseID, DiseaseName AS DiseaseCategoryName FROM Diagnosis JOIN Disease ON Disease.DiseaseID = Diagnosis.DiseaseID"
    With rst
        .open qryPro.FltQry(sql1), conn, 3, 4
        If .RecordCount > 0 Then
            .movefirst
            Do While Not .EOF
                diseaseID = Trim(.Fields("DiseaseID"))
                diseaseName = Trim(.Fields("DiseaseCategoryName"))

                If UCase(CStr(yearId)) = UCase(diseaseID) Then
                    dyHt = dyHt & "<option value=""" & CStr(diseaseID) & """ selected>" & diseaseName & "</option>"
                Else
                    dyHt = dyHt & "<option value=""" & CStr(diseaseID) & """>" & diseaseName & "</option>"
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
    FormatDate = Year(dateValue) & "-" & Right("0" & Month(dateValue), 2) & "-" & Right("0" & day(dateValue), 2)
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
    ' Dummy function for branch name. Replace with actual implementation.
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
