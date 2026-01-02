'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
response.write "Hello Joe"

Dim periodStart, periodEnd, dateArr, datePeriod

'Retrieve query parameters
'datePeriod = Trim(Request.QueryString("Dateperiod"))
'selectedVisitStatusID = Trim(Request.QueryString("visitStatusID"))

'Parse date period
'If datePeriod <> "" Then
'    dateArr = Split(datePeriod, "||")
'    periodStart = dateArr(0)
'    periodEnd = dateArr(1)
'End If

Styling
'MultiSelectStyles

response.write "<!DOCTYPE html>"
response.write "<html lang='en'>"
response.write "<head>"
response.write "<meta charset='UTF-8'>"
response.write "<meta name='viewport' content='width=device-width, initial-scale=1.0'>"
response.write "<title>Visitation Analysis</title>"

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

' Output dropdown
response.write "<div class='header'>"
  
    'Output HTML Form for date selection
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
    response.write "        </div>"
    response.write "    </div> "
    response.write "   </form>"
response.write "</div>"

If (periodStart <> "" And periodEnd <> "") Then
    response.write "<h2 class='font-style'>SHOWING DATA FROM: " & periodStart & " TO: " & periodEnd & "</h2>"
Else
    response.write "<h2 class='font-style'>SHOWING DATA FROM: 2018-01-01 TO: 2018-01-31</h2>"
End If

response.write "<div id='yearlyTab' class='tab-content active'>"
response.write "  <div class='chart-container'>"
response.write "    <div id='yearlyChartDiv' class='chart'></div>"
response.write "  </div>"

' table

  response.write "      <table style=""width:100%"" id=""yearlyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
  response.write "      <thead class=""table-dark"">"
  response.write "        <tr>"
  response.write "             <th>No.</th>"
  response.write "             <th>Visit Year</th>"
  response.write "             <th>Visitation CostF</th>"
  response.write "             <th>Visitation Cost</th>"
  response.write "             <th>Previous Cost</th>"
  response.write "             <th>Difference</th>"
  response.write "             <th>Percent Change</th>"
  response.write "             <th>Yearly Contribution (%)</th>"
  response.write "        </tr>"
  response.write "       </thead>"
  response.write "    </table>"
response.write "</div>"

'Response.Write " it works"
response.write "</body>"
response.write "</html>"

dispVisitCostAnalysis


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
    sql = sql & "WHERE Visitation.VisitDate BETWEEN '2018-01-01' AND '2026-12-31' "
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
    
    response.write sql
    
    'Initialize and open database connection for main data
    Set rstMain = CreateObject("ADODB.Recordset")
    rstMain.open sql, conn, 3, 4
    
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
    
    ' DataTable Initialization
    response.write "<script>"
    response.write "var dbDataYearly = " & jsonData & ";"
    response.write "    new DataTable('#yearlyTable', {"
    response.write "        data: dbDataYearly.data,"
    response.write "        columns: ["
    response.write "            { data: 'counter' },"
    response.write "            { data: 'VisitYear' },"
    response.write "            { data: 'VisitationCostF' },"
    response.write "            { data: 'VisitationCost' },"
    response.write "            { data: 'PrevCost' },"
    response.write "            { data: 'Diff' },"
    response.write "            { data: 'PercentageChange' },"
    response.write "            { data: 'ContPercentage' },"
    response.write "        ],"
        
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
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
        
        response.write " .pagination {"
        response.write "    text-align: center;"
        response.write "    margin: 20px 0;"
        response.write " }"
        
        response.write " .pagination a {"
        response.write "    margin: 0 5px;"
        response.write "    padding: 10px 15px;"
        response.write "    background-color: #f1f1f1;"
        response.write "    border: 1px solid #ccc;"
        response.write "    text-decoration: none;"
        response.write "    color: #333;"
        response.write " }"
        
        response.write " .pagination a:hover {"
        response.write "    background-color: #ddd;"
        response.write " }"
        
        response.write " .font-style {"
        response.write "    font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif;"
        response.write " }"
        
        response.write " #pharmacy {"
        response.write "    padding-bottom: 10px;"
        response.write " }"
        response.write " </style>"
        
End Sub


'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
