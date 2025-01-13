'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

response.Clear
conn.commandTimeOut = 7200
Dim periodStart, periodEnd, datePeriod, dateArr, sltInsuranceTypIDs, idsArr
Dim id, formattedIDs

Styling
MultiSelectStyles

datePeriod = Trim(request.querystring("Dateperiod"))
sltInsuranceTypIDs = Trim(request.querystring("InsuranceTypeID"))

 'Parse date period
If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    periodStart = dateArr(0)
    periodEnd = dateArr(1)
End If

' Format selected drug store IDs
If sltInsuranceTypIDs <> "" Then
    idsArr = Split(sltInsuranceTypIDs, ",")
    For Each id In idsArr
        formattedIDs = formattedIDs & "'" & Trim(id) & "',"
    Next
    ' Remove the trailing comma
    formattedIDs = Left(formattedIDs, Len(formattedIDs) - 1)
End If
    
response.write "<!DOCTYPE html>"
response.write "<html lang='en'>"
response.write "<head>"
response.write "<meta charset='UTF-8'>"
response.write "<meta name='viewport' content='width=device-width, initial-scale=1.0'>"
response.write "<title>Insurance Type Analysis</title>"

response.write "<script src='https://cdn.plot.ly/plotly-latest.min.js'></script>"
response.write "<script src='https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js'></script>"

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

' Construct SQL query for dropdown options (all insurance types)
    sql = "select InsuranceTypeID, InsuranceTypeName from InsuranceType"
    ' Initialize and open database connection for dropdown options
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.open sql, conn, 3, 4

    ' Populate dropdown options
    dropdownOptions = ""

    With rstDropdown
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                optionHTML = "<option value='" & .Fields("InsuranceTypeID") & "'>" & .Fields("InsuranceTypeName") & "</option>"
                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With

    ' Close dropdown recordset
    rstDropdown.Close
    Set rstDropdown = Nothing
    
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
response.write "  <div class='tab-button active' onclick='openTab(event, ""yearlyTab"")'>Yearly Insurance Type</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""quarterlyTab"")'>Quarterly Insurance Type</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""monthlyTab"")'>Monthly Insurance Type</div>"
response.write "  <div class='tab-button' onclick='openTab(event, ""weeklyTab"")'>Weekly Insurance Type</div>"

response.write "</div>"
    
   
    response.write "<div>"
        response.write "        <label for='insuranceType' class='font-style'>Select Insurance Type:</label><br>"
        response.write "        <select id='insuranceType' name='insuranceType' multiple class='mult-select-tag'>"
        response.write dropdownOptions
        response.write "        </select>"
    response.write "</div>"
    
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
        response.write "<h2 class='font-style'>SHOWING DATA FROM: 2018-01-01 TO: 2022-01-31</h2>"
    End If


    'Yearly tab starts here

    response.write "<div id='yearlyTab' class='tab-content active'>"
    response.write "  <div class='chart-container'>"
    response.write "    <div id='yearlyChartDiv' class='chart'></div>"
    
    response.write "  </div>"
    
    
    response.write "  <div class='chart-container'>"
    
    response.write "    <div id='yearlyChartDivGender' class='chart'></div>"
    response.write "  </div>"
    
        response.write "      <table style=""width:100%"" id=""yearlyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
        response.write "      <thead class=""table-dark"">"
        response.write "            <tr>"
        response.write "                <th>No.</th>"
        response.write "                <th>Insurance Type Name</th>"
        response.write "                <th>Year</th>"
        response.write "                <th>Year Total</th>"
        response.write "                <th>Previous Year Total</th>"
        response.write "                <th>Difference</th>"
        response.write "                <th>Percentage Change</th>"
        response.write "                <th>Yearly Contribution (%)</th>"
        response.write "            </tr>"
        response.write "        </thead>"
        response.write "    </table>"
        
    response.write "</div>"
    
    response.write "<script>"
        
        ' Output scripts
        response.write "    new MultiSelectTag('insuranceType', {"
        response.write "        rounded: true,"
        response.write "        shadow: true,"
        response.write "        placeholder: 'Search',"
        response.write "        tagColor: {"
        response.write "            textColor: '#327b2c',"
        response.write "            borderColor: '#92e681',"
        response.write "            bgColor: '#eaffe6',"
        response.write "        },"
        response.write "        onChange: function (values) {"
        response.write "            console.log(values);"
        response.write "        },"
        response.write "    });"
        
        response.write "    function updateUrl() {"
        response.write "        const fromDate = document.getElementById('from').value;"
        response.write "        const toDate = document.getElementById('to').value;"
        response.write "        const insuranceTypes = Array.from(document.getElementById('insuranceType').selectedOptions).map(option => option.value).join(',');"
        response.write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp';"
        response.write "        const params = new URLSearchParams({"
        response.write "            PrintLayoutName: 'dispInsuranceTypeAnalysis',"
        response.write "            PositionForTableName: 'WorkingDay',"
        response.write "            WorkingDayID: '',"
        response.write "            Dateperiod: `${fromDate}||${toDate}`,"
        response.write "            InsuranceTypeID: insuranceTypes"
        response.write "        });"
        response.write "        const newUrl = baseUrl + '?' + params.toString();"
        response.write "        window.location.href = newUrl;"
        response.write "        console.log(newUrl);"
        response.write "    }"
    response.write "</script>"
    
    yearlyInsuranceTypeAnalysis
    
    response.write "</body>"
    response.write "</html>"
    
    Sub yearlyInsuranceTypeAnalysis()
    Dim sql, rst
    Set rst = CreateObject("ADODB.Recordset")
    
    ' Construct SQL query for main data
    sql = "WITH visitationCTE AS (" & vbCrLf
    sql = sql & "    SELECT SUM(visitcost) final_amount," & vbCrLf
    sql = sql & "           VisitationID, InsuranceTypeName," & vbCrLf
    sql = sql & "           YEAR(visitdate) [Year]" & vbCrLf
    sql = sql & "    FROM Visitation" & vbCrLf
    sql = sql & "    JOIN InsuranceType ON Visitation.InsuranceTypeID = InsuranceType.InsuranceTypeID" & vbCrLf
    
     If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "    WHERE visitdate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'" & vbCrLf
    Else
        sql = sql & "    WHERE visitdate BETWEEN '2018-01-01' AND '2022-12-31'" & vbCrLf
    End If
    
    If sltInsuranceTypIDs <> "" Then
        sql = sql & "    AND InsuranceType.InsuranceTypeID IN(" & formattedIDs & ") " & vbCrLf
    Else
        sql = sql & "    AND InsuranceType.InsuranceTypeID IS NOT NULL " & vbCrLf
    End If
    'sql = sql & "    AND InsuranceTypeName like '%112%'" & vbCrLf
    sql = sql & "    GROUP BY" & vbCrLf
    sql = sql & "    YEAR(visitdate),VisitationID,InsuranceTypeName" & vbCrLf
    sql = sql & ")," & vbCrLf
    
    sql = sql & "treatchargesCTE AS (" & vbCrLf
    sql = sql & "    SELECT SUM(finalamt) final_amount,VisitationID," & vbCrLf
    sql = sql & "           YEAR(ConsultReviewDate) [Year],InsuranceTypeID" & vbCrLf
    sql = sql & "    FROM treatcharges" & vbCrLf
    sql = sql & "    GROUP BY " & vbCrLf
    sql = sql & "    YEAR(ConsultReviewDate),VisitationID, InsuranceTypeID" & vbCrLf
    sql = sql & ")," & vbCrLf
    
    sql = sql & "investigationCTE AS (" & vbCrLf
    sql = sql & "    SELECT SUM(finalamt) final_amount,VisitationID, InsuranceTypeID," & vbCrLf
    sql = sql & "           YEAR(requestdate) [Year]" & vbCrLf
    sql = sql & "    FROM Investigation" & vbCrLf
    sql = sql & "    GROUP BY" & vbCrLf
    sql = sql & "    YEAR(requestdate),VisitationID, InsuranceTypeID" & vbCrLf
    sql = sql & ")," & vbCrLf
    
    sql = sql & "investigation2CTE AS (" & vbCrLf
    sql = sql & "    SELECT SUM(finalamt) final_amount, VisitationID, InsuranceTypeID," & vbCrLf
    sql = sql & "           YEAR(requestdate) [inv2RequestYear]" & vbCrLf
    sql = sql & "    FROM Investigation2" & vbCrLf
    sql = sql & "    GROUP BY " & vbCrLf
    sql = sql & "    YEAR(requestdate),VisitationID, InsuranceTypeID" & vbCrLf
    sql = sql & ")," & vbCrLf
    
    sql = sql & "drugsaleitemsCTE AS (" & vbCrLf
    sql = sql & "    SELECT SUM(finalamt) final_amount, VisitationID, InsuranceTypeID," & vbCrLf
    sql = sql & "           YEAR(dispensedate) saleItemYear" & vbCrLf
    sql = sql & "    FROM drugsaleitems" & vbCrLf
    sql = sql & "    GROUP BY" & vbCrLf
    sql = sql & "    YEAR(dispensedate),VisitationID, InsuranceTypeID" & vbCrLf
    sql = sql & ")," & vbCrLf
    
    sql = sql & "drugsaleitems2CTE AS (" & vbCrLf
    sql = sql & "    SELECT SUM(finalamt) final_amount, VisitationID, InsuranceTypeID," & vbCrLf
    sql = sql & "           YEAR(dispensedate) SaleItem2Year" & vbCrLf
    sql = sql & "    FROM DrugSaleItems2" & vbCrLf
    sql = sql & "    GROUP BY " & vbCrLf
    sql = sql & "    YEAR(dispensedate), VisitationID, InsuranceTypeID" & vbCrLf
    sql = sql & ")," & vbCrLf
    
    sql = sql & "aggregatedCTE AS (" & vbCrLf
    sql = sql & "    SELECT " & vbCrLf
    sql = sql & "        visitationCTE.[Year]," & vbCrLf
    sql = sql & "        SUM (ISNULL(visitationCTE.final_amount,0)+ISNULL(treatchargesCTE.final_amount,0)+" & vbCrLf
    sql = sql & "        ISNULL(investigationCTE.final_amount,0)+ISNULL(investigation2CTE.final_amount,0)+" & vbCrLf
    sql = sql & "        ISNULL(drugsaleitemsCTE.final_amount,0)+ISNULL(drugsaleitems2CTE.final_amount,0)) [TotalbyYear], InsuranceTypeName" & vbCrLf
    sql = sql & "    FROM visitationCTE" & vbCrLf
    sql = sql & "    LEFT JOIN treatchargesCTE on visitationCTE.VisitationID = treatchargesCTE.VisitationID" & vbCrLf
    sql = sql & "    LEFT JOIN investigationCTE on visitationCTE.VisitationID = investigationCTE.VisitationID" & vbCrLf
    sql = sql & "    LEFT JOIN investigation2CTE on visitationCTE.VisitationID = investigation2CTE.VisitationID" & vbCrLf
    sql = sql & "    LEFT JOIN drugsaleitemsCTE on visitationCTE.VisitationID = drugsaleitemsCTE.VisitationID" & vbCrLf
    sql = sql & "    LEFT JOIN drugsaleitems2CTE on visitationCTE.VisitationID = drugsaleitems2CTE.VisitationID" & vbCrLf
    sql = sql & "    GROUP BY visitationCTE.[Year],InsuranceTypeName" & vbCrLf
    sql = sql & ")," & vbCrLf
    
    sql = sql & "AnalysisCTE AS (" & vbCrLf
    sql = sql & "    SELECT " & vbCrLf
    sql = sql & "        TotalbyYear ," & vbCrLf
    sql = sql & "        [Year]," & vbCrLf
    sql = sql & "        LAG(TotalbyYear) OVER ( ORDER BY  [Year]) [PrevYearTote]," & vbCrLf
    sql = sql & "        (TotalbyYear - LAG(TotalbyYear) OVER ( ORDER BY  [Year])) [DiffTote]," & vbCrLf
    sql = sql & "        ((TotalbyYear - LAG(TotalbyYear) OVER ( ORDER BY  [Year])) * 100.0/TotalbyYear) PercentageChangeInTote," & vbCrLf
    sql = sql & "        (TotalbyYear*100.0/ SUM (TotalbyYear) OVER ()) [PercentContToTote]," & vbCrLf
    sql = sql & "        InsuranceTypeName" & vbCrLf
    sql = sql & "    FROM aggregatedCTE" & vbCrLf
    sql = sql & ")" & vbCrLf
    
    sql = sql & "SELECT InsuranceTypeName, [Year], TotalbyYear, [PrevYearTote], [DiffTote], PercentageChangeInTote, [PercentContToTote]" & vbCrLf
    sql = sql & "FROM AnalysisCTE"

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
            jsonData = jsonData & """InsuranceTypeName"":""" & CStr(rst.Fields("InsuranceTypeName").Value) & ""","
            jsonData = jsonData & """Year"":""" & rst.Fields("Year").Value & ""","
            jsonData = jsonData & """TotalbyYear"":""" & rst.Fields("TotalbyYear").Value & ""","
            jsonData = jsonData & """PrevYearTote"":""" & rst.Fields("PrevYearTote").Value & ""","
            jsonData = jsonData & """DiffTote"":""" & rst.Fields("DiffTote").Value & ""","
            jsonData = jsonData & """PercentageChangeInTote"":""" & rst.Fields("PercentageChangeInTote").Value & ""","
            jsonData = jsonData & """PercentContToTote"":""" & rst.Fields("PercentContToTote").Value & """"
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
    response.write "        x: revenueSourcesYearly.map(pair => pair.Year),"
    response.write "        y: revenueSourcesYearly.map(pair => pair.TotalbyYear),"
    response.write "        type: 'bar',"
    response.write "        text: revenueSourcesYearly.map(pair => 'InsuranceTypeName: ' + pair.InsuranceTypeName + '<br>Year: ' + pair.Year + '<br>Difference: ' + pair.DiffTote + '<br>Percentage Change: ' + pair.PercentageChangeInTote + '<br>Cont Percentage: ' + pair.PercentContToTote + ' '),"
    response.write "        textposition: 'auto',"
    response.write "        texttemplate: '%{y}',"
    'response.write "        hovertemplate: '%{text}',"
    response.write "        marker: {"
    response.write "            color: revenueSourcesYearly.map((_, index) => colors[index % colors.length])"
    response.write "        }"
    response.write "    };"

    ' Define the layout for the chart
    response.write "    var barLayout = {"
    response.write "        title:  'Insurance Type Analysis Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " ',"
    response.write "        xaxis: { title: 'Year' },"
    response.write "        yaxis: { title: 'Total Per Year' },"
    response.write "        height: 600, width: window.innerWidth * 0.95,"
    response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
    response.write "    };"

    ' Create the bar chart
    response.write "    Plotly.newPlot('yearlyChartDiv', [trace], barLayout);"
    response.write "});"

    ' DataTable Initialization
    response.write "    new DataTable('#yearlyTable', {"
    response.write "        data: dbDataYearly.data,"
    response.write "        columns: ["
    response.write "            { data: 'counter' },"
    response.write "            { data: 'InsuranceTypeName' },"
    response.write "            { data: 'Year' },"
    response.write "            { data: 'TotalbyYear' },"
    response.write "            { data: 'PrevYearTote' },"
    response.write "            { data: 'DiffTote' },"
    response.write "            { data: 'PercentageChangeInTote' },"
    response.write "            { data: 'PercentContToTote' }"
    response.write "        ],"
    
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: '" & brnchName & " Insurance Analysis From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: '" & brnchName & " Insurance Analysis From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: '" & brnchName & " Insurance Analysis From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'print',"
    response.write "                text: 'PRINT',"
    response.write "                title: '" & brnchName & " Insurance Analysis From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
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

Sub MultiSelectStyles()
     response.write "    <style>" & vbCrLf
    response.write "        .mult-select-tag {" & vbCrLf
    response.write "            display: flex;" & vbCrLf
    response.write "            width: 300px;" & vbCrLf
    response.write "            flex-direction: column;" & vbCrLf
    response.write "            align-items: center;" & vbCrLf
    response.write "            position: relative;" & vbCrLf
    response.write "            --tw-shadow: 0 1px 3px 0 rgb(0 0 0 / 0.1), 0 1px 2px -1px rgb(0 0 0 / 0.1);" & vbCrLf
    response.write "            --tw-shadow-color: 0 1px 3px 0 var(--tw-shadow-color), 0 1px 2px -1px var(--tw-shadow-color);" & vbCrLf
    response.write "            --border-color: rgb(218, 221, 224);" & vbCrLf
    response.write "            font-family: Verdana, sans-serif;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .wrapper {" & vbCrLf
    response.write "            width: 100%;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .body {" & vbCrLf
    response.write "            display: flex;" & vbCrLf
    response.write "            border: 1px solid var(--border-color);" & vbCrLf
    response.write "            background: #fff;" & vbCrLf
    response.write "            min-height: 2.15rem;" & vbCrLf
    response.write "            width: 100%;" & vbCrLf
    response.write "            min-width: 14rem;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .input-container {" & vbCrLf
    response.write "            display: flex;" & vbCrLf
    response.write "            flex-wrap: wrap;" & vbCrLf
    response.write "            flex: 1 1 auto;" & vbCrLf
    response.write "            padding: 0.1rem;" & vbCrLf
    response.write "            align-items: center;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .input-body {" & vbCrLf
    response.write "            display: flex;" & vbCrLf
    response.write "            width: 100%;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .input {" & vbCrLf
    response.write "            flex: 1;" & vbCrLf
    response.write "            background: 0 0;" & vbCrLf
    response.write "            border-radius: 0.25rem;" & vbCrLf
    response.write "            padding: 0.45rem;" & vbCrLf
    response.write "            margin: 10px;" & vbCrLf
    response.write "            color: #2d3748;" & vbCrLf
    response.write "            outline: 0;" & vbCrLf
    response.write "            border: 1px solid var(--border-color);" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .btn-container {" & vbCrLf
    response.write "            color: #e2ebf0;" & vbCrLf
    response.write "            padding: 0.5rem;" & vbCrLf
    response.write "            display: flex;" & vbCrLf
    response.write "            border-left: 1px solid var(--border-color);" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag button {" & vbCrLf
    response.write "            cursor: pointer;" & vbCrLf
    response.write "            width: 100%;" & vbCrLf
    response.write "            color: #718096;" & vbCrLf
    response.write "            outline: 0;" & vbCrLf
    response.write "            height: 100%;" & vbCrLf
    response.write "            border: none;" & vbCrLf
    response.write "            padding: 0;" & vbCrLf
    response.write "            background: 0 0;" & vbCrLf
    response.write "            background-image: none;" & vbCrLf
    response.write "            -webkit-appearance: none;" & vbCrLf
    response.write "            text-transform: none;" & vbCrLf
    response.write "            margin: 0;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag button:first-child {" & vbCrLf
    response.write "            width: 1rem;" & vbCrLf
    response.write "            height: 90%;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .drawer {" & vbCrLf
    response.write "            position: absolute;" & vbCrLf
    response.write "            background: #fff;" & vbCrLf
    response.write "            max-height: 15rem;" & vbCrLf
    response.write "            z-index: 40;" & vbCrLf
    response.write "            top: 98%;" & vbCrLf
    response.write "            width: 100%;" & vbCrLf
    response.write "            overflow-y: scroll;" & vbCrLf
    response.write "            border: 1px solid var(--border-color);" & vbCrLf
    response.write "            border-radius: 0.25rem;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag ul {" & vbCrLf
    response.write "            list-style-type: none;" & vbCrLf
    response.write "            padding: 0.5rem;" & vbCrLf
    response.write "            margin: 0;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag ul li {" & vbCrLf
    response.write "            padding: 0.5rem;" & vbCrLf
    response.write "            border-radius: 0.25rem;" & vbCrLf
    response.write "            cursor: pointer;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag ul li:hover {" & vbCrLf
    response.write "            background: rgb(243 244 246);" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .item-container {" & vbCrLf
    response.write "            display: flex;" & vbCrLf
    response.write "            justify-content: center;" & vbCrLf
    response.write "            align-items: center;" & vbCrLf
    response.write "            padding: 0.2rem 0.4rem;" & vbCrLf
    response.write "            margin: 0.2rem;" & vbCrLf
    response.write "            font-weight: 500;" & vbCrLf
    response.write "            border: 1px solid;" & vbCrLf
    response.write "            border-radius: 9999px;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .item-label {" & vbCrLf
    response.write "            max-width: 100%;" & vbCrLf
    response.write "            line-height: 1;" & vbCrLf
    response.write "            font-size: 0.75rem;" & vbCrLf
    response.write "            font-weight: 400;" & vbCrLf
    response.write "            flex: 0 1 auto;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .item-close-container {" & vbCrLf
    response.write "            display: flex;" & vbCrLf
    response.write "            flex: 1 1 auto;" & vbCrLf
    response.write "            flex-direction: row-reverse;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .item-close-svg {" & vbCrLf
    response.write "            width: 1rem;" & vbCrLf
    response.write "            margin-left: 0.5rem;" & vbCrLf
    response.write "            height: 1rem;" & vbCrLf
    response.write "            cursor: pointer;" & vbCrLf
    response.write "            border-radius: 9999px;" & vbCrLf
    response.write "            display: block;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .hidden {" & vbCrLf
    response.write "            display: none;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .shadow {" & vbCrLf
    response.write "            box-shadow: var(--tw-ring-offset-shadow, 0 0 #0000), var(--tw-ring-shadow, 0 0 #0000), var(--tw-shadow);" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "        .mult-select-tag .rounded {" & vbCrLf
    response.write "            border-radius: 0.375rem;" & vbCrLf
    response.write "        }" & vbCrLf
    response.write "    </style>" & vbCrLf
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
