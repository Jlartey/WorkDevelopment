VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRTLYO_MostDespensedDrugsBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
    response.Clear

    conn.CommandTimeout = 7200
    Dim periodStart, periodEnd, brnchID, jsonData, dropdownvalue
    Dim page_title
    page_title = ""

    If Len(Trim(Request.QueryString("selectedValue"))) > 1 Then
        periodStart = Trim(Request.QueryString("selectedValue"))
        periodEnd = Trim(Request.QueryString("selectedValue1"))
        brnchID = Trim(Request.QueryString("branID"))
        dropdownvalue = Trim(Request.QueryString("dropdownvalue"))
       

        dropdownvalue = CInt(dropdownvalue)
        If Len(dropdownvalue) = 0 Then
            dropdownvalue = 10
        End If

        periodStart = FormatDate(periodStart)
        periodEnd = FormatDate(periodEnd)
        brnchID = Trim(brnchID)
    Else
        periodStart = FormatDate(Now - 1)
        periodEnd = FormatDate(Now)
        brnchID = "B001"
        dropdownvalue = 10

        response.write "<script>"
        ' Response.Write "function hideSpinner() {"
        response.write "    document.getElementById('spinner').style.display = 'none';"
        ' Response.Write "}"
        response.write "</script>"
    End If

    response.write "<!DOCTYPE html>"
    response.write "<html lang=""en"">"
    response.write ""
    response.write "<head>"
    response.write "    <title>Most Dispensed Medications </title>"
    response.write "    <meta charset=""utf-8"">"
    response.write "    <meta http-equiv=""X-UA-Compatible"" content=""IE=Edge"">"
    response.write "    <meta name=""viewport"" content=""width=device-width, initial-scale=1, shrink-to-fit=no"">"
    response.write "    <!--[ Favicon]-->"
    response.write "    <link rel=""icon"" type=""image/x-icon"" href=""images/banner1.bmp"">"
    response.write "    <link rel=""icon"" type=""image/png"" sizes=""16x16"" href=""images/banner1.bmp"">"
    response.write "    <link rel=""icon"" type=""image/png"" sizes=""32x32"" href=""images/banner1.bmp"">"
    response.write "    <link rel=""apple-touch-icon"" sizes=""180x180"" href=""images/banner1.bmp"">"
    response.write ""
    response.write "    <!--[ Template main css file ]-->"
    response.write "    <link rel=""stylesheet"" href=""BusinessAnalytics/assets/css/style.min.css"">"

   
    response.write "<script src='BusinessAnalytics/assets/bundles/plotly.min.js'></script>"

    response.write "<script src='BusinessAnalytics/test3styles/styles/jsdelivrv530.js'></script>"
    response.write "<script src='BusinessAnalytics/test3styles/styles/jsdelivrselect2v410.js'></script>"
    response.write "<script src='BusinessAnalytics/test3styles/styles/jsdelivrjqueryv360.js'></script>"
    response.write "<script src='BusinessAnalytics/test3styles/styles/cloudflarev027.js'></script>"
    response.write "<script src='BusinessAnalytics/test3styles/styles/cdnjscloudflarev027.js'></script>"
    response.write "<script src='BusinessAnalytics/test3styles/styles/cdndatatablesv111.min.js'></script>"

    response.write "<script src='BusinessAnalytics/test3styles/styles/multi-select-tag.js'></script>"
    response.write "<script src='BusinessAnalytics/test3styles/styles/sweetalert2.all.min.js'></script>"

    response.write "    <link rel=""stylesheet"" href=""BusinessAnalytics/test3styles/styles/sweetalert2.min.css"">"
    response.write "    <link rel=""stylesheet"" href=""BusinessAnalytics/test3styles/styles/habibmhamadimultiselecttag.css"">"
    response.write "    <link rel=""stylesheet"" href=""BusinessAnalytics/test3styles/styles/tdatatables.min.css"">"
    response.write "    <link rel=""stylesheet"" href=""BusinessAnalytics/test3styles/styles/jsdelivrv530.css"">"
    response.write "    <link rel=""stylesheet"" href=""BusinessAnalytics/test3styles/styles/jsdelivrv502.css"">"
    response.write "    <link rel=""stylesheet"" href=""BusinessAnalytics/test3styles/styles/jsdelivrselect2v410.css"">"
     response.write "    <link rel=""stylesheet"" href=""BusinessAnalytics/assets/css/style.min.css"">"
   
   ' style for tooltip
response.write "<style>"
response.write ".custom-tooltip { position: relative; cursor: help; }"
response.write ".custom-tooltip::after {"
response.write "    content: attr(data-tooltip);"
response.write "    position: absolute;"
response.write "    bottom: 100%;"
response.write "    left: 50%;"
response.write "    transform: translateX(-50%);"
'Response.Write "    background-color: aqua;"
response.write "    background-color: #003366;"
response.write "    color: white;"
response.write "    padding: 8px 12px;"
response.write "    font-size: 14px;"
response.write "    border-radius: 4px;"
response.write "    white-space: pre-wrap;"
response.write "    z-index: 9999;"
response.write "    display: none;"
response.write "    min-width: 320px;"
response.write "    box-shadow: 0 4px 12px rgba(0,0,0,0.25);"
response.write "    line-height: 1.4;"
response.write "}"
response.write ".custom-tooltip:hover::after {"
response.write "    display: block;"
response.write "}"
response.write "</style>"

   
   
    response.write "</head>"
    response.write "<body data-theme=""theme-PurpleHeart"" class=""svgstroke-a bg-gradient"">"
    ' loading----------------------------------------------------------------------------------------------
    response.write "<style>"
    response.write "    #spinner {"
    response.write "        position: fixed;"
    response.write "        top: 0;"
    response.write "        left: 0;"
    response.write "        width: 100%;"
    response.write "        height: 100%;"
    ' Response.Write "        background: rgba(255, 255, 255, 0.7);"
    response.write "        display: none;" ' Initially hidden
    response.write "        align-items: center;"
    response.write "        justify-content: center;"
    response.write "        backdrop-filter: blur(4px);"
    response.write "        z-index: 9999;"
    response.write "    }"
    response.write "    .spinner-border {"
    response.write "        width: 3rem;"
    response.write "        height: 3rem;"
    response.write "    }"
    response.write "    .spinner-grow {"
    response.write "        width: 3rem;"
    response.write "        height: 3rem;"
    response.write "        margin-left: 1rem;"
    response.write "    }"
    response.write "</style>"

    response.write "<div id='spinner'>"
    response.write "    <div class=""spinner-border"" role=""status"">"
    response.write "        <span class=""visually-hidden"">Loading...</span>"
    response.write "    </div>"
    response.write "    <div class=""spinner-grow"" role=""status"">"
    response.write "        <span class=""visually-hidden"">Loading...</span>"
    response.write "    </div>"
    response.write "</div>"
    ShowLoading
    ' ----------------------------------------------------------------------------------------------
    InitPageScript
    ' ----------------------------------------------------------------------------------------------
    response.write ""
    response.write "    <!-- <main class=""container-fluid px-0""> -->"
    ' ----------------------------------------------------------------------------------------------
    filters
    ' ----------------------------------------------------------------------------------------------
    response.write "            <div class=""card bg-transparent border-0"">"
    response.write "                <ul class=""nav nav-tabs li_animate"" role=""tablist"">"
    response.write "                    <li class=""nav-item"" role=""presentation""><a class=""nav-link active"" data-bs-toggle=""tab"" href=""#Yearly"" aria-selected=""true"" role=""tab""><span>Yearly Dispensing Trends</span></a></li>"
    response.write "                    <li class=""nav-item"" role=""presentation""><a class=""nav-link"" data-bs-toggle=""tab"" href=""#Quarterly"" aria-selected=""false"" tabindex=""-1"" role=""tab""><span class=""d-none d-sm-inline-flex ps-1"">Quarterly Dispensing Trends</span></a></li>"
    response.write "                    <li class=""nav-item"" role=""presentation""><a class=""nav-link"" data-bs-toggle=""tab"" href=""#Monthly"" aria-selected=""false"" tabindex=""-1"" role=""tab""><span class=""d-none d-sm-inline-flex ps-1"">Monthly Dispensing Trends</span></a></li>"
    response.write "                    <li class=""nav-item"" role=""presentation""><a class=""nav-link"" data-bs-toggle=""tab"" href=""#Weekly"" aria-selected=""false"" tabindex=""-1"" role=""tab""><span class=""d-none d-sm-inline-flex ps-1"">Weekly Dispensing Trends</span></a></li>"
    response.write "                </ul>"
    response.write "                <div class=""tab-content bg-card p-3 border border-top-0 rounded-bottom"">"
    response.write "                    <!--[ Start:: Yearly]-->"
    response.write "                    <div class=""tab-pane fade show active"" id=""Yearly"" role=""tabpanel"">"
    response.write "                        <ul class=""row g-3 list-unstyled li_animate mb-0 "">"
    response.write "                            <div class=""col-12 mb-5 "">"
    
    get_annual_most_dispensed_medications

    response.write "                            </div>"
    response.write "                        </ul>"
    response.write "                    </div>"
    response.write "                    <!--[ Start:: Quarterly]-->"
    response.write "                    <div class=""tab-pane fade "" id=""Quarterly"" role=""tabpanel"">"
    response.write "                        <ul class=""row g-3 list-unstyled li_animate mb-0 "">"
    response.write "                            <div class=""col-12 mb-5 "">"
    
    quarterly_most_dispensed_medications

    response.write "                            </div>"
    response.write "                        </ul>"
    response.write "                    </div>"
    response.write "                    <!--[ Start:: Monthly]-->"
    response.write "                    <div class=""tab-pane fade "" id=""Monthly"" role=""tabpanel"">"
    response.write "                        <ul class=""row g-3 list-unstyled li_animate mb-0 "">"
    response.write "                            <div class=""col-12 mb-5 "">"

    monthly_most_dispensed_medications

    response.write "                            </div>"
    response.write "                        </ul>"
    response.write "                    </div>"
    response.write "                    <!--[ Start:: Weekly]-->"
    response.write "                    <div class=""tab-pane fade "" id=""Weekly"" role=""tabpanel"">"
    response.write "                        <ul class=""row g-3 list-unstyled li_animate mb-0 "">"
    response.write "                            <div class=""col-12 mb-5 "">"

    weekly_most_dispensed_medications

    response.write "                            </div>"
    response.write "                        </ul>"
    response.write "                    </div>"
    response.write "                </div>"
    response.write "            </div>"
    response.write ""
    response.write "        </div>"
    response.write ""
    response.write "        </div>"
    response.write "        </div>"
    response.write ""
    response.write "    </main>"
    response.write "    <script src=""BusinessAnalytics/assets/bundles/libscripts.bundle.js ""></script>"
    response.write "    <script src=""BusinessAnalytics/assets/js/main.js ""></script>"
    response.write ""
    response.write "</body>"
    response.write ""
    response.write "</html>"

Sub get_annual_most_dispensed_medications()

    response.write "  <div class='chart-container'>"
    response.write "    <div id='yearlyChartDivGender' class='chart table-responsive'>"
    response.write "  </div>"
    response.write "   <div class=""card"">"
    response.write "   <div class=""card-body"">"
    response.write "       <div class='table-responsive'>"
    response.write "      <table class=""table dataTable align-middle table-hover table-bordered"" style=""width: 100%;"" id=""interVisitTable"" >"
    response.write "      <thead >"
    response.write "              <tr>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
     response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "              </tr>"
    response.write "            </thead>"
    response.write "        </table>"
    response.write "       </div>"
    response.write "      </div>"
    response.write "  </div>"

Dim sql, rst, cmd
    Set rst = CreateObject("ADODB.RecordSet")
    Set cmd = CreateObject("ADODB.Command")

    ' sql = "exec sp_yearly_most_dispensed_medications '" & brnchID & "','" & periodStart & "','" & periodEnd & "'    "
    sql = "exec sp_yearly_most_dispensed_medications '" & brnchID & "','" & periodStart & "','" & periodEnd & "', " & dropdownvalue & ""


    ' Set up the Command object
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1
    Set rst = cmd.Execute

Dim jsonAnnualData

If Not rst.EOF And Not IsNull(rst.Fields(0).value) Then
    
    If Trim(rst.Fields(0).value) <> "" Then
        jsonAnnualData = "{""data"":" & rst.Fields(0).value & "}"
    Else
        jsonAnnualData = "{""data"":[]}"
    End If
Else
    jsonAnnualData = "{""data"":[]}"
End If



rst.Close
Set rst = Nothing

'##################################################################################
    ' Send the data to the client-side


response.write "<script>"
response.write "var dbDataYearlyGender = " & jsonAnnualData & ";"
response.write "document.addEventListener('DOMContentLoaded', function() {"
response.write "    var data = dbDataYearlyGender.data;"

' Unique years and drugs
response.write "    var years = [...new Set(data.map(d => d.DispenseYear))];"
response.write "    var drugNames = [...new Set(data.map(d => d.DrugName))];"

response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"

response.write "    var traces = drugNames.map(function(drug, index) {"
response.write "        var yValues = years.map(function(year) {"
response.write "            var entry = data.find(function(d) { return d.DrugName === drug && d.DispenseYear === year; });"
response.write "            return entry ? parseFloat(entry.YearQty) : 0;"
response.write "        });"

response.write "        var hoverText = years.map(function(year) {"
response.write "            var entry = data.find(function(d) { return d.DrugName === drug && d.DispenseYear === year; });"
response.write "            return entry ? "
response.write "                'Year: ' + entry.DispenseYear + '<br>' +"
response.write "                'Drug: ' + entry.DrugName + '<br>' +"
response.write "                'Quantity Dispensed: ' + parseFloat(entry.YearQty).toLocaleString() + '<br>' +"
response.write "                'Prev. Year\'s Quantity: ' + (entry.PrevYearQty ? parseFloat(entry.PrevYearQty).toLocaleString() : 'N/A') + '<br>' +"

response.write "                'Difference: ' + (entry.AbsoluteChange ? parseFloat(entry.AbsoluteChange).toLocaleString() : 'N/A') + '<br>' +"

response.write "                '% Change: ' + (entry.PercentageChange ? parseFloat(entry.PercentageChange).toLocaleString() + '%' : 'N/A') + '<br>' +"

response.write "                'Cumulative Quantity: ' + parseFloat(entry.CumulativeQty).toLocaleString() : 'No data';"
response.write "        });"

response.write "        return {"
response.write "            x: years,"
response.write "            y: yValues,"
response.write "            name: drug,"
response.write "            type: 'bar',"
response.write "            marker: { color: colors[index % colors.length] },"
response.write "            text: yValues.map(v => v.toLocaleString()),"
response.write "            textposition: 'auto',"
response.write "            hovertemplate: hoverText"
response.write "        };"
response.write "    });"

' Theme support
response.write "    function getCurrentTheme() {"
response.write "        return localStorage.getItem('theme') || 'light';"
response.write "    }"
response.write "    var currentTheme = getCurrentTheme();"
response.write "    var textColor = currentTheme === 'dark' ? '#ffffff' : '#000000';"
response.write "    var chartBackgroundColor = currentTheme === 'dark' ? '#333' : '#f0f0f0';"
response.write "    var gridColorTheme = currentTheme === 'dark' ? '#444' : '#ccc';"

' Layout
response.write "    var layout = {"
response.write "        title: {"
response.write "            text: 'Yearly Dispensing Trends Between " & FormatDateNew(periodStart) & " and " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "',"
response.write "            font: { color: textColor }"
response.write "        },"
response.write "        font: { color: textColor },"
response.write "        xaxis: {"
response.write "            title: 'Year',"
response.write "            type: 'category',"
response.write "            titlefont: { color: textColor },"
response.write "            tickfont: { color: textColor },"
response.write "            gridcolor: gridColorTheme"
response.write "        },"
response.write "        yaxis: {"
response.write "            title: 'Quantity Dispensed',"
response.write "            titlefont: { color: textColor },"
response.write "            tickfont: { color: textColor },"
response.write "            gridcolor: gridColorTheme"
response.write "        },"
response.write "        barmode: 'group',"
response.write "        paper_bgcolor: chartBackgroundColor,"
response.write "        plot_bgcolor: chartBackgroundColor,"
response.write "        height: 600,"
response.write "        width: window.innerWidth * 0.95,"
response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
response.write "        legend: {"
response.write "            orientation: 'h',"
response.write "            x: 0.5,"
response.write "            xanchor: 'center',"
response.write "            y: -0.2,"
response.write "            yanchor: 'top',"
response.write "            font: { color: textColor }"
response.write "        }"
response.write "    };"

response.write "    Plotly.newPlot('yearlyChartDivGender', traces, layout);"
response.write "});"
response.write "</script>"



'##################################################################################

    ' DataTable Initialization
    response.write "<script>"
    response.write "    new DataTable('#interVisitTable', {"
    response.write "        data: dbDataYearlyGender.data,"
    response.write "        columns: ["
    response.write "            { data: 'counter',title:'S/No.' },"
    response.write "            { data: 'DispenseYear',title:'Year' },"
'    response.write "            { data: 'DrugName',title:'Drug' },"
    
    response.write "            {"
response.write "                data: 'DrugName',"
response.write "                title: 'Drug',"
response.write "                render: function(data, type, row) {"
response.write "                    if (type === 'display') {"
response.write "                        var tooltip = row.PrescriptionNarrative ? row.PrescriptionNarrative.replace(/\""/g, '&quot;') : '';"
response.write "                        var displayValue = data;"
response.write "                        return '<span class=\""custom-tooltip\"" data-tooltip=\""' + tooltip + '\"">' + displayValue + '</span>';"
response.write "                    }"
response.write "                    return data;"
response.write "                }"
response.write "            },"

    
    
     response.write "            { data: 'YearQtyFormatted',title:'Qty Dispensed' },"
   response.write "            { data: 'PrevYearQty', title:'Prev.  Year\'s Qty', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
   response.write "            { data: 'AbsoluteChange', title:'Difference', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
               
'    response.write "            { data: 'PercentageChange',title:'% Change' },"
response.write "            { data: 'PercentageChange', title:'% Change', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
    response.write "            { data: 'CumulativeQty', title:'Cumulative Qty', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
    response.write "            { data: 'PrescriptionNarrative',title:'Prescription Narrative',visible:false },"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: 'Yearly Dispensing Trends Between : " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: 'Yearly Dispensing Trends Between : " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: 'Yearly Dispensing Trends Between : " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
    response.write "            },"
'    response.write "            {"
'    response.write "                extend: 'print',"
'    response.write "                text: 'PRINT',"
'    response.write "                title: 'Yearly Dispensing Trends Between : " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
'    response.write "            }"
    response.write "        ]"
    response.write "    });"
    response.write "</script>"
    
End Sub

Sub quarterly_most_dispensed_medications()

    response.write "  <div class='chart-container'>"
    response.write "    <div id='quarterlyChartDiv' class='chart table-responsive'>"
    response.write "  </div>"
    response.write "   <div class=""card"">"
    response.write "   <div class=""card-body"">"
    response.write "       <div class='table-responsive'>"
    response.write "      <table class=""table dataTable align-middle table-hover table-bordered"" style=""width: 100%;"" id=""proportionTable"" >"
    response.write "      <thead >"
    response.write "              <tr>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
'    response.write "                <th></th>"
    response.write "               </tr>"
    response.write "            </thead>"
    response.write "        </table>"
    response.write "       </div>"
    response.write "      </div>"
    response.write "  </div>"


Dim sql, rst, cmd
    Set rst = CreateObject("ADODB.RecordSet")
    Set cmd = CreateObject("ADODB.Command")

    ' sql = "exec sp_quarterly_most_dispensed_medications '" & brnchID & "','" & periodStart & "','" & periodEnd & "'"
    sql = "exec sp_quarterly_most_dispensed_medications '" & brnchID & "','" & periodStart & "','" & periodEnd & "', " & dropdownvalue & ""


    ' Set up the Command object
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1
    Set rst = cmd.Execute

Dim jsonQuarterData

If Not rst.EOF And Not IsNull(rst.Fields(0).value) Then
    
    If Trim(rst.Fields(0).value) <> "" Then
        jsonQuarterData = "{""data"":" & rst.Fields(0).value & "}"
    Else
        jsonQuarterData = "{""data"":[]}"
    End If
Else
    jsonQuarterData = "{""data"":[]}"
End If


rst.Close
Set rst = Nothing

'##################################################################################

response.write "<script>"
response.write "var dbDataQuarterlyGender = " & jsonQuarterData & ";"
response.write "document.addEventListener('DOMContentLoaded', function() {"

response.write "    var revenueSourcesQuarterly = dbDataQuarterlyGender.data;"

' Color palette for drug traces
response.write "    var colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'];"

response.write "    var quarters = [...new Set(revenueSourcesQuarterly.map(d => d.DispenseYear + ' Q' + d.DispenseQuarter))];"

' Sort chronologically
response.write "    quarters.sort((a, b) => {"
response.write "        var [yearA, qA] = a.split(' Q').map(Number);"
response.write "        var [yearB, qB] = b.split(' Q').map(Number);"
response.write "        return yearA - yearB || qA - qB;"
response.write "    });"

' Unique drug names
response.write "    var drugNames = [...new Set(revenueSourcesQuarterly.map(d => d.DrugName))];"

' Build a trace per drug
response.write "    var traces = drugNames.map(function(drug, index) {"
response.write "        return {"
response.write "            name: drug,"
response.write "            type: 'bar',"
response.write "            marker: { color: colors[index % colors.length] },"
response.write "            x: quarters,"
response.write "            y: quarters.map(function(q) {"
response.write "                var item = revenueSourcesQuarterly.find(function(d) {"
response.write "                    return d.DrugName === drug && (d.DispenseYear + ' Q' + d.DispenseQuarter) === q;"
response.write "                });"
response.write "                return item ? parseFloat(item.QuarterQty) : 0;"
response.write "            }),"
response.write "            text: quarters.map(function(q) {"
response.write "                var item = revenueSourcesQuarterly.find(function(d) {"
response.write "                    return d.DrugName === drug && (d.DispenseYear + ' Q' + d.DispenseQuarter) === q;"
response.write "                });"
response.write "                return item ? parseFloat(item.QuarterQty).toLocaleString() : '';"
response.write "            }),"
response.write "            hovertemplate: quarters.map(function(q) {"
response.write "                var item = revenueSourcesQuarterly.find(function(d) {"
response.write "                    return d.DrugName === drug && (d.DispenseYear + ' Q' + d.DispenseQuarter) === q;"
response.write "                });"
response.write "                return item ? "
response.write "                    'Year: ' + item.DispenseYear + '<br>' +"
response.write "                    'Quarter: Q' + item.DispenseQuarter + '<br>' +"
response.write "                    'Drug: ' + item.DrugName + '<br>' +"
response.write "                    'Quantity Dispensed: ' + parseFloat(item.QuarterQty).toLocaleString() + '<br>' +"
response.write "                    'Prev. Year Qty: ' + (item.PrevYearQty ? parseFloat(item.PrevYearQty).toLocaleString() : 'N/A') + '<br>' +"
response.write "                    'Difference: ' + (item.AbsoluteChange ? parseFloat(item.AbsoluteChange).toLocaleString() : 'N/A') + '<br>' +"
response.write "                    '% Change: ' + (item.PercentageChange ? parseFloat(item.PercentageChange).toLocaleString() + '%' : 'N/A') + '<br>' +"
response.write "                    'Cumulative Qty: ' + parseFloat(item.CumulativeQty).toLocaleString()"
response.write "                    : 'No data';"
response.write "            }),"
response.write "            textposition: 'auto'"
response.write "        };"
response.write "    });"

' Theme support
response.write "    function getCurrentTheme() {"
response.write "        return localStorage.getItem('theme') || 'light';"
response.write "    }"
response.write "    var currentTheme = getCurrentTheme();"
response.write "    var textColor = currentTheme === 'dark' ? '#ffffff' : '#000000';"
response.write "    var chartBackgroundColor = currentTheme === 'dark' ? '#333' : '#f0f0f0';"
response.write "    var gridColorTheme = currentTheme === 'dark' ? '#444' : '#ccc';"

' Layout
response.write "    var layout = {"
response.write "        title: {"
response.write "            text: 'Quarterly Dispensing Trends<br>Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "',"
response.write "            font: { color: textColor }"
response.write "        },"
response.write "        font: { color: textColor },"
response.write "        xaxis: {"
response.write "            title: '',"
response.write "            type: 'category',"
response.write "            titlefont: { color: textColor },"
response.write "            tickfont: { color: textColor },"
response.write "            gridcolor: gridColorTheme"
response.write "        },"
response.write "        yaxis: {"
response.write "            title: 'Dispensed Quantity',"
response.write "            titlefont: { color: textColor },"
response.write "            tickfont: { color: textColor },"
response.write "            gridcolor: gridColorTheme"
response.write "        },"
response.write "        paper_bgcolor: chartBackgroundColor,"
response.write "        plot_bgcolor: chartBackgroundColor,"
response.write "        height: 450,"
response.write "        width: window.innerWidth * 0.98,"
response.write "        barmode: 'stack',"
response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
response.write "        legend: {"
response.write "            orientation: 'h',"
response.write "            x: 0.5,"
response.write "            xanchor: 'center',"
response.write "            y: -0.2,"
response.write "            yanchor: 'top',"
response.write "            font: { color: textColor }"
response.write "        }"
response.write "    };"

response.write "    Plotly.newPlot('quarterlyChartDiv', traces, layout);"
response.write "});"
response.write "</script>"



'##################################################################################

    ' DataTable Initialization
   response.write "<script>"
response.write "    document.addEventListener('DOMContentLoaded', function() {"

response.write "        var dbDataQuarterlyGender = " & jsonQuarterData & "; "
response.write ""

response.write "        new DataTable('#proportionTable', {"
response.write "            data: dbDataQuarterlyGender.data,"
response.write "            columns: ["
response.write "            { data: 'counter',title:'S/No.' },"

response.write "            {"
response.write "                data: null,"
response.write "                title: 'Period',"
response.write "                render: function(data, type, row) {"
response.write "                    return row.DispenseYear + ' Q' + row.DispenseQuarter;"
response.write "                }"
response.write "            },"


'response.write "            { data: 'DrugName',title:'Drug' },"

response.write "            {"
response.write "                data: 'DrugName',"
response.write "                title: 'Drug',"
response.write "                render: function(data, type, row) {"
response.write "                    if (type === 'display') {"
response.write "                        var tooltip = row.PrescriptionNarrative ? row.PrescriptionNarrative.replace(/\""/g, '&quot;') : '';"
response.write "                        var displayValue = data;"
response.write "                        return '<span class=\""custom-tooltip\"" data-tooltip=\""' + tooltip + '\"">' + displayValue + '</span>';"
response.write "                    }"
response.write "                    return data;"
response.write "                }"
response.write "            },"

response.write "            { data: 'QuarterQty', title:'Qty Dispensed', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
response.write "            { data: 'PrevQuarterQty', title:'Prev.  Quarter\'s Qty', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
               
response.write "            { data: 'AbsoluteChange', title:'Difference', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"

response.write "            { data: 'PercentageChange', title:'% Change', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
    response.write "            { data: 'CumulativeQty', title:'Cumulative Qty', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
    response.write "            { data: 'PrescriptionNarrative',title:'Prescription Narrative', visible:false },"
response.write "            ],"
response.write "            lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
response.write "            dom: 'lBfrtip',"
response.write "            buttons: ["
response.write "                {"
response.write "                    extend: 'csv',"
response.write "                    text: 'CSV',"
response.write "                    title: 'Quarterly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
response.write "                },"
response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: 'Quarterly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
response.write "                },"
response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: 'Quarterly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
response.write "                },"
'response.write "                {"
'response.write "                    extend: 'print',"
'response.write "                    text: 'PRINT',"
'response.write "                    title: 'Quarterly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
'response.write "                }"
response.write "            ]"
response.write "        });"
response.write "    });"
response.write "</script>"

    
End Sub

Sub monthly_most_dispensed_medications()

        response.write "  <div class='chart-container'>"
        response.write "    <div id='monthlyChartDiv' class='chart table-responsive'>"
        response.write "  </div>"
        response.write "   <div class=""card"">"
        response.write "   <div class=""card-body"">"
        response.write "       <div class='table-responsive'>"
        response.write "      <table class=""table dataTable align-middle table-hover table-bordered"" style=""width: 100%;"" id=""monthlyTable"" >"
        response.write "      <thead >"
        response.write "              <tr>"
        response.write "                <th></th>"
        response.write "                <th></th>"
        response.write "                <th></th>"
        response.write "                <th></th>"
        response.write "                <th></th>"
        response.write "                <th></th>"
        response.write "                <th></th>"
        response.write "                <th></th>"
        response.write "                <th></th>"
        response.write "                <th></th>"
        response.write "              </tr>"
        response.write "            </thead>"
        response.write "        </table>"
        response.write "       </div>"
        response.write "      </div>"
        response.write "  </div>"

 
 Dim sql, rst, cmd
    Set rst = CreateObject("ADODB.RecordSet")
    Set cmd = CreateObject("ADODB.Command")

    ' sql = "exec sp_monthly_most_dispensed_medications '" & brnchID & "','" & periodStart & "','" & periodEnd & "'"
    sql = "exec sp_monthly_most_dispensed_medications '" & brnchID & "','" & periodStart & "','" & periodEnd & "', " & dropdownvalue & ""

    
    ' Set up the Command object
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1
    Set rst = cmd.Execute

Dim jsonMonthlyData

If Not rst.EOF And Not IsNull(rst.Fields(0).value) Then
    
    If Trim(rst.Fields(0).value) <> "" Then
        jsonMonthlyData = "{""data"":" & rst.Fields(0).value & "}"
    Else
        jsonMonthlyData = "{""data"":[]}"
    End If
Else
    jsonMonthlyData = "{""data"":[]}"
End If


rst.Close
Set rst = Nothing
 
response.write "<script>"
response.write "var dbDataGenderMonthly = " & jsonMonthlyData & ";"
response.write "document.addEventListener('DOMContentLoaded', function() {"
response.write "    var data = dbDataGenderMonthly.data;"

' Month names
response.write "    var monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];"

' Extract unique sorted months
response.write "    var uniqueMonths = [...new Set(data.map(d => +d.DispenseMonth))].sort((a, b) => a - b);"
response.write "    var periods = uniqueMonths.map(m => ({ mon: m, label: monthNames[m - 1] }));"

' Extract unique (drug + year) combinations
response.write "    var drugYearKeys = [...new Set(data.map(d => d.DrugName + '|' + d.DispenseYear))];"

' Color palette
response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"

' Build traces
response.write "    var traces = drugYearKeys.map((key, i) => {"
response.write "        var parts = key.split('|');"
response.write "        var drug = parts[0];"
response.write "        var year = parts[1];"
response.write "        var y = [], hover = [];"
response.write "        periods.forEach(p => {"
response.write "            var match = data.find(d => d.DrugName === drug && d.DispenseYear.toString() === year && +d.DispenseMonth === p.mon);"
response.write "            if (match) {"
response.write "                y.push(parseFloat(match.QuarterQty));"
response.write "                hover.push("
response.write "                    'Drug: ' + match.DrugName + '<br>' +"
response.write "                    'Year: ' + match.DispenseYear + '<br>' +"
response.write "                    'Month: ' + p.label + '<br>' +"
response.write "                    'Dispensed Quantity: ' + parseFloat(match.QuarterQty).toLocaleString() + '<br>' +"
response.write "                    'Prev. Month\'s Qty: ' + (isNaN(parseFloat(match.PrevQuarterQty)) || match.PrevQuarterQty === null ? 'NA' : parseFloat(match.PrevQuarterQty).toLocaleString()) + '<br>' +"
response.write "                    'Difference: ' + (isNaN(parseFloat(match.AbsoluteChange)) || match.AbsoluteChange === null ? 'NA' : parseFloat(match.AbsoluteChange).toLocaleString()) + '<br>' +"
response.write "                    '% Change: ' + (isNaN(parseFloat(match.PercentageChange)) || match.PercentageChange === null ? 'NA' : parseFloat(match.PercentageChange).toLocaleString()+'%') + '<br>' +"
response.write "                    'Cumulative Qty: ' + parseFloat(match.CumulativeQty).toLocaleString()"
response.write "                );"
response.write "            } else {"
response.write "                y.push(null);"
response.write "                hover.push('No data');"
response.write "            }"
response.write "        });"
response.write "        return {"
response.write "            x: periods.map(p => p.label),"
response.write "            y: y,"
response.write "            type: 'scatter',"
response.write "            mode: 'lines+markers',"
response.write "            name: drug + ' ' + year,"
response.write "            text: hover,"
response.write "            hovertemplate: '%{text}',"
response.write "            marker: { color: colors[i % colors.length] }"
response.write "        };"
response.write "    });"

' Theme detection
response.write "    var theme = localStorage.getItem('theme') || 'light';"
response.write "    var textColor = theme === 'dark' ? '#fff' : '#000';"
response.write "    var bg = theme === 'dark' ? '#333' : '#f0f0f0';"
response.write "    var grid = theme === 'dark' ? '#444' : '#ccc';"

' Layout
response.write "    var layout = {"
response.write "        title: {"
response.write "            text: 'Monthly Dispensing Trends<br>Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "',"
response.write "            font: { color: textColor }"
response.write "        },"
response.write "        font: { color: textColor },"
response.write "        xaxis: {"
response.write "            title: '',"
response.write "            tickangle: 0,"
response.write "            gridcolor: grid"
response.write "        },"
response.write "        yaxis: {"
response.write "            title: 'Dispensed Quantity',"
response.write "            gridcolor: grid"
response.write "        },"
response.write "        paper_bgcolor: bg,"
response.write "        plot_bgcolor: bg,"
response.write "        legend: {"
response.write "            orientation: 'h',"
response.write "            x: 0.5,"
response.write "            y: -0.2,"
response.write "            xanchor: 'center',"
response.write "            font: { color: textColor }"
response.write "        },"
response.write "        height: 600,"
response.write "        width: window.innerWidth * 0.98"
response.write "    };"

' Render the chart
response.write "    Plotly.newPlot('monthlyChartDiv', traces, layout);"
response.write "});"
response.write "</script>"

 '####################################################################################################
    

    ' DataTable Initialization
    response.write "<script>"
    response.write "    new DataTable('#monthlyTable', {"
    response.write "        data: dbDataGenderMonthly.data,"
    response.write "        columns: ["
    response.write "            { data: 'counter',title:'S/No.' },"
    response.write "            { data: 'DispenseYear' ,title:'Year'},"
    response.write "            { data: 'DispenseMonthName',title:'Month' },"
'    response.write "            { data: 'DrugName',title:'Drug' },"
response.write "            {"
response.write "                data: 'DrugName',"
response.write "                title: 'Drug',"
response.write "                render: function(data, type, row) {"
response.write "                    if (type === 'display') {"
response.write "                        var tooltip = row.PrescriptionNarrative ? row.PrescriptionNarrative.replace(/\""/g, '&quot;') : '';"
response.write "                        var displayValue = data;"
response.write "                        return '<span class=\""custom-tooltip\"" data-tooltip=\""' + tooltip + '\"">' + displayValue + '</span>';"
response.write "                    }"
response.write "                    return data;"
response.write "                }"
response.write "            },"

    response.write "            { data: 'QuarterQty', title:'Qty Dispensed', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
   response.write "            { data: 'PrevQuarterQty', title:'Prev.  Month\'s Qty', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
               
   response.write "            { data: 'AbsoluteChange', title:'Difference', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"

response.write "            { data: 'PercentageChange', title:'% Change', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
    response.write "            { data: 'CumulativeQty', title:'Cumulative Qty', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
    response.write "            { data: 'PrescriptionNarrative',title:'Prescription Narrative',visible:false },"
    response.write "        ],"
    response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    response.write "        dom: 'lBfrtip',"
    response.write "        buttons: ["
    response.write "            {"
    response.write "                extend: 'csv',"
    response.write "                text: 'CSV',"
    response.write "                title: 'Monthly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'excel',"
    response.write "                text: 'EXCEL',"
    response.write "                title: 'Monthly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
    response.write "            },"
    response.write "            {"
    response.write "                extend: 'pdf',"
    response.write "                text: 'PDF',"
    response.write "                title: 'Monthly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "',"
    response.write "            },"
'    response.write "            {"
'    response.write "                extend: 'print',"
'    response.write "                text: 'PRINT',"
'    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
'    response.write "            }"
    response.write "        ]"
    response.write "    });"
    response.write "</script>"
    
End Sub

Sub weekly_most_dispensed_medications()
   

    response.write "  <div class='chart-container'>"
    response.write "    <div id='weeklyVisitsChartDiv' class='chart table-responsive'>"
    response.write "  </div>"
    response.write "   <div class=""card"">"
    response.write "   <div class=""card-body"">"
    response.write "       <div class='table-responsive'>"
    response.write "      <table class=""table dataTable align-middle table-hover table-bordered"" style=""width: 100%;"" id=""weeklyTable"" >"
    response.write "      <thead >"
    response.write "              <tr>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "                <th></th>"
    response.write "              </tr>"
    response.write "            </thead>"
    response.write "        </table>"
    response.write "       </div>"
    response.write "      </div>"
    response.write "  </div>"

   
   Dim sql, rst, cmd
    Set rst = CreateObject("ADODB.RecordSet")
    Set cmd = CreateObject("ADODB.Command")

    ' sql = "exec sp_weekly_most_dispensed_medications '" & brnchID & "','" & periodStart & "','" & periodEnd & "'"
    sql = "exec sp_weekly_most_dispensed_medications '" & brnchID & "','" & periodStart & "','" & periodEnd & "', " & dropdownvalue & ""


    ' Set up the Command object
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1
    Set rst = cmd.Execute

Dim jsonWeeklyData

If Not rst.EOF And Not IsNull(rst.Fields(0).value) Then
    
    If Trim(rst.Fields(0).value) <> "" Then
        jsonWeeklyData = "{""data"":" & rst.Fields(0).value & "}"
    Else
        jsonWeeklyData = "{""data"":[]}"
    End If
Else
    jsonWeeklyData = "{""data"":[]}"
End If


rst.Close
Set rst = Nothing

'####################################################################################################

response.write "<script>"
response.write "var dbDataWeekly = " & jsonWeeklyData & ";"
response.write "document.addEventListener('DOMContentLoaded', function() {"
response.write "    var data = dbDataWeekly.data;"

' Extract unique sorted week numbers
response.write "    var uniqueWeeks = [...new Set(data.map(d => +d.DispenseWeek))].sort((a, b) => a - b);"
response.write "    var weeks = uniqueWeeks.map(w => ({ wk: w, label:  w }));"

' Extract unique (DrugName + Year) combinations for legend
response.write "    var drugYearKeys = [...new Set(data.map(d => d.DrugName + '|' + d.DispenseYear))];"

' Color palette
response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"

' Build traces per drug-year
response.write "    var traces = drugYearKeys.map((key, i) => {"
response.write "        var parts = key.split('|');"
response.write "        var drug = parts[0];"
response.write "        var year = parts[1];"
response.write "        var y = [], hover = [];"
response.write "        weeks.forEach(w => {"
response.write "            var match = data.find(d => d.DrugName === drug && d.DispenseYear.toString() === year && +d.DispenseWeek === w.wk);"
response.write "            if (match) {"
response.write "                y.push(parseFloat(match.QuarterQty));"
response.write "                hover.push("
response.write "                    'Drug: ' + match.DrugName + '<br>' +"
response.write "                    'Year: ' + match.DispenseYear + '<br>' +"
response.write "                    'Week: ' + w.label + '<br>' +"
response.write "                    'Dispensed Quantity: ' + parseFloat(match.QuarterQty).toLocaleString() + '<br>' +"
response.write "                    'Prev. Week\'s Qty: ' + (isNaN(parseFloat(match.PrevQuarterQty)) || match.PrevQuarterQty === null ? 'NA' : parseFloat(match.PrevQuarterQty).toLocaleString()) + '<br>' +"
response.write "                    'Difference: ' + (isNaN(parseFloat(match.AbsoluteChange)) || match.AbsoluteChange === null ? 'NA' : parseFloat(match.AbsoluteChange).toLocaleString()) + '<br>' +"
response.write "                    '% Change: ' + (isNaN(parseFloat(match.PercentageChange)) || match.PercentageChange === null ? 'NA' : parseFloat(match.PercentageChange).toLocaleString() + '%') + '<br>' +"
response.write "                    'Cumulative Qty: ' + (isNaN(parseFloat(match.CumulativeQty)) || match.CumulativeQty === null ? 'NA' : parseFloat(match.CumulativeQty).toLocaleString())"
response.write "                );"
response.write "            } else {"
response.write "                y.push(null);"
response.write "                hover.push('No data');"
response.write "            }"
response.write "        });"
response.write "        return {"
response.write "            x: weeks.map(w => w.label),"
response.write "            y: y,"
response.write "            type: 'scatter',"
response.write "            mode: 'lines+markers',"
response.write "            name: drug + ' ' + year,"
response.write "            text: hover,"
response.write "            hovertemplate: '%{text}',"
response.write "            marker: { color: colors[i % colors.length] }"
response.write "        };"
response.write "    });"

' Theme handling
response.write "    var theme = localStorage.getItem('theme') || 'light';"
response.write "    var textColor = theme === 'dark' ? '#fff' : '#000';"
response.write "    var bg = theme === 'dark' ? '#333' : '#f0f0f0';"
response.write "    var grid = theme === 'dark' ? '#444' : '#ccc';"

' Layout configuration
response.write "    var layout = {"
response.write "        title: {"
response.write "            text: 'Weekly Dispensing Trends<br>Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "',"
response.write "            font: { color: textColor }"
response.write "        },"
response.write "        font: { color: textColor },"
response.write "        xaxis: {"
response.write "            title: 'Week Number',"
response.write "            tickangle: 0,"
response.write "            gridcolor: grid"
response.write "        },"
response.write "        yaxis: {"
response.write "            title: 'Dispensed Quantity',"
response.write "            gridcolor: grid"
response.write "        },"
response.write "        paper_bgcolor: bg,"
response.write "        plot_bgcolor: bg,"
response.write "        legend: {"
response.write "            orientation: 'h',"
response.write "            x: 0.5,"
response.write "            y: -0.2,"
response.write "            xanchor: 'center',"
response.write "            font: { color: textColor }"
response.write "        },"
response.write "        height: 600,"
response.write "        width: window.innerWidth * 0.98"
response.write "    };"

response.write "    Plotly.newPlot('weeklyVisitsChartDiv', traces, layout);"
response.write "});"
response.write "</script>"

   
   
response.write "<script>"
response.write "document.addEventListener('DOMContentLoaded', function() {"

response.write "    var ageGroup = " & jsonWeeklyData & ";"

response.write "    new DataTable('#weeklyTable', {"
'response.write "        data: ageGroup.dataAgeGroup.data,"
response.write "        data: dbDataWeekly.data,"
response.write "        columns: ["
 response.write "            { data: 'counter',title:'S/No.' },"
    response.write "            { data: 'DispenseYear' ,title:'Year'},"
    response.write "            { data: 'DispenseWeek',title:'Week' },"
'    response.write "            { data: 'DrugName',title:'Drug' },"
    
    response.write "            {"
response.write "                data: 'DrugName',"
response.write "                title: 'Drug',"
response.write "                render: function(data, type, row) {"
response.write "                    if (type === 'display') {"
response.write "                        var tooltip = row.PrescriptionNarrative ? row.PrescriptionNarrative.replace(/\""/g, '&quot;') : '';"
response.write "                        var displayValue = data;"
response.write "                        return '<span class=\""custom-tooltip\"" data-tooltip=\""' + tooltip + '\"">' + displayValue + '</span>';"
response.write "                    }"
response.write "                    return data;"
response.write "                }"
response.write "            },"
    response.write "            { data: 'QuarterQty', title:'Qty Dispensed', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
   response.write "            { data: 'PrevQuarterQty', title:'Prev.  Week\'s Qty', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
               
   response.write "            { data: 'AbsoluteChange', title:'Difference', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"

response.write "            { data: 'PercentageChange', title:'% Change', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
    response.write "            { data: 'CumulativeQty', title:'Cumulative Qty', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
    response.write "            { data: 'PrescriptionNarrative',title:'Prescription Narrative' ,visible:false},"
response.write "        ],"
response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
response.write "        dom: 'lBfrtip',"
response.write "        buttons: ["
response.write "            {"
response.write "                extend: 'csv',"
response.write "                text: 'CSV',"
response.write "                title: 'Weekly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
response.write "            },"
response.write "            {"
response.write "                extend: 'excel',"
response.write "                text: 'EXCEL',"
response.write "                title: 'Weekly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
response.write "            },"
response.write "            {"
response.write "                extend: 'pdf',"
response.write "                text: 'PDF',"
response.write "                title: 'Weekly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
response.write "            }"
response.write "        ]"
response.write "    });"

response.write "    hideSpinner();"
response.write "});"

' Function to hide spinner
response.write "function hideSpinner() {"
response.write "    document.getElementById('spinner').style.display = 'none';"
response.write "}"
response.write "</script>"


End Sub

Sub filters()
    response.write "    <main class=""container-fluid"">"
    response.write "        <div class=""content"">"
    response.write "            <!-- start: page body area -->"
    response.write "            <!-- <div class=""px-xl-5 px-lg-4 px-3 py-3 page-body""> -->"
    response.write "            <div class="" px-0 py-3 page-body"">"
    response.write "                <div class=""card mb-1"">"
    response.write "                    <div class=""card-body"">"
    response.write "                        <div class=""row g-3"">"
    SetBranch
    SetNumberDropdown
    ' SetAgeGroup
    response.write "                            <div class=""col-md-6 col-lg-3 col-xl-3"">"
    response.write "                                <div class=""form-group"">"
    response.write "                                    <div class=""form-group mb-3"">"
    response.write "                                        <small class=""form-label text-muted"">From:</small>"
    response.write "                                        <input type=""date"" class=""form-control"" id ='startDate'>"
    response.write "                                    </div>"
    response.write "                                </div>"
    response.write "                            </div>"
    response.write "                            <div class=""col-md-6 col-lg-3 col-xl-3"">"
    response.write "                                <div class=""form-group"">"
    response.write "                                    <div class=""form-group mb-3"">"
    response.write "                                        <small class=""form-label text-muted"">To:</small>"
    response.write "                                        <input type=""date"" class=""form-control"" id ='endDate'>"
    response.write "                                    </div>"
    response.write "                                </div>"
    response.write "                            </div>"
    response.write "                            <div class=""col-md-6 col-lg-3 col-xl-3 d-flex align-items-center justify-content-center "">"
    response.write "                                <button type=""submit "" class=""btn btn-primary"" onclick=""PeriodOnclick()"">Apply</button>"
    response.write "                            </div>"
    response.write "                        </div>"
    response.write "                    </div>"
    response.write "                </div>"
    response.write "            </div>"
    response.write ""
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

    htStr = htStr & "                    function getQueryParam(param) { " & vbCrLf
    htStr = htStr & "                        let urlParams = new URLSearchParams(window.location.search); " & vbCrLf
    htStr = htStr & "                        return urlParams.get(param); " & vbCrLf
    htStr = htStr & "                    } " & vbCrLf

    htStr = htStr & "          function formatDate(dateString) { " & vbCrLf
    htStr = htStr & "                   dateString = String(dateString).trim(); " & vbCrLf
    htStr = htStr & "                  var date = new Date(dateString); " & vbCrLf
    htStr = htStr & "                   if (isNaN(date)) { " & vbCrLf
    htStr = htStr & "                       console.error('Invalid date string'); " & vbCrLf
    htStr = htStr & "                      return null; " & vbCrLf
    htStr = htStr & "                  } " & vbCrLf
    htStr = htStr & "                 var year = date.getFullYear(); " & vbCrLf
    htStr = htStr & "                var month = String(date.getMonth() + 1).padStart(2, '0'); " & vbCrLf
    htStr = htStr & "                var day = String(date.getDate()).padStart(2, '0'); " & vbCrLf

    htStr = htStr & "                 return `${year}-${month}-${day}`; " & vbCrLf
    htStr = htStr & "             } " & vbCrLf

    htStr = htStr & "  document.addEventListener('DOMContentLoaded', (event) => { " & vbCrLf

    htStr = htStr & "   var dropdown = document.getElementById('Branchs'); " & vbCrLf
    htStr = htStr & "                 var defaultBranch = getQueryParam('branID'); " & vbCrLf

    htStr = htStr & "   var dropdownNumberSelect = document.getElementById('NumberSelect'); " & vbCrLf
    htStr = htStr & "                 var defaultNumberSelect = getQueryParam('dropdownvalue'); " & vbCrLf

    htStr = htStr & "             for (var i = 0; i < dropdown.options.length; i++) { " & vbCrLf
    htStr = htStr & "               if (dropdown.options[i].value === defaultBranch) { " & vbCrLf
    htStr = htStr & "                   dropdown.selectedIndex = i; " & vbCrLf
    htStr = htStr & "                 break; " & vbCrLf
    htStr = htStr & "             } " & vbCrLf
    htStr = htStr & "         } " & vbCrLf

    htStr = htStr & "             for (var i = 0; i < dropdownNumberSelect.options.length; i++) { " & vbCrLf
    htStr = htStr & "               if (dropdownNumberSelect.options[i].value === defaultNumberSelect) { " & vbCrLf
    htStr = htStr & "                   dropdownNumberSelect.selectedIndex = i; " & vbCrLf
    htStr = htStr & "                 break; " & vbCrLf
    htStr = htStr & "             } " & vbCrLf
    htStr = htStr & "         } " & vbCrLf

    htStr = htStr & "                 var selectedValue = getQueryParam('selectedValue'); " & vbCrLf
    htStr = htStr & "                     if (selectedValue === null) {   " & vbCrLf
    htStr = htStr & "          var startDateInput = document.getElementById('startDate'); " & vbCrLf
    htStr = htStr & "          var today = new Date(); " & vbCrLf
    htStr = htStr & "          var day = ('0' + today.getDate()).slice(-2); " & vbCrLf
    htStr = htStr & "          var month = ('0' + (today.getMonth() + 1)).slice(-2); " & vbCrLf
    htStr = htStr & "          var todayString = today.getFullYear() + '-' + month + '-' + day; " & vbCrLf
    htStr = htStr & "          startDateInput.value = todayString; " & vbCrLf
    htStr = htStr & "                     } else { " & vbCrLf
    htStr = htStr & "                         var formattedDate = formatDate(selectedValue); " & vbCrLf
    htStr = htStr & "                         if (formattedDate) { " & vbCrLf
    htStr = htStr & "                         console.log('Formatted Date: ' + formattedDate); " & vbCrLf
    htStr = htStr & "                          } else { " & vbCrLf
    htStr = htStr & "                                    console.error('Invalid date string'); " & vbCrLf
    htStr = htStr & "                        } " & vbCrLf

    htStr = htStr & "          var startDateInput = document.getElementById('startDate'); " & vbCrLf
    htStr = htStr & "          startDateInput.value = formattedDate; " & vbCrLf
    htStr = htStr & "                     } " & vbCrLf
    htStr = htStr & "                 var selectedValue1 = getQueryParam('selectedValue1'); " & vbCrLf
    htStr = htStr & "                     if (selectedValue1 === null) { " & vbCrLf
    htStr = htStr & "          var endDateInput = document.getElementById('endDate'); " & vbCrLf
    htStr = htStr & "          var today = new Date(); " & vbCrLf
    htStr = htStr & "          var day = ('0' + today.getDate()).slice(-2); " & vbCrLf
    htStr = htStr & "          var month = ('0' + (today.getMonth() + 1)).slice(-2); " & vbCrLf
    htStr = htStr & "          var todayString = today.getFullYear() + '-' + month + '-' + day; " & vbCrLf
    htStr = htStr & "          endDateInput.value = todayString; " & vbCrLf
    htStr = htStr & "                     } else { " & vbCrLf
    htStr = htStr & "                         var formattedDate = formatDate(selectedValue1); " & vbCrLf
    htStr = htStr & "                         if (formattedDate) { " & vbCrLf
    htStr = htStr & "                         console.log('Formatted Date: ' + formattedDate); " & vbCrLf
    htStr = htStr & "                          } else { " & vbCrLf
    htStr = htStr & "                                    console.error('Invalid date string'); " & vbCrLf
    htStr = htStr & "                        } " & vbCrLf

    htStr = htStr & "          var endDateInput = document.getElementById('endDate'); " & vbCrLf
    htStr = htStr & "          endDateInput.value = formattedDate; " & vbCrLf
    htStr = htStr & "                     } " & vbCrLf

    htStr = htStr & "      }); " & vbCrLf

    htStr = htStr & "var branchID1" & vbCrLf
    htStr = htStr & "function MonthOnchange(){" & vbCrLf
    htStr = htStr & "var ur, mth, sp, ordByTyp, ms, emr, str;" & vbCrLf
    htStr = htStr & "mth =  document.getElementById('NoOfDays').value;" & vbCrLf
    htStr = htStr & "dayid = 0;" & vbCrLf
    htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MostDespensedDrugsBA&PositionForTableName=WorkingDay';" & vbCrLf
    htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&month=' + mth  + ' &yearid=' + dayid + '&NoGlobal=1';" & vbCrLf
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
    htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MostDespensedDrugsBA&PositionForTableName=WorkingDay';" & vbCrLf
    htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&month=' + mth  + ' &yearid=' + dayid + '&NoGlobal=1' ;" & vbCrLf
    htStr = htStr & "window.location.href = processurl(ur);" & vbCrLf
    htStr = htStr & "}" & vbCrLf
    'emrdata
    htStr = htStr & "function PeriodOnclick(){ " & vbCrLf
    htStr = htStr & "var branchID1 =  document.getElementById('Branchs').value;" & vbCrLf
    htStr = htStr & "var dropdownNumberSelect1 =  document.getElementById('NumberSelect').value;" & vbCrLf
    htStr = htStr & "var startDate1 =  document.getElementById('startDate').value;" & vbCrLf
    htStr = htStr & "var endDate1 =  document.getElementById('endDate').value;" & vbCrLf
    htStr = htStr & "startDate1 = startDate1.trimEnd();" & vbCrLf
    htStr = htStr & "endDate1 = endDate1.trimEnd();" & vbCrLf
    htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MostDespensedDrugsBA&PositionForTableName=WorkingDay';" & vbCrLf
    ' htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&selectedValue=' + startDate1  + ' &selectedValue1=' + endDate1 + ' &branID=' + branchID1 + '&NoGlobal=1';" & vbCrLf
    htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&selectedValue=' + startDate1 + '&selectedValue1=' + endDate1 + '&branID=' + branchID1 + '&dropdownvalue=' + dropdownNumberSelect1 + '&NoGlobal=1';" & vbCrLf

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
    htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MostDespensedDrugsBA&PositionForTableName=WorkingDay';" & vbCrLf
    htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&selectedValue=' + startDate1  + ' &selectedValue1=' + endDate1 + ' &branID=' + branchID1 + '&NoGlobal=1';" & vbCrLf
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

    ' htStr = htStr & "</script>"
    response.write htStr
    js = js & "<script>" & vbCrLf
    js = js & "  " & vbCrLf
    js = js & "  " & vbCrLf
    js = js & "</script>"
    response.write js
End Sub

' Sub SetBranch()
'     Set rst = CreateObject("ADODB.Recordset")
'     dyHt = "<div class='col-md-6 col-lg-3 col-xl-3'>"
'     dyHt = dyHt & "<div class='form-group'>"
'     dyHt = dyHt & "<small class='form-label text-muted'>Facility:</small>"
'     dyHt = dyHt & "<select class='form-select' name='Branchs' id='Branchs' onchange='BranchOnchange()'>"
    
'     ' Default option
'     dyHt = dyHt & "<option value=''></option>"

'     sql0 = "select BranchID, BranchName from Branch b"
    
'     ' Open the recordset and loop through the results
'     With rst
'         .Open qryPro.FltQry(sql0), conn, 3, 4
'         If .RecordCount > 0 Then
'             .MoveFirst
'             Do While Not .EOF
'                 branchIDD = Trim(.Fields("BranchID"))
'                 branchName11 = Trim(.Fields("BranchName"))

'                 ' Check if this branch is selected
'                 If UCase(CStr(yearid)) = UCase(branchIDD) Then
'                     dyHt = dyHt & "<option value='" & CStr(branchIDD) & "' selected>" & branchName11 & "</option>"
'                 Else
'                     dyHt = dyHt & "<option value='" & CStr(branchIDD) & "'>" & branchName11 & "</option>"
'                 End If

'                 rst.MoveNext
'                 Response.flush
'             Loop
'         End If
'         .Close
'     End With
    
'     ' Close the select and divs
'     dyHt = dyHt & "</select>"
'     dyHt = dyHt & "</div>"
'     dyHt = dyHt & "</div>"
    
'     ' Output the final HTML
'     Response.Write dyHt

'     Set rst = Nothing
' End Sub

' Sub SetBranch()
'     Dim rst, dyHt, sql0, branchIDD, branchName11, isFirst

'     Set rst = CreateObject("ADODB.Recordset")
'     isFirst = True

'     dyHt = "<div class='col-md-6 col-lg-3 col-xl-3'>"
'     dyHt = dyHt & "<div class='form-group'>"
'     dyHt = dyHt & "<small class='form-label text-muted'>Facility:</small>"
'     dyHt = dyHt & "<select class='form-select' name='Branchs' id='Branchs' onchange='BranchOnchange()'>"

'     ' SQL query
'     sql0 = "SELECT BranchID, BranchName FROM Branch b"

'     ' Open the recordset and loop through the results
'     With rst
'         .Open qryPro.FltQry(sql0), conn, 3, 4
'         If .RecordCount > 0 Then
'             .MoveFirst
'             Do While Not .EOF
'                 branchIDD = Trim(.Fields("BranchID"))
'                 branchName11 = Trim(.Fields("BranchName"))

'                 ' Select the first record by default
'                 If isFirst Then
'                     dyHt = dyHt & "<option value='" & CStr(branchIDD) & "' selected>" & branchName11 & "</option>"
'                     isFirst = False
'                 Else
'                     dyHt = dyHt & "<option value='" & CStr(branchIDD) & "'>" & branchName11 & "</option>"
'                 End If

'                 .MoveNext
'                 Response.Flush
'             Loop
'         End If
'         .Close
'     End With

'     ' Close the select and divs
'     dyHt = dyHt & "</select>"
'     dyHt = dyHt & "</div>"
'     dyHt = dyHt & "</div>"

'     ' Output the final HTML
'     Response.Write dyHt

'     Set rst = Nothing
' End Sub


Sub SetBranch()
    Dim rst, dyHt, sql0, branchIDD, branchName11, rowIndex

    Set rst = CreateObject("ADODB.Recordset")
    rowIndex = 0

    dyHt = "<div class='col-md-6 col-lg-3 col-xl-3'>"
    dyHt = dyHt & "<div class='form-group'>"
    dyHt = dyHt & "<small class='form-label text-muted'>Facility:</small>"
    dyHt = dyHt & "<select class='form-select' name='Branchs' id='Branchs' onchange='BranchOnchange()'>"

    ' SQL query
    sql0 = "SELECT BranchID, BranchName FROM Branch b"

    ' Open the recordset and loop through the results
    With rst
        .open qryPro.FltQry(sql0), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                rowIndex = rowIndex + 1
                branchIDD = Trim(.Fields("BranchID"))
                branchName11 = Trim(.Fields("BranchName"))

                ' ? Select the second record by default
                If rowIndex = 2 Then
                    dyHt = dyHt & "<option value='" & CStr(branchIDD) & "' selected>" & branchName11 & "</option>"
                Else
                    dyHt = dyHt & "<option value='" & CStr(branchIDD) & "'>" & branchName11 & "</option>"
                End If

                .MoveNext
                response.flush
            Loop
        End If
        .Close
    End With

    ' Close the select and divs
    dyHt = dyHt & "</select>"
    dyHt = dyHt & "</div>"
    dyHt = dyHt & "</div>"

    ' Output the final HTML
    response.write dyHt

    Set rst = Nothing
End Sub


Sub SetNumberDropdown()
    Dim dyHt, numbers, num, defaultNum

    ' Define the list of numbers
'    numbers = Array(10, 15, 20, 25, 30)
    numbers = Array(10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65)
    defaultNum = 10 ' Default selection

    dyHt = "<div class='col-md-6 col-lg-3 col-xl-3'>"
    dyHt = dyHt & "<div class='form-group'>"
    dyHt = dyHt & "<small class='form-label text-muted'>Select Top Number:</small>"
    dyHt = dyHt & "<select class='form-select' name='NumberSelect' id='NumberSelect' onchange='NumberOnchange()'>"

    ' Loop through numbers
    For Each num In numbers
        If num = defaultNum Then
            dyHt = dyHt & "<option value='" & num & "' selected>" & num & "</option>"
        Else
            dyHt = dyHt & "<option value='" & num & "'>" & num & "</option>"
        End If
    Next

    dyHt = dyHt & "</select>"
    dyHt = dyHt & "</div>"
    dyHt = dyHt & "</div>"

    response.write dyHt
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

    ' Array of month names
    Dim monthNames
    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    Dim monthName
    monthName = monthNames(CInt(monthPart) - 1) ' Subtract 1 for zero-based index

    formatedDate = dayPart & "-" & monthName & "-" & yearPart
    FormatDateNew = formatedDate
End Function

Sub ShowLoading()
    response.write "<script>"
    response.write "        var spinner = document.getElementById('spinner');"
    response.write "            spinner.style.display = 'flex';"  ' Show the spinner
    response.write "</script>"

End Sub

Sub HideLoading()
    response.write "<script>"
    response.write "        var spinner = document.getElementById('spinner');"
    response.write "            spinner.style.display = 'none';"  ' Hide the spinner
    response.write "</script>"

End Sub



'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
