VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRTLYO_PrescriptionErrorRatesBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
    Response.Clear

    conn.CommandTimeout = 7200
    Dim periodStart, periodEnd, brnchID, jsonData
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

    Response.Write "<!DOCTYPE html>"
    Response.Write "<html lang=""en"">"
    Response.Write ""
    Response.Write "<head>"
    Response.Write "    <title>Prescription Error Rates </title>"
    Response.Write "    <meta charset=""utf-8"">"
    Response.Write "    <meta http-equiv=""X-UA-Compatible"" content=""IE=Edge"">"
    Response.Write "    <meta name=""viewport"" content=""width=device-width, initial-scale=1, shrink-to-fit=no"">"
    Response.Write "    <!--[ Favicon]-->"
    Response.Write "    <link rel=""icon"" type=""image/x-icon"" href=""images/banner1.bmp"">"
    Response.Write "    <link rel=""icon"" type=""image/png"" sizes=""16x16"" href=""images/banner1.bmp"">"
    Response.Write "    <link rel=""icon"" type=""image/png"" sizes=""32x32"" href=""images/banner1.bmp"">"
    Response.Write "    <link rel=""apple-touch-icon"" sizes=""180x180"" href=""images/banner1.bmp"">"
    Response.Write ""
    Response.Write "    <!--[ Template main css file ]-->"
    Response.Write "    <link rel=""stylesheet"" href=""BusinessAnalytics/assets/css/style.min.css"">"

   
    Response.Write "<script src='BusinessAnalytics/assets/bundles/plotly.min.js'></script>"

    Response.Write "<script src='BusinessAnalytics/test3styles/styles/jsdelivrv530.js'></script>"
    Response.Write "<script src='BusinessAnalytics/test3styles/styles/jsdelivrselect2v410.js'></script>"
    Response.Write "<script src='BusinessAnalytics/test3styles/styles/jsdelivrjqueryv360.js'></script>"
    Response.Write "<script src='BusinessAnalytics/test3styles/styles/cloudflarev027.js'></script>"
    Response.Write "<script src='BusinessAnalytics/test3styles/styles/cdnjscloudflarev027.js'></script>"
    Response.Write "<script src='BusinessAnalytics/test3styles/styles/cdndatatablesv111.min.js'></script>"

    Response.Write "<script src='BusinessAnalytics/test3styles/styles/multi-select-tag.js'></script>"
    Response.Write "<script src='BusinessAnalytics/test3styles/styles/sweetalert2.all.min.js'></script>"

    Response.Write "    <link rel=""stylesheet"" href=""BusinessAnalytics/test3styles/styles/sweetalert2.min.css"">"
    Response.Write "    <link rel=""stylesheet"" href=""BusinessAnalytics/test3styles/styles/habibmhamadimultiselecttag.css"">"
    Response.Write "    <link rel=""stylesheet"" href=""BusinessAnalytics/test3styles/styles/tdatatables.min.css"">"
    Response.Write "    <link rel=""stylesheet"" href=""BusinessAnalytics/test3styles/styles/jsdelivrv530.css"">"
    Response.Write "    <link rel=""stylesheet"" href=""BusinessAnalytics/test3styles/styles/jsdelivrv502.css"">"
    Response.Write "    <link rel=""stylesheet"" href=""BusinessAnalytics/test3styles/styles/jsdelivrselect2v410.css"">"
     Response.Write "    <link rel=""stylesheet"" href=""BusinessAnalytics/assets/css/style.min.css"">"
   
    
        
     ' style for tooltip
Response.Write "<style>"
Response.Write ".custom-tooltip { position: relative; cursor: help; }"
Response.Write ".custom-tooltip::after {"
Response.Write "    content: attr(data-tooltip);"
Response.Write "    position: absolute;"
Response.Write "    bottom: 100%;"
Response.Write "    left: 50%;"
Response.Write "    transform: translateX(-50%);"
'Response.Write "    background-color: aqua;"
Response.Write "    background-color: #003366;"
Response.Write "    color: white;"
Response.Write "    padding: 8px 12px;"
Response.Write "    font-size: 14px;"
Response.Write "    border-radius: 4px;"
Response.Write "    white-space: pre-wrap;"
Response.Write "    z-index: 9999;"
Response.Write "    display: none;"
Response.Write "    min-width: 320px;"
Response.Write "    box-shadow: 0 4px 12px rgba(0,0,0,0.25);"
Response.Write "    line-height: 1.4;"
Response.Write "}"
Response.Write ".custom-tooltip:hover::after {"
Response.Write "    display: block;"
Response.Write "}"
Response.Write "</style>"
    
    
    Response.Write "</head>"
    Response.Write "<body data-theme=""theme-PurpleHeart"" class=""svgstroke-a bg-gradient"">"
    ' loading----------------------------------------------------------------------------------------------
    Response.Write "<style>"
    Response.Write "    #spinner {"
    Response.Write "        position: fixed;"
    Response.Write "        top: 0;"
    Response.Write "        left: 0;"
    Response.Write "        width: 100%;"
    Response.Write "        height: 100%;"
    ' Response.Write "        background: rgba(255, 255, 255, 0.7);"
    Response.Write "        display: none;" ' Initially hidden
    Response.Write "        align-items: center;"
    Response.Write "        justify-content: center;"
    Response.Write "        backdrop-filter: blur(4px);"
    Response.Write "        z-index: 9999;"
    Response.Write "    }"
    Response.Write "    .spinner-border {"
    Response.Write "        width: 3rem;"
    Response.Write "        height: 3rem;"
    Response.Write "    }"
    Response.Write "    .spinner-grow {"
    Response.Write "        width: 3rem;"
    Response.Write "        height: 3rem;"
    Response.Write "        margin-left: 1rem;"
    Response.Write "    }"
    Response.Write "</style>"

    Response.Write "<div id='spinner'>"
    Response.Write "    <div class=""spinner-border"" role=""status"">"
    Response.Write "        <span class=""visually-hidden"">Loading...</span>"
    Response.Write "    </div>"
    Response.Write "    <div class=""spinner-grow"" role=""status"">"
    Response.Write "        <span class=""visually-hidden"">Loading...</span>"
    Response.Write "    </div>"
    Response.Write "</div>"
    ShowLoading
    ' ----------------------------------------------------------------------------------------------
    InitPageScript
    ' ----------------------------------------------------------------------------------------------
    Response.Write ""
    Response.Write "    <!-- <main class=""container-fluid px-0""> -->"
    ' ----------------------------------------------------------------------------------------------
    filters
    ' ----------------------------------------------------------------------------------------------
    Response.Write "            <div class=""card bg-transparent border-0"">"
    Response.Write "                <ul class=""nav nav-tabs li_animate"" role=""tablist"">"
    Response.Write "                    <li class=""nav-item"" role=""presentation""><a class=""nav-link active"" data-bs-toggle=""tab"" href=""#Yearly"" aria-selected=""true"" role=""tab""><span>Annual Prescription Error Trends</span></a></li>"
    Response.Write "                    <li class=""nav-item"" role=""presentation""><a class=""nav-link"" data-bs-toggle=""tab"" href=""#Quarterly"" aria-selected=""false"" tabindex=""-1"" role=""tab""><span class=""d-none d-sm-inline-flex ps-1"">Quarterly Prescription Error Trends</span></a></li>"
'    Response.Write "                    <li class=""nav-item"" role=""presentation""><a class=""nav-link"" data-bs-toggle=""tab"" href=""#Monthly"" aria-selected=""false"" tabindex=""-1"" role=""tab""><span class=""d-none d-sm-inline-flex ps-1"">Monthly Prescription Error Trends</span></a></li>"
    Response.Write "                    <li class=""nav-item"" role=""presentation""><a class=""nav-link"" data-bs-toggle=""tab"" href=""#Weekly"" aria-selected=""false"" tabindex=""-1"" role=""tab""><span class=""d-none d-sm-inline-flex ps-1"">Prescription Error Details</span></a></li>"
    Response.Write "                </ul>"
    Response.Write "                <div class=""tab-content bg-card p-3 border border-top-0 rounded-bottom"">"
    Response.Write "                    <!--[ Start:: Yearly]-->"
    Response.Write "                    <div class=""tab-pane fade show active"" id=""Yearly"" role=""tabpanel"">"
    Response.Write "                        <ul class=""row g-3 list-unstyled li_animate mb-0 "">"
    Response.Write "                            <div class=""col-12 mb-5 "">"
    
    get_annual_prescription_appropriateness

    Response.Write "                            </div>"
    Response.Write "                        </ul>"
    Response.Write "                    </div>"
    Response.Write "                    <!--[ Start:: Quarterly]-->"
    Response.Write "                    <div class=""tab-pane fade "" id=""Quarterly"" role=""tabpanel"">"
    Response.Write "                        <ul class=""row g-3 list-unstyled li_animate mb-0 "">"
    Response.Write "                            <div class=""col-12 mb-5 "">"
    
    get_quarterlyl_prescription_appropriateness

    Response.Write "                            </div>"
    Response.Write "                        </ul>"
    Response.Write "                    </div>"
    Response.Write "                    <!--[ Start:: Monthly]-->"
    Response.Write "                    <div class=""tab-pane fade "" id=""Monthly"" role=""tabpanel"">"
    Response.Write "                        <ul class=""row g-3 list-unstyled li_animate mb-0 "">"
    Response.Write "                            <div class=""col-12 mb-5 "">"

'    monthly_most_dispensed_medications

    Response.Write "                            </div>"
    Response.Write "                        </ul>"
    Response.Write "                    </div>"
    Response.Write "                    <!--[ Start:: Weekly]-->"
    Response.Write "                    <div class=""tab-pane fade "" id=""Weekly"" role=""tabpanel"">"
    Response.Write "                        <ul class=""row g-3 list-unstyled li_animate mb-0 "">"
    Response.Write "                            <div class=""col-12 mb-5 "">"

    detailed_prescription_errors

    Response.Write "                            </div>"
    Response.Write "                        </ul>"
    Response.Write "                    </div>"
    Response.Write "                </div>"
    Response.Write "            </div>"
    Response.Write ""
    Response.Write "        </div>"
    Response.Write ""
    Response.Write "        </div>"
    Response.Write "        </div>"
    Response.Write ""
    Response.Write "    </main>"
    Response.Write "    <script src=""BusinessAnalytics/assets/bundles/libscripts.bundle.js ""></script>"
    Response.Write "    <script src=""BusinessAnalytics/assets/js/main.js ""></script>"
    Response.Write ""
    Response.Write "</body>"
    Response.Write ""
    Response.Write "</html>"

Sub get_annual_prescription_appropriateness()

    Response.Write "  <div class='chart-container'>"
    Response.Write "    <div id='yearlyChartDivGender' class='chart table-responsive'>"
    Response.Write "  </div>"
    Response.Write "   <div class=""card"">"
    Response.Write "   <div class=""card-body"">"
    Response.Write "       <div class='table-responsive'>"
    Response.Write "      <table class=""table dataTable align-middle table-hover table-bordered"" style=""width: 100%;"" id=""interVisitTable"" >"
    Response.Write "      <thead >"
    Response.Write "              <tr>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
     Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "              </tr>"
    Response.Write "            </thead>"
    Response.Write "        </table>"
    Response.Write "       </div>"
    Response.Write "      </div>"
    Response.Write "  </div>"


Dim sql, rst, cmd, jsonAnnualData
Set rst = CreateObject("ADODB.RecordSet")
Set cmd = CreateObject("ADODB.Command")

sql = "exec dbo.sp_annual_appropriateness1 '" & periodStart & "','" & periodEnd & "'"

cmd.ActiveConnection = conn
cmd.CommandText = sql
cmd.CommandType = 1

Set rst = cmd.Execute

' Handle possible no data or NULL field
If Not rst.EOF And Not IsNull(rst.Fields(0).value) Then
    ' Ensure it's valid JSON (basic check)
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


'response.write "<script>"
'response.write "var dbDataYearlyGender = " & jsonAnnualData & ";"
'response.write "document.addEventListener('DOMContentLoaded', function() {"
'response.write "    var data = dbDataYearlyGender.data;"
'
'' Unique years and drugs
'response.write "    var years = [...new Set(data.map(d => d.VisitYear))];"
'response.write "    var drugNames = [...new Set(data.map(d => d.Category))];"
'
'response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
'
'response.write "    var traces = drugNames.map(function(drug, index) {"
'response.write "        var yValues = years.map(function(year) {"
'response.write "            var entry = data.find(function(d) { return d.Category === drug && d.VisitYear === year; });"
'response.write "            return entry ? parseFloat(entry.CurrentYearCount) : 0;"
'response.write "        });"
'
'response.write "        var hoverText = years.map(function(year) {"
'response.write "            var entry = data.find(function(d) { return d.Category === drug && d.VisitYear === year; });"
'response.write "            return entry ? "
'response.write "                'Year: ' + entry.VisitYear + '<br>' +"
'response.write "                'Total Dispense Count: ' + (entry.DispenseCount ? parseFloat(entry.DispenseCount).toLocaleString() : 'N/A') + '<br>' +"
'response.write "                'Appropriateness: ' + (entry.Appropriateness ? entry.Appropriateness : 'N/A') + '<br>' +"
'
'response.write "                'Error Rate Percent: ' + (entry.ErrorRatePercent ? parseFloat(entry.ErrorRatePercent).toLocaleString() : 'N/A') + '<br>' +"
'
'
'response.write "                'Appropriateness Of: ' + entry.Category + '<br>' +"
'response.write "                'Current Year Error Counts: ' + parseFloat(entry.CurrentYearCount).toLocaleString() + '<br>' +"
'response.write "                'Prev. Year\'s Error Counts: ' + (entry.PreviousYearCount ? parseFloat(entry.PreviousYearCount).toLocaleString() : 'N/A') + '<br>' +"
'
'response.write "                'Difference: ' + (entry.AbsoluteChange ? parseFloat(entry.AbsoluteChange).toLocaleString() : 'N/A') + '<br>' +"
'
'response.write "                '% Change: ' + (entry.YoY_PercentChange ? parseFloat(entry.YoY_PercentChange).toLocaleString() + '%' : 'N/A') + '<br>' +"
'
'response.write "                'Narrative: ' + entry.QI_Narrative : 'No data';"
'response.write "        });"
'
'response.write "        return {"
'response.write "            x: years,"
'response.write "            y: yValues,"
'response.write "            name: drug,"
'response.write "            type: 'bar',"
'response.write "            marker: { color: colors[index % colors.length] },"
'response.write "            text: yValues.map(v => v.toLocaleString()),"
'response.write "            textposition: 'auto',"
'response.write "            hovertemplate: hoverText"
'response.write "        };"
'response.write "    });"
'
'' Theme support
'response.write "    function getCurrentTheme() {"
'response.write "        return localStorage.getItem('theme') || 'light';"
'response.write "    }"
'response.write "    var currentTheme = getCurrentTheme();"
'response.write "    var textColor = currentTheme === 'dark' ? '#ffffff' : '#000000';"
'response.write "    var chartBackgroundColor = currentTheme === 'dark' ? '#333' : '#f0f0f0';"
'response.write "    var gridColorTheme = currentTheme === 'dark' ? '#444' : '#ccc';"
'
'' Layout
'response.write "    var layout = {"
'response.write "        title: {"
'response.write "            text: 'Yearly Dispensing Trends Between " & FormatDateNew(periodStart) & " and " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "',"
'response.write "            font: { color: textColor }"
'response.write "        },"
'response.write "        font: { color: textColor },"
'response.write "        xaxis: {"
'response.write "            title: 'Year',"
'response.write "            type: 'category',"
'response.write "            titlefont: { color: textColor },"
'response.write "            tickfont: { color: textColor },"
'response.write "            gridcolor: gridColorTheme"
'response.write "        },"
'response.write "        yaxis: {"
'response.write "            title: 'Quantity Dispensed',"
'response.write "            titlefont: { color: textColor },"
'response.write "            tickfont: { color: textColor },"
'response.write "            gridcolor: gridColorTheme"
'response.write "        },"
'response.write "        barmode: 'group',"
'response.write "        paper_bgcolor: chartBackgroundColor,"
'response.write "        plot_bgcolor: chartBackgroundColor,"
'response.write "        height: 600,"
'response.write "        width: window.innerWidth * 0.95,"
'response.write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
'response.write "        legend: {"
'response.write "            orientation: 'h',"
'response.write "            x: 0.5,"
'response.write "            xanchor: 'center',"
'response.write "            y: -0.2,"
'response.write "            yanchor: 'top',"
'response.write "            font: { color: textColor }"
'response.write "        }"
'response.write "    };"
'
'response.write "    Plotly.newPlot('yearlyChartDivGender', traces, layout);"
'response.write "});"
'response.write "</script>"

Response.Write "<script>"
Response.Write "var dbDataYearlyGender = " & jsonAnnualData & ";"
Response.Write "document.addEventListener('DOMContentLoaded', function() {"
Response.Write "    var data = dbDataYearlyGender.data;"

' Unique years, categories, and appropriateness labels
Response.Write "    var years = [...new Set(data.map(d => d.VisitYear))];"
Response.Write "    var groups = [...new Set(data.map(d => d.Category + ' - ' + (d.Appropriateness || 'Unknown')))];"

Response.Write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"

' Build stacked bar traces
Response.Write "    var traces = groups.map(function(groupLabel, index) {"
Response.Write "        var split = groupLabel.split(' - ');"
Response.Write "        var category = split[0];"
Response.Write "        var appropriateness = split[1];"

Response.Write "        var yValues = years.map(function(year) {"
Response.Write "            var entry = data.find(function(d) {"
Response.Write "                return d.Category === category && (d.Appropriateness || 'Unknown') === appropriateness && d.VisitYear === year;"
Response.Write "            });"
Response.Write "            return entry ? parseFloat(entry.CurrentYearCount || 0) : 0;"
Response.Write "        });"

Response.Write "        var hoverText = years.map(function(year) {"
Response.Write "            var entry = data.find(function(d) {"
Response.Write "                return d.Category === category && (d.Appropriateness || 'Unknown') === appropriateness && d.VisitYear === year;"
Response.Write "            });"
Response.Write "            return entry ? "
Response.Write "                'Year: ' + entry.VisitYear + '<br>' +"
Response.Write "                'Appropriateness Of: ' + entry.Category + '<br>' +"
Response.Write "                'Appropriateness: ' + (entry.Appropriateness || 'N/A') + '<br>' +"
Response.Write "                'Dispensed: ' + (entry.DispenseCount ? parseFloat(entry.DispenseCount).toLocaleString() : 'N/A') + '<br>' +"
Response.Write "                'Current Year Errors: ' + parseFloat(entry.CurrentYearCount).toLocaleString() + '<br>' +"
Response.Write "                'Previous Year Errors: ' + (entry.PreviousYearCount ? parseFloat(entry.PreviousYearCount).toLocaleString() : 'N/A') + '<br>' +"
Response.Write "                'Change: ' + (entry.AbsoluteChange ? parseFloat(entry.AbsoluteChange).toLocaleString() : 'N/A') + '<br>' +"
Response.Write "                'Error Rate: ' + (entry.ErrorRatePercent ? parseFloat(entry.ErrorRatePercent).toFixed(3) + '%' : 'N/A') + '<br>' +"
Response.Write "                'YoY % Change: ' + (entry.YoY_PercentChange ? parseFloat(entry.YoY_PercentChange).toFixed(1) + '%' : 'N/A') + '<br>' +"
Response.Write "                'Narrative: ' + (entry.QI_Narrative || 'N/A') : 'No data';"
Response.Write "        });"

Response.Write "        return {"
Response.Write "            x: years,"
Response.Write "            y: yValues,"
Response.Write "            name: groupLabel,"
Response.Write "            type: 'bar',"
Response.Write "            marker: { color: colors[index % colors.length] },"
Response.Write "            text: yValues.map(v => v.toLocaleString()),"
Response.Write "            textposition: 'auto',"
Response.Write "            hovertext: hoverText,"
Response.Write "            hoverinfo: 'text'"
Response.Write "        };"
Response.Write "    });"

' Theme and layout setup
Response.Write "    function getCurrentTheme() {"
Response.Write "        return localStorage.getItem('theme') || 'light';"
Response.Write "    }"
Response.Write "    var currentTheme = getCurrentTheme();"
Response.Write "    var textColor = currentTheme === 'dark' ? '#ffffff' : '#000000';"
Response.Write "    var chartBackgroundColor = currentTheme === 'dark' ? '#333' : '#f0f0f0';"
Response.Write "    var gridColorTheme = currentTheme === 'dark' ? '#444' : '#ccc';"

Response.Write "    var layout = {"
Response.Write "        title: {"
Response.Write "            text: 'Annual Prescription Error Trends <br>Between " & FormatDateNew(periodStart) & " and " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "',"
Response.Write "            font: { color: textColor }"
Response.Write "        },"
Response.Write "        font: { color: textColor },"
Response.Write "        xaxis: {"
Response.Write "            title: 'Year',"
Response.Write "            type: 'category',"
Response.Write "            titlefont: { color: textColor },"
Response.Write "            tickfont: { color: textColor },"
Response.Write "            gridcolor: gridColorTheme"
Response.Write "        },"
Response.Write "        yaxis: {"
Response.Write "            title: 'Error Counts',"
Response.Write "            titlefont: { color: textColor },"
Response.Write "            tickfont: { color: textColor },"
Response.Write "            gridcolor: gridColorTheme"
Response.Write "        },"
Response.Write "        barmode: 'stack',"
Response.Write "        paper_bgcolor: chartBackgroundColor,"
Response.Write "        plot_bgcolor: chartBackgroundColor,"
Response.Write "        height: 600,"
Response.Write "        width: window.innerWidth * 0.95,"
Response.Write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
Response.Write "        legend: {"
Response.Write "            orientation: 'h',"
Response.Write "            x: 0.5,"
Response.Write "            xanchor: 'center',"
Response.Write "            y: -0.2,"
Response.Write "            yanchor: 'top',"
Response.Write "            font: { color: textColor }"
Response.Write "        }"
Response.Write "    };"

Response.Write "    Plotly.newPlot('yearlyChartDivGender', traces, layout);"
Response.Write "});"


    '    // ? Chart click ? disease breakdown page
  Response.Write "      document.getElementById('yearlySamePeriodChartDiv').on('plotly_click', () => {"
  Response.Write "          window.location.href = 'disease-breakdown.html';"
  Response.Write "      });"



Response.Write "</script>"


'##################################################################################

    ' DataTable Initialization
    Response.Write "<script>"
    Response.Write "    new DataTable('#interVisitTable', {"
    Response.Write "        data: dbDataYearlyGender.data,"
    Response.Write "        columns: ["
    Response.Write "            { data: 'counter',title:'S/No.' },"
    Response.Write "            { data: 'DispenseYear',title:'Year' },"
    
     Response.Write "            { data: 'DispenseCount', title:'Total Dispense Count', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
               
               
               
    Response.Write "            { data: 'PrescriptionErrorCount', title:'Prescription Error Count', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
'     Response.Write "            { data: 'ErrorRatePercent', title:'Error Rate Percent', render: function(data) { " & _
'               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
         ' --- Narrative Column (with tooltip) ---
Response.Write "            {"
Response.Write "                data: 'ErrorRatePercent',"
Response.Write "                title: 'Error Rate Percent',"
Response.Write "                render: function(data, type, row) {"
Response.Write "                    if (type === 'display') {"
Response.Write "                        var tooltip = row.QI_Narrative ? row.QI_Narrative.replace(/""/g, '&quot;') : '';"
Response.Write "                        var val = parseFloat(data);"
Response.Write "                        var displayValue = isNaN(val) ? '0' : val.toLocaleString();"
Response.Write "                        return '<span class=""custom-tooltip"" data-tooltip=""" + "' + tooltip + '" + """>' + displayValue + '</span>'; "
Response.Write "                    }"
Response.Write "                    return data;"
Response.Write "                }"
Response.Write "            },"
    
    
    Response.Write "            { data: 'Category',title:'Appropriateness Of' },"

    
    
     Response.Write "            { data: 'Appropriateness',title:'Appropriateness' },"
    


               
   Response.Write "            { data: 'CurrentYearCount', title:'Current Year Count', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
   Response.Write "            { data: 'PreviousYearCount', title:'Prev. Year Count', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
   Response.Write "            { data: 'AbsoluteChange', title:'Difference', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
               
'    response.write "            { data: 'PercentageChange',title:'% Change' },"
Response.Write "            { data: 'YoY_PercentChange', title:'% Change', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
               

               
               
    
   
    Response.Write "            { data: 'QI_Narrative',title:'Narrative',visible:false },"
    Response.Write "        ],"
    Response.Write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    Response.Write "        dom: 'lBfrtip',"
    Response.Write "        buttons: ["
    Response.Write "            {"
    Response.Write "                extend: 'csv',"
    Response.Write "                text: 'CSV',"
    Response.Write "                title: 'Annual Prescription Error Trends Between : " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'excel',"
    Response.Write "                text: 'EXCEL',"
    Response.Write "                title: 'Annual Prescription Error Trends Between : " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'pdf',"
    Response.Write "                text: 'PDF',"
    Response.Write "                title: 'Annual Prescription Error Trends Between : " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
    Response.Write "            },"
'    response.write "            {"
'    response.write "                extend: 'print',"
'    response.write "                text: 'PRINT',"
'    response.write "                title: 'Yearly Dispensing Trends Between : " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
'    response.write "            }"
    Response.Write "        ]"
    Response.Write "    });"
    Response.Write "</script>"
    
End Sub

Sub get_quarterlyl_prescription_appropriateness()

    Response.Write "  <div class='chart-container'>"
    Response.Write "    <div id='quarterlyChartDiv' class='chart table-responsive'>"
    Response.Write "  </div>"
    Response.Write "   <div class=""card"">"
    Response.Write "   <div class=""card-body"">"
    Response.Write "       <div class='table-responsive'>"
    Response.Write "      <table class=""table dataTable align-middle table-hover table-bordered"" style=""width: 100%;"" id=""proportionTable"" >"
    Response.Write "      <thead >"
    Response.Write "              <tr>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "                <th></th>"
    Response.Write "               </tr>"
    Response.Write "            </thead>"
    Response.Write "        </table>"
    Response.Write "       </div>"
    Response.Write "      </div>"
    Response.Write "  </div>"


Dim sql, rst, cmd
    Set rst = CreateObject("ADODB.RecordSet")
    Set cmd = CreateObject("ADODB.Command")

    sql = "exec [dbo].[sp_quarterly_appropriateness] '" & periodStart & "','" & periodEnd & "'"

    ' Set up the Command object
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1
    Set rst = cmd.Execute

Dim jsonQuarterData

' Handle possible no data or NULL field
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

Response.Write "<script>"
Response.Write "var dbDataQuarterlyGender = " & jsonQuarterData & ";"
Response.Write "document.addEventListener('DOMContentLoaded', function() {"

Response.Write "    var revenueSourcesQuarterly = dbDataQuarterlyGender.data;"

' Color palette for drug traces
Response.Write "    var colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'];"

Response.Write "    var quarters = [...new Set(revenueSourcesQuarterly.map(d => d.DispenseYear + ' Q' + d.DispenseQuarter))];"

' Sort chronologically
Response.Write "    quarters.sort((a, b) => {"
Response.Write "        var [yearA, qA] = a.split(' Q').map(Number);"
Response.Write "        var [yearB, qB] = b.split(' Q').map(Number);"
Response.Write "        return yearA - yearB || qA - qB;"
Response.Write "    });"

' Unique drug names
Response.Write "    var drugNames = [...new Set(revenueSourcesQuarterly.map(d => d.Category))];"

' Build a trace per drug
Response.Write "    var traces = drugNames.map(function(drug, index) {"
Response.Write "        return {"
Response.Write "            name: drug,"
Response.Write "            type: 'bar',"
Response.Write "            marker: { color: colors[index % colors.length] },"
Response.Write "            x: quarters,"
Response.Write "            y: quarters.map(function(q) {"
Response.Write "                var item = revenueSourcesQuarterly.find(function(d) {"
Response.Write "                    return d.Category === drug && (d.DispenseYear + ' Q' + d.DispenseQuarter) === q;"
Response.Write "                });"
Response.Write "                return item ? parseFloat(item.CurrentQuarterCount) : 0;"
Response.Write "            }),"
Response.Write "            text: quarters.map(function(q) {"
Response.Write "                var item = revenueSourcesQuarterly.find(function(d) {"
Response.Write "                    return d.Category === drug && (d.DispenseYear + ' Q' + d.DispenseQuarter) === q;"
Response.Write "                });"
Response.Write "                return item ? parseFloat(item.PrescriptionErrorCount).toLocaleString() : '';"
Response.Write "            }),"
Response.Write "            hovertemplate: quarters.map(function(q) {"
Response.Write "                var item = revenueSourcesQuarterly.find(function(d) {"
Response.Write "                    return d.Category === drug && (d.DispenseYear + ' Q' + d.DispenseQuarter) === q;"
Response.Write "                });"
Response.Write "                return item ? "
Response.Write "                    'Year: ' + item.DispenseYear + '<br>' +"
Response.Write "                    'Quarter Q: ' + item.DispenseQuarter + '<br>' +"
Response.Write "                    'Appropriateness Of: ' + item.Category + '<br>' +" '
Response.Write "                    'Error Counts: ' + parseFloat(item.PrescriptionErrorCount).toLocaleString() + '<br>' +"
Response.Write "                    'Error Rate Percent: ' + parseFloat(item.ErrorRatePercent).toLocaleString() + '<br>' +"
Response.Write "                    'Prev. Quarter Counts: ' + (item.PreviousQuarterCount ? parseFloat(item.PreviousQuarterCount).toLocaleString() : 'N/A') + '<br>' +"
Response.Write "                    'Difference: ' + (item.AbsoluteChange ? parseFloat(item.AbsoluteChange).toLocaleString() : 'N/A') + '<br>' +"
Response.Write "                    '% Change: ' + (item.QoQ_PercentChange ? parseFloat(item.QoQ_PercentChange).toLocaleString() + '%' : 'N/A') + '<br>' +"
Response.Write "                    'Cumulative Qty: ' + parseFloat(item.CumulativeQty).toLocaleString()"
Response.Write "                    : 'No data';"
Response.Write "            }),"
Response.Write "            textposition: 'auto'"
Response.Write "        };"
Response.Write "    });"

' Theme support
Response.Write "    function getCurrentTheme() {"
Response.Write "        return localStorage.getItem('theme') || 'light';"
Response.Write "    }"
Response.Write "    var currentTheme = getCurrentTheme();"
Response.Write "    var textColor = currentTheme === 'dark' ? '#ffffff' : '#000000';"
Response.Write "    var chartBackgroundColor = currentTheme === 'dark' ? '#333' : '#f0f0f0';"
Response.Write "    var gridColorTheme = currentTheme === 'dark' ? '#444' : '#ccc';"

' Layout
Response.Write "    var layout = {"
Response.Write "        title: {"
Response.Write "            text: 'Quarterly Prescription Error Trends<br>Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "',"
Response.Write "            font: { color: textColor }"
Response.Write "        },"
Response.Write "        font: { color: textColor },"
Response.Write "        xaxis: {"
Response.Write "            title: '',"
Response.Write "            type: 'category',"
Response.Write "            titlefont: { color: textColor },"
Response.Write "            tickfont: { color: textColor },"
Response.Write "            gridcolor: gridColorTheme"
Response.Write "        },"
Response.Write "        yaxis: {"
Response.Write "            title: 'Dispensed Quantity',"
Response.Write "            titlefont: { color: textColor },"
Response.Write "            tickfont: { color: textColor },"
Response.Write "            gridcolor: gridColorTheme"
Response.Write "        },"
Response.Write "        paper_bgcolor: chartBackgroundColor,"
Response.Write "        plot_bgcolor: chartBackgroundColor,"
Response.Write "        height: 450,"
Response.Write "        width: window.innerWidth * 0.98,"
Response.Write "        barmode: 'stack',"
Response.Write "        margin: { t: 50, b: 80, l: 60, r: 10 },"
Response.Write "        legend: {"
Response.Write "            orientation: 'h',"
Response.Write "            x: 0.5,"
Response.Write "            xanchor: 'center',"
Response.Write "            y: -0.2,"
Response.Write "            yanchor: 'top',"
Response.Write "            font: { color: textColor }"
Response.Write "        }"
Response.Write "    };"

Response.Write "    Plotly.newPlot('quarterlyChartDiv', traces, layout);"
Response.Write "});"
Response.Write "</script>"



'##################################################################################

    ' DataTable Initialization
   Response.Write "<script>"
Response.Write "    document.addEventListener('DOMContentLoaded', function() {"

Response.Write "        var dbDataQuarterlyGender = " & jsonQuarterData & "; "
Response.Write ""

Response.Write "        new DataTable('#proportionTable', {"
Response.Write "            data: dbDataQuarterlyGender.data,"
'response.write "            columns: ["
'response.write "            { data: 'counter',title:'S/No.' },"
'
'response.write "            {"
'response.write "                data: null,"
'response.write "                title: 'Period',"
'response.write "                render: function(data, type, row) {"
'response.write "                    return row.DispenseYear + ' Q' + row.DispenseQuarter;"
'response.write "                }"
'response.write "            },"
'
'
'response.write "            { data: 'DrugName',title:'Drug' },"
'
'response.write "            { data: 'QuarterQty', title:'Qty Dispensed', render: function(data) { " & _
'               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
'
'response.write "            { data: 'PrevQuarterQty', title:'Prev.  Quarter\'s Qty', render: function(data) { " & _
'               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
'
'response.write "            { data: 'AbsoluteChange', title:'Difference', render: function(data) { " & _
'               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
'
'response.write "            { data: 'PercentageChange', title:'% Change', render: function(data) { " & _
'               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
'
'    response.write "            { data: 'CumulativeQty', title:'Cumulative Qty', render: function(data) { " & _
'               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
'
'    response.write "            { data: 'PrescriptionNarrative',title:'Prescription Narrative' },"

 Response.Write "        columns: ["
    Response.Write "            { data: 'counter',title:'S/No.' },"
    Response.Write "            { data: 'DispenseYear',title:'Year' },"
     Response.Write "            { data: 'DispenseCount', title:'Total Dispense Count', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
               
    Response.Write "            { data: 'PrescriptionErrorCount', title:'Prescription Error Count', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
     Response.Write "            { data: 'ErrorRatePercent', title:'Error Rate Percent', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString()+'%' ; } },"
    
    
    Response.Write "            { data: 'Category',title:'Appropriateness Of' },"
    
    
     Response.Write "            { data: 'Appropriateness',title:'Appropriateness' },"
   Response.Write "            { data: 'CurrentQuarterCount', title:'Current Quarter Count', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
   Response.Write "            { data: 'PreviousQuarterCount', title:'Prev. Quarter Count', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
   Response.Write "            { data: 'AbsoluteChange', title:'Difference', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
               
'    response.write "            { data: 'PercentageChange',title:'% Change' },"
Response.Write "            { data: 'QoQ_PercentChange', title:'% Change', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() +'%'; } },"
    
   
    Response.Write "            { data: 'QI_Narrative',title:'Narrative' },"


Response.Write "            ],"
Response.Write "            lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
Response.Write "            dom: 'lBfrtip',"
Response.Write "            buttons: ["
Response.Write "                {"
Response.Write "                    extend: 'csv',"
Response.Write "                    text: 'CSV',"
Response.Write "                    title: 'Quarterly Prescription Errors Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
Response.Write "                },"
Response.Write "                {"
Response.Write "                    extend: 'excel',"
Response.Write "                    text: 'EXCEL',"
Response.Write "                    title: 'Quarterly Prescription Errors Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
Response.Write "                },"
'response.write "                {"
'response.write "                    extend: 'pdf',"
'response.write "                    text: 'PDF',"
'response.write "                    title: 'Quarterly Prescription Errors Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
'response.write "                },"
'response.write "                {"
'response.write "                    extend: 'print',"
'response.write "                    text: 'PRINT',"
'response.write "                    title: 'Quarterly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
'response.write "                }"
Response.Write "            ]"
Response.Write "        });"
Response.Write "    });"
Response.Write "</script>"

    
End Sub

Sub monthly_most_dispensed_medications()

        Response.Write "  <div class='chart-container'>"
        Response.Write "    <div id='monthlyChartDiv' class='chart table-responsive'>"
        Response.Write "  </div>"
        Response.Write "   <div class=""card"">"
        Response.Write "   <div class=""card-body"">"
        Response.Write "       <div class='table-responsive'>"
        Response.Write "      <table class=""table dataTable align-middle table-hover table-bordered"" style=""width: 100%;"" id=""monthlyTable"" >"
        Response.Write "      <thead >"
        Response.Write "              <tr>"
        Response.Write "                <th></th>"
        Response.Write "                <th></th>"
        Response.Write "                <th></th>"
        Response.Write "                <th></th>"
        Response.Write "                <th></th>"
        Response.Write "                <th></th>"
        Response.Write "                <th></th>"
        Response.Write "                <th></th>"
        Response.Write "                <th></th>"
        Response.Write "                <th></th>"
        Response.Write "              </tr>"
        Response.Write "            </thead>"
        Response.Write "        </table>"
        Response.Write "       </div>"
        Response.Write "      </div>"
        Response.Write "  </div>"

 
 Dim sql, rst, cmd
    Set rst = CreateObject("ADODB.RecordSet")
    Set cmd = CreateObject("ADODB.Command")

    sql = "exec sp_monthly_most_dispensed_medications '" & brnchID & "','" & periodStart & "','" & periodEnd & "'"

    ' Set up the Command object
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1
    Set rst = cmd.Execute

Dim jsonMonthlyData

If Not rst.EOF Then
    jsonMonthlyData = "{""data"":" & rst.Fields(0).value & "}"
Else
    jsonMonthlyData = "{""data"":[]}"
End If


rst.Close
Set rst = Nothing
 
Response.Write "<script>"
Response.Write "var dbDataGenderMonthly = " & jsonMonthlyData & ";"
Response.Write "document.addEventListener('DOMContentLoaded', function() {"
Response.Write "    var data = dbDataGenderMonthly.data;"

' Month names
Response.Write "    var monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];"

' Extract unique sorted months
Response.Write "    var uniqueMonths = [...new Set(data.map(d => +d.DispenseMonth))].sort((a, b) => a - b);"
Response.Write "    var periods = uniqueMonths.map(m => ({ mon: m, label: monthNames[m - 1] }));"

' Extract unique (drug + year) combinations
Response.Write "    var drugYearKeys = [...new Set(data.map(d => d.DrugName + '|' + d.DispenseYear))];"

' Color palette
Response.Write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"

' Build traces
Response.Write "    var traces = drugYearKeys.map((key, i) => {"
Response.Write "        var parts = key.split('|');"
Response.Write "        var drug = parts[0];"
Response.Write "        var year = parts[1];"
Response.Write "        var y = [], hover = [];"
Response.Write "        periods.forEach(p => {"
Response.Write "            var match = data.find(d => d.DrugName === drug && d.DispenseYear.toString() === year && +d.DispenseMonth === p.mon);"
Response.Write "            if (match) {"
Response.Write "                y.push(parseFloat(match.QuarterQty));"
Response.Write "                hover.push("
Response.Write "                    'Drug: ' + match.DrugName + '<br>' +"
Response.Write "                    'Year: ' + match.DispenseYear + '<br>' +"
Response.Write "                    'Month: ' + p.label + '<br>' +"
Response.Write "                    'Dispensed Quantity: ' + parseFloat(match.QuarterQty).toLocaleString() + '<br>' +"
Response.Write "                    'Prev. Month\'s Qty: ' + (isNaN(parseFloat(match.PrevQuarterQty)) || match.PrevQuarterQty === null ? 'NA' : parseFloat(match.PrevQuarterQty).toLocaleString()) + '<br>' +"
Response.Write "                    'Difference: ' + (isNaN(parseFloat(match.AbsoluteChange)) || match.AbsoluteChange === null ? 'NA' : parseFloat(match.AbsoluteChange).toLocaleString()) + '<br>' +"
Response.Write "                    '% Change: ' + (isNaN(parseFloat(match.PercentageChange)) || match.PercentageChange === null ? 'NA' : parseFloat(match.PercentageChange).toLocaleString()+'%') + '<br>' +"
Response.Write "                    'Cumulative Qty: ' + parseFloat(match.CumulativeQty).toLocaleString()"
Response.Write "                );"
Response.Write "            } else {"
Response.Write "                y.push(null);"
Response.Write "                hover.push('No data');"
Response.Write "            }"
Response.Write "        });"
Response.Write "        return {"
Response.Write "            x: periods.map(p => p.label),"
Response.Write "            y: y,"
Response.Write "            type: 'scatter',"
Response.Write "            mode: 'lines+markers',"
Response.Write "            name: drug + ' ' + year,"
Response.Write "            text: hover,"
Response.Write "            hovertemplate: '%{text}',"
Response.Write "            marker: { color: colors[i % colors.length] }"
Response.Write "        };"
Response.Write "    });"

' Theme detection
Response.Write "    var theme = localStorage.getItem('theme') || 'light';"
Response.Write "    var textColor = theme === 'dark' ? '#fff' : '#000';"
Response.Write "    var bg = theme === 'dark' ? '#333' : '#f0f0f0';"
Response.Write "    var grid = theme === 'dark' ? '#444' : '#ccc';"

' Layout
Response.Write "    var layout = {"
Response.Write "        title: {"
Response.Write "            text: 'Monthly Dispensing Trends<br>Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "',"
Response.Write "            font: { color: textColor }"
Response.Write "        },"
Response.Write "        font: { color: textColor },"
Response.Write "        xaxis: {"
Response.Write "            title: '',"
Response.Write "            tickangle: 0,"
Response.Write "            gridcolor: grid"
Response.Write "        },"
Response.Write "        yaxis: {"
Response.Write "            title: 'Dispensed Quantity',"
Response.Write "            gridcolor: grid"
Response.Write "        },"
Response.Write "        paper_bgcolor: bg,"
Response.Write "        plot_bgcolor: bg,"
Response.Write "        legend: {"
Response.Write "            orientation: 'h',"
Response.Write "            x: 0.5,"
Response.Write "            y: -0.2,"
Response.Write "            xanchor: 'center',"
Response.Write "            font: { color: textColor }"
Response.Write "        },"
Response.Write "        height: 600,"
Response.Write "        width: window.innerWidth * 0.98"
Response.Write "    };"

' Render the chart
Response.Write "    Plotly.newPlot('monthlyChartDiv', traces, layout);"
Response.Write "});"
Response.Write "</script>"

 '####################################################################################################
    

    ' DataTable Initialization
    Response.Write "<script>"
    Response.Write "    new DataTable('#monthlyTable', {"
    Response.Write "        data: dbDataGenderMonthly.data,"
    Response.Write "        columns: ["
    Response.Write "            { data: 'counter',title:'S/No.' },"
    Response.Write "            { data: 'DispenseYear' ,title:'Year'},"
    Response.Write "            { data: 'DispenseMonthName',title:'Month' },"
    Response.Write "            { data: 'DrugName',title:'Drug' },"
'    response.write "            { data: 'QuarterQty',title:'Qty Dispensed' },"
    Response.Write "            { data: 'QuarterQty', title:'Qty Dispensed', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
   Response.Write "            { data: 'PrevQuarterQty', title:'Prev.  Month\'s Qty', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
               
   Response.Write "            { data: 'AbsoluteChange', title:'Difference', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"

Response.Write "            { data: 'PercentageChange', title:'% Change', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
    Response.Write "            { data: 'CumulativeQty', title:'Cumulative Qty', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
    
    Response.Write "            { data: 'PrescriptionNarrative',title:'Prescription Narrative' },"
    Response.Write "        ],"
    Response.Write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    Response.Write "        dom: 'lBfrtip',"
    Response.Write "        buttons: ["
    Response.Write "            {"
    Response.Write "                extend: 'csv',"
    Response.Write "                text: 'CSV',"
    Response.Write "                title: 'Monthly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'excel',"
    Response.Write "                text: 'EXCEL',"
    Response.Write "                title: 'Monthly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'pdf',"
    Response.Write "                text: 'PDF',"
    Response.Write "                title: 'Monthly Dispensing Trends Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "',"
    Response.Write "            },"
'    response.write "            {"
'    response.write "                extend: 'print',"
'    response.write "                text: 'PRINT',"
'    response.write "                title: '" & brnchName & " Inter Visit Intervals From: " & FormatDateNew(periodStart) & " To: " & FormatDateNew(periodEnd) & "'"
'    response.write "            }"
    Response.Write "        ]"
    Response.Write "    });"
    Response.Write "</script>"
    
End Sub

Sub detailed_prescription_errors()
   

    Response.Write "  <div class='chart-container'>"
    Response.Write "    <div id='weeklyVisitsChartDiv' class='chart table-responsive'>"
    Response.Write "  </div>"
    Response.Write "   <div class=""card"">"
    Response.Write "   <div class=""card-body"">"
    Response.Write "       <div class='table-responsive'>"
    Response.Write "<h2 style='text-align:center;color:blue; margin-top:20px;'>Prescription Errors Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "</h2>"
    Response.Write "      <table class=""table dataTable align-middle table-hover table-bordered"" style=""width: 100%;"" id=""weeklyTable"" >"
'    response.write "      <thead >"
'    response.write "              <tr>"
'    response.write "                <th></th>"
'    response.write "                <th></th>"
'    response.write "                <th></th>"
'    response.write "                <th></th>"
'    response.write "                <th></th>"
'    response.write "                <th></th>"
'    response.write "                <th></th>"
'    response.write "                <th></th>"
'    response.write "                <th></th>"
'    response.write "                <th></th>"
'     response.write "                <th></th>"
'    response.write "                <th></th>"
'    response.write "              </tr>"
'    response.write "            </thead>"

Response.Write "      <thead>"
Response.Write "        <tr>"
Response.Write "          <th rowspan='2'>S/No.</th>"
Response.Write "          <th rowspan='2'>Date</th>"
Response.Write "          <th rowspan='2'>Type Of Patient</th>"
Response.Write "          <th rowspan='2'>Name Of Patient</th>"
Response.Write "          <th rowspan='2'>Age</th>"
Response.Write "          <th rowspan='2'>Drug</th>"
'response.write "          <th colspan='4'>AppropriatenessOf</th>"
Response.Write "          <th colspan='4' style='text-align: center;'>Appropriateness Of</th>"

Response.Write "          <th rowspan='2'>Therapeutic Duplication</th>"
Response.Write "          <th rowspan='2'>Comments</th>"
Response.Write "        </tr>"

Response.Write "        <tr>"
Response.Write "          <th>Drug</th>"
Response.Write "          <th>Dose</th>"
Response.Write "          <th>Frequency</th>"
Response.Write "          <th>Route</th>"
Response.Write "        </tr>"
Response.Write "      </thead>"

    Response.Write "        </table>"
    Response.Write "       </div>"
    Response.Write "      </div>"
    Response.Write "  </div>"

   
   Dim sql, rst, cmd
    Set rst = CreateObject("ADODB.RecordSet")
    Set cmd = CreateObject("ADODB.Command")

    sql = "exec sp_detailed_prescription_errors '" & periodStart & "','" & periodEnd & "'"

    ' Set up the Command object
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1
    Set rst = cmd.Execute

Dim jsonWeeklyData

'If Not rst.EOF Then
'    jsonWeeklyData = "{""data"":" & rst.Fields(0).value & "}"
'Else
'    jsonWeeklyData = "{""data"":[]}"
'End If
'
'
'rst.Close
'Set rst = Nothing

' Handle possible no data or NULL field
If Not rst.EOF And Not IsNull(rst.Fields(0).value) Then
    ' Ensure it's valid JSON (basic check)
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

'response.write "<script>"
'response.write "var dbDataWeekly = " & jsonWeeklyData & ";"
'response.write "document.addEventListener('DOMContentLoaded', function() {"
'response.write "    var data = dbDataWeekly.data;"
'
'' Extract unique sorted week numbers
'response.write "    var uniqueWeeks = [...new Set(data.map(d => +d.DispenseWeek))].sort((a, b) => a - b);"
'response.write "    var weeks = uniqueWeeks.map(w => ({ wk: w, label:  w }));"
'
'' Extract unique (DrugName + Year) combinations for legend
'response.write "    var drugYearKeys = [...new Set(data.map(d => d.DrugName + '|' + d.DispenseYear))];"
'
'' Color palette
'response.write "    var colors = ['#FF6347', '#4682B4', '#32CD32', '#FFD700', '#4B0082', '#FF69B4', '#8B4513', '#00CED1', '#DC143C', '#2F4F4F'];"
'
'' Build traces per drug-year
'response.write "    var traces = drugYearKeys.map((key, i) => {"
'response.write "        var parts = key.split('|');"
'response.write "        var drug = parts[0];"
'response.write "        var year = parts[1];"
'response.write "        var y = [], hover = [];"
'response.write "        weeks.forEach(w => {"
'response.write "            var match = data.find(d => d.DrugName === drug && d.DispenseYear.toString() === year && +d.DispenseWeek === w.wk);"
'response.write "            if (match) {"
'response.write "                y.push(parseFloat(match.QuarterQty));"
'response.write "                hover.push("
'response.write "                    'Drug: ' + match.DrugName + '<br>' +"
'response.write "                    'Year: ' + match.DispenseYear + '<br>' +"
'response.write "                    'Week: ' + w.label + '<br>' +"
'response.write "                    'Dispensed Quantity: ' + parseFloat(match.QuarterQty).toLocaleString() + '<br>' +"
'response.write "                    'Prev. Week\'s Qty: ' + (isNaN(parseFloat(match.PrevQuarterQty)) || match.PrevQuarterQty === null ? 'NA' : parseFloat(match.PrevQuarterQty).toLocaleString()) + '<br>' +"
'response.write "                    'Difference: ' + (isNaN(parseFloat(match.AbsoluteChange)) || match.AbsoluteChange === null ? 'NA' : parseFloat(match.AbsoluteChange).toLocaleString()) + '<br>' +"
'response.write "                    '% Change: ' + (isNaN(parseFloat(match.PercentageChange)) || match.PercentageChange === null ? 'NA' : parseFloat(match.PercentageChange).toLocaleString() + '%') + '<br>' +"
'response.write "                    'Cumulative Qty: ' + (isNaN(parseFloat(match.CumulativeQty)) || match.CumulativeQty === null ? 'NA' : parseFloat(match.CumulativeQty).toLocaleString())"
'response.write "                );"
'response.write "            } else {"
'response.write "                y.push(null);"
'response.write "                hover.push('No data');"
'response.write "            }"
'response.write "        });"
'response.write "        return {"
'response.write "            x: weeks.map(w => w.label),"
'response.write "            y: y,"
'response.write "            type: 'scatter',"
'response.write "            mode: 'lines+markers',"
'response.write "            name: drug + ' ' + year,"
'response.write "            text: hover,"
'response.write "            hovertemplate: '%{text}',"
'response.write "            marker: { color: colors[i % colors.length] }"
'response.write "        };"
'response.write "    });"
'
'' Theme handling
'response.write "    var theme = localStorage.getItem('theme') || 'light';"
'response.write "    var textColor = theme === 'dark' ? '#fff' : '#000';"
'response.write "    var bg = theme === 'dark' ? '#333' : '#f0f0f0';"
'response.write "    var grid = theme === 'dark' ? '#444' : '#ccc';"
'
'' Layout configuration
'response.write "    var layout = {"
'response.write "        title: {"
'response.write "            text: 'Weekly Dispensing Trends<br>Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "',"
'response.write "            font: { color: textColor }"
'response.write "        },"
'response.write "        font: { color: textColor },"
'response.write "        xaxis: {"
'response.write "            title: 'Week Number',"
'response.write "            tickangle: 0,"
'response.write "            gridcolor: grid"
'response.write "        },"
'response.write "        yaxis: {"
'response.write "            title: 'Dispensed Quantity',"
'response.write "            gridcolor: grid"
'response.write "        },"
'response.write "        paper_bgcolor: bg,"
'response.write "        plot_bgcolor: bg,"
'response.write "        legend: {"
'response.write "            orientation: 'h',"
'response.write "            x: 0.5,"
'response.write "            y: -0.2,"
'response.write "            xanchor: 'center',"
'response.write "            font: { color: textColor }"
'response.write "        },"
'response.write "        height: 600,"
'response.write "        width: window.innerWidth * 0.98"
'response.write "    };"
'
'response.write "    Plotly.newPlot('weeklyVisitsChartDiv', traces, layout);"
'response.write "});"
'response.write "</script>"

   
   
Response.Write "<script>"
Response.Write "document.addEventListener('DOMContentLoaded', function() {"
Response.Write "var dbDataWeekly = " & jsonWeeklyData & ";"
'response.write "    var ageGroup = " & jsonWeeklyData & ";"

Response.Write "    new DataTable('#weeklyTable', {"
'response.write "        data: ageGroup.dataAgeGroup.data,"
Response.Write "        data: dbDataWeekly.data,"
Response.Write "        columns: ["
 Response.Write "            { data: 'counter',title:'S/No.' },"
    Response.Write "            { data: 'visitdate' ,title:'Date'},"
    Response.Write "            { data: 'TypeOfPatient',title:'Type Of Patient' },"
    Response.Write "            { data: 'patientname',title:'Name Of Patient' },"
    Response.Write "            { data: 'Age', title:'Age', render: function(data) { " & _
               "var val = parseFloat(data); return isNaN(val) ? 'NA' : val.toLocaleString() ; } },"
        Response.Write "            { data: 'drugname' ,title:'Drug'},"
        Response.Write "            { data: 'AppropriatenessOfDrug' ,title:'Drug'},"
        Response.Write "            { data: 'AppropriatenessOfDose' ,title:'Dose'},"
          Response.Write "            { data: 'AppropriatenessOfFrequency' ,title:'Frequency'},"
        Response.Write "            { data: 'AppropriatenessOfRoute' ,title:'Route'},"
        Response.Write "            { data: 'TherapeuticDuplication' ,title:'Therapeutic Duplication'},"
        
    Response.Write "            { data: 'Comments',title:'Comments' },"
Response.Write "        ],"
Response.Write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
Response.Write "        dom: 'lBfrtip',"
Response.Write "        buttons: ["
Response.Write "            {"
Response.Write "                extend: 'csv',"
Response.Write "                text: 'CSV',"
Response.Write "                title: 'Prescription Errors Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
Response.Write "            },"
Response.Write "            {"
Response.Write "                extend: 'excel',"
Response.Write "                text: 'EXCEL',"
Response.Write "                title: 'Prescription Errors Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
Response.Write "            },"
'response.write "            {"
'response.write "                extend: 'pdf',"
'response.write "                text: 'PDF',"
'response.write "                title: 'Prescription Errors Between " & FormatDateNew(periodStart) & " And " & FormatDateNew(periodEnd) & " - " & GetComboName("Branch", brnchID) & "'"
'response.write "            }"
Response.Write "        ]"
Response.Write "    });"

Response.Write "    hideSpinner();"
Response.Write "});"

' Function to hide spinner
Response.Write "function hideSpinner() {"
Response.Write "    document.getElementById('spinner').style.display = 'none';"
Response.Write "}"
Response.Write "</script>"


End Sub

Sub filters()
    Response.Write "    <main class=""container-fluid"">"
    Response.Write "        <div class=""content"">"
    Response.Write "            <!-- start: page body area -->"
    Response.Write "            <!-- <div class=""px-xl-5 px-lg-4 px-3 py-3 page-body""> -->"
    Response.Write "            <div class="" px-0 py-3 page-body"">"
    Response.Write "                <div class=""card mb-1"">"
    Response.Write "                    <div class=""card-body"">"
    Response.Write "                        <div class=""row g-3"">"
    SetBranch
    ' SetAgeGroup
    Response.Write "                            <div class=""col-md-6 col-lg-3 col-xl-3"">"
    Response.Write "                                <div class=""form-group"">"
    Response.Write "                                    <div class=""form-group mb-3"">"
    Response.Write "                                        <small class=""form-label text-muted"">From:</small>"
    Response.Write "                                        <input type=""date"" class=""form-control"" id ='startDate'>"
    Response.Write "                                    </div>"
    Response.Write "                                </div>"
    Response.Write "                            </div>"
    Response.Write "                            <div class=""col-md-6 col-lg-3 col-xl-3"">"
    Response.Write "                                <div class=""form-group"">"
    Response.Write "                                    <div class=""form-group mb-3"">"
    Response.Write "                                        <small class=""form-label text-muted"">To:</small>"
    Response.Write "                                        <input type=""date"" class=""form-control"" id ='endDate'>"
    Response.Write "                                    </div>"
    Response.Write "                                </div>"
    Response.Write "                            </div>"
    Response.Write "                            <div class=""col-md-6 col-lg-3 col-xl-3 d-flex align-items-center justify-content-center "">"
    Response.Write "                                <button type=""submit "" class=""btn btn-primary"" onclick=""PeriodOnclick()"">Apply</button>"
    Response.Write "                            </div>"
    Response.Write "                        </div>"
    Response.Write "                    </div>"
    Response.Write "                </div>"
    Response.Write "            </div>"
    Response.Write ""
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

    htStr = htStr & "             for (var i = 0; i < dropdown.options.length; i++) { " & vbCrLf
    htStr = htStr & "               if (dropdown.options[i].value === defaultBranch) { " & vbCrLf
    htStr = htStr & "                   dropdown.selectedIndex = i; " & vbCrLf
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
    htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionErrorRatesBA&PositionForTableName=WorkingDay';" & vbCrLf
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
    htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionErrorRatesBA&PositionForTableName=WorkingDay';" & vbCrLf
    htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&month=' + mth  + ' &yearid=' + dayid + '&NoGlobal=1' ;" & vbCrLf
    htStr = htStr & "window.location.href = processurl(ur);" & vbCrLf
    htStr = htStr & "}" & vbCrLf
    'emrdata
    htStr = htStr & "function PeriodOnclick(){ " & vbCrLf
    htStr = htStr & "var branchID1 =  document.getElementById('Branchs').value;" & vbCrLf
    htStr = htStr & "var startDate1 =  document.getElementById('startDate').value;" & vbCrLf
    htStr = htStr & "var endDate1 =  document.getElementById('endDate').value;" & vbCrLf
    htStr = htStr & "startDate1 = startDate1.trimEnd();" & vbCrLf
    htStr = htStr & "endDate1 = endDate1.trimEnd();" & vbCrLf
    htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionErrorRatesBA&PositionForTableName=WorkingDay';" & vbCrLf
    htStr = htStr & "ur = ur + '&WorkingDayID=DAY20160401&selectedValue=' + startDate1  + ' &selectedValue1=' + endDate1 + ' &branID=' + branchID1 + '&NoGlobal=1';" & vbCrLf
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
    htStr = htStr & "ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionErrorRatesBA&PositionForTableName=WorkingDay';" & vbCrLf
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

    htStr = htStr & "</script>"
    Response.Write htStr
    js = js & "<script>" & vbCrLf
    js = js & "  " & vbCrLf
    js = js & "  " & vbCrLf
    js = js & "</script>"
    Response.Write js
End Sub

Sub SetBranch()
    Set rst = CreateObject("ADODB.Recordset")
    dyHt = "<div class='col-md-6 col-lg-3 col-xl-3'>"
    dyHt = dyHt & "<div class='form-group'>"
    dyHt = dyHt & "<small class='form-label text-muted'>Facility:</small>"
    dyHt = dyHt & "<select class='form-select' name='Branchs' id='Branchs' onchange='BranchOnchange()'>"
    
    ' Default option
    dyHt = dyHt & "<option value=''></option>"

    sql0 = "select BranchID, BranchName from Branch b"
    
    ' Open the recordset and loop through the results
    With rst
        .open qryPro.FltQry(sql0), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                branchIDD = Trim(.Fields("BranchID"))
                branchName11 = Trim(.Fields("BranchName"))

                ' Check if this branch is selected
                If UCase(CStr(yearId)) = UCase(branchIDD) Then
                    dyHt = dyHt & "<option value='" & CStr(branchIDD) & "' selected>" & branchName11 & "</option>"
                Else
                    dyHt = dyHt & "<option value='" & CStr(branchIDD) & "'>" & branchName11 & "</option>"
                End If

                rst.MoveNext
                Response.flush
            Loop
        End If
        .Close
    End With
    
    ' Close the select and divs
    dyHt = dyHt & "</select>"
    dyHt = dyHt & "</div>"
    dyHt = dyHt & "</div>"
    
    ' Output the final HTML
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

    ' Array of month names
    Dim monthNames
    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    Dim monthName
    monthName = monthNames(CInt(monthPart) - 1) ' Subtract 1 for zero-based index

    formatedDate = dayPart & "-" & monthName & "-" & yearPart
    FormatDateNew = formatedDate
End Function

Sub ShowLoading()
    Response.Write "<script>"
    Response.Write "        var spinner = document.getElementById('spinner');"
    Response.Write "            spinner.style.display = 'flex';"  ' Show the spinner
    Response.Write "</script>"

End Sub

Sub HideLoading()
    Response.Write "<script>"
    Response.Write "        var spinner = document.getElementById('spinner');"
    Response.Write "            spinner.style.display = 'none';"  ' Hide the spinner
    Response.Write "</script>"

End Sub



'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>









