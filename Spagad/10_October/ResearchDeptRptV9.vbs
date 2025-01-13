'I need to imporove upon this one
addCSS
DatePicker
Dim sql, rst, datePeriod, cnt
datePeriod = Split(Request.QueryString("DatePeriod"), "||")

Sub DatePicker()
    response.write "<form id='dateForm'>"
    response.write "    <div class='container' style='display: flex;'>"
    response.write "        <div>"
    response.write "            <label for='from'>From</label>"
    response.write "            <input type='date' name='from' id='from'>"
    response.write "        </div>"
    response.write "        <div>"
    response.write "            <label for='to'>To</label>"
    response.write "            <input type='date' name='to' id='to'>"
    response.write "        </div>"
    response.write "        <div>"
    response.write "            <label for='researchForm'>Select Form</label>"
    response.write "            <select name='researchForm' id='researchForm' class='cta-form'>"
    response.write "                <option value=''>Select Form</option>"
    response.write "                <option value='SRS'>SRS</option>"
    response.write "                <option value='HOOS'>HOOS</option>"
    response.write "                <option value='KOOS'>KOOS</option>"
    response.write "            </select>"
    response.write "        </div>"
    response.write "        <div>"
    response.write "            <button type='button' onclick='processPrint()'>Process Print</button>"
    response.write "        </div>"
    response.write "    </div>"
    response.write "</form>"
    response.write "<script>"
    response.write "    function processPrint() {"
    response.write "        const fromDate = document.getElementById('from').value;"
    response.write "        const toDate = document.getElementById('to').value;"
    response.write "        const researchForm = document.getElementById('researchForm').value;"
    response.write "        if (!fromDate || !toDate || !researchForm) {"
    response.write "            alert('Please select the form and date range.');"
    response.write "            return;"
    response.write "        }"
    response.write "        const url = 'http://172.2.2.33/hms/wpgPrtPrintLayoutAll.asp?';"
    response.write "        const params = new URLSearchParams({"
    response.write "            PrintLayoutName: 'DispResearchRpt',"
    response.write "            DatePeriod: fromDate + '||' + toDate,"
    response.write "            ResearchForm: researchForm"
    response.write "        });"
    response.write "        window.location.href = url + params.toString();"
    response.write "    }"
    response.write "</script>"
End Sub

Sub displaySRS()
    ' Your existing SRS report generation code
End Sub

Sub displayHOOS()
    ' Your existing HOOS report generation code
End Sub

Sub displayKOOS()
    ' Your existing KOOS report generation code
End Sub

Sub addCSS()
    response.write "<style>"
    response.write "  .container { display: flex; justify-content: space-between; padding: 20px; }"
    response.write "  .cta-form { padding: 6px; font-size: 15px; font-family: inherit; color: inherit; border: none; background-color: #f2f2f2; border-radius: 9px; box-shadow: 0 1px 2px rgba(0, 0, 0, 0.1); }"
    response.write "  table { width: 100%; border-collapse: collapse; margin: 20px 0; }"
    response.write "  th, td { padding: 8px; text-align: left; border: 1px solid #ddd; }"
    response.write "  tr:nth-child(even) { background-color: #f2f2f2; }"
    response.write "</style>"
End Sub

Sub showReport()
    Select Case Request.QueryString("ResearchForm")
        Case "SRS"
            displaySRS()
        Case "HOOS"
            displayHOOS()
        Case "KOOS"
            displayKOOS()
        Case Else
            response.write "<p>Please select a valid form and try again.</p>"
    End Select
End Sub

DatePicker
showReport
