Sub DatePicker()
    response.write "<form id='dateForm'> "
    response.write "    <div class='container'> "
    response.write "        <div> "
    response.write "            <label for='from'>From</label> "
    response.write "            <input type='date' name='from' id='from'> "
    response.write "        </div> "
    response.write "        <div> "
    response.write "            <label for='to'>To</label> "
    response.write "            <input type='date' name='to' id='to'> "
    response.write "        </div> "
    response.write "    </div> "
    response.write "</form> "
    
    ' Add CSS for styling
    response.write "<style>"
    response.write "  .container { display: flex; justify-content: center; align-items: center; gap: 20px; margin: 20px auto; padding: 10px; background-color: #f9f9f9; border-radius: 10px; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); }"
    response.write "  label { font-weight: bold; }"
    response.write "  input[type='date'] { padding: 5px; border: 1px solid #ccc; border-radius: 5px; }"
    response.write "</style>"

    ' Set default values and auto-update the URL on input change
    response.write "<script>"
    response.write "    window.onload = function() {"
    response.write "        const today = new Date();"
    response.write "        const sixMonthsAgo = new Date();"
    response.write "        sixMonthsAgo.setMonth(today.getMonth() - 6);"
    response.write "        document.getElementById('from').value = sixMonthsAgo.toISOString().split('T')[0];"
    response.write "        document.getElementById('to').value = today.toISOString().split('T')[0];"
    response.write "    };"
    response.write "    document.getElementById('from').addEventListener('input', updateUrl);"
    response.write "    document.getElementById('to').addEventListener('input', updateUrl);"
    response.write "    function updateUrl() { "
    response.write "        const fromDate = document.getElementById('from').value || new Date().toISOString().split('T')[0];"
    response.write "        const toDate = document.getElementById('to').value || new Date().toISOString().split('T')[0];"
    response.write "        const baseUrl = 'http://172.2.2.33/hms/wpgPrtPrintLayoutAll.asp';"
    response.write "        const params = new URLSearchParams({ "
    response.write "            PrintLayoutName: 'DispResearchRpt', "
    response.write "            PositionForTableName: 'WorkingDay', "
    response.write "            WorkingDayID: '', "
    response.write "            DatePeriod: fromDate + '||' + toDate"
    response.write "        }); "
    response.write "        const newUrl = baseUrl + '?' + params.toString(); "
    response.write "        window.location.href = newUrl; "
    response.write "    } "
    response.write "</script>"
End Sub
