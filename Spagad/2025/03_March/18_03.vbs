Sub tableStyles()
    response.write "<style>"
        ' [Previous table styles remain unchanged] '
        
        response.write ".header-container {"
        response.write "    display: flex;"
        response.write "    justify-content: center;"
        response.write "    width: 100%;"
        response.write "    margin: 20px 0;"
        response.write "}"
        
        response.write ".filters {"
        response.write "    display: flex;"
        response.write "    flex-direction: column;"
        response.write "    width: 400px;"
        response.write "    padding: 20px;"
        response.write "    border: 1px solid #ccc;"
        response.write "    background-color: #f9f9f9;"
        response.write "    border-radius: 8px;"
        response.write "}"
        
        response.write ".filter-row {"
        response.write "    margin-bottom: 15px;"
        response.write "    width: 100%;"
        response.write "}"
        
        response.write ".date-container {"
        response.write "    display: flex;"
        response.write "    justify-content: space-between;"
        response.write "    gap: 10px;"
        response.write "}"
        
        response.write ".date-field {"
        response.write "    flex: 1;"
        response.write "}"
        
        response.write "input[type='date'] {"
        response.write "    width: 100%;"
        response.write "    padding: 8px;"
        response.write "    border-radius: 4px;"
        response.write "    border: 1px solid #aaa;"
        response.write "    font-size: 14px;"
        response.write "}"
        
        ' [Previous styles for .font-style, .myselect, .mybutton remain unchanged] '
        
    response.write "</style>"
End Sub

Sub header()
    Dim dropdownOptions

    sql = "SELECT MedicalOutcomeID, MedicalOutcomeName FROM MedicalOutcome"
    
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.open sql, conn, 3, 4

    dropdownOptions = "<option value=''>" & "All" & "</option>"

    With rstDropdown
        If .recordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                optionHTML = "<option value='" & .fields("MedicalOutcomeID") & "'>" & .fields("MedicalOutcomeName") & "</option>"
                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With

    rstDropdown.Close
    Set rstDropdown = Nothing

    response.write "<div class='header-container'>"
    response.write "<div class='filters'>"
    
    ' Medical Outcome Row
    response.write "<div class='filter-row'>"
    response.write "    <label for='medicalOutcome' class='font-style'>Select Medical Outcome:</label>"
    response.write "    <select class='myselect' id='medicalOutcome' name='medicalOutcome'>"
    response.write dropdownOptions
    response.write "    </select>"
    response.write "</div>"
    
    ' Date Filters Row
    response.write "<div class='filter-row'>"
    response.write "    <div class='date-container'>"
    response.write "        <div class='date-field'>"
    response.write "            <label for='from' class='font-style'>From</label>"
    response.write "            <input type='date' name='from' id='from'>"
    response.write "        </div>"
    response.write "        <div class='date-field'>"
    response.write "            <label for='to' class='font-style'>To</label>"
    response.write "            <input type='date' name='to' id='to'>"
    response.write "        </div>"
    response.write "    </div>"
    response.write "</div>"
    
    ' Button Row
    response.write "<div class='filter-row'>"
    response.write "    <button class='mybutton' type='button' onclick='updateUrl()'>Show Data</button>"
    response.write "</div>"
    
    response.write "</div>"
    response.write "</div>"

    ' [JavaScript remains unchanged] '
    response.write "<script>"
    response.write "    function updateUrl() {"
    response.write "        const fromDate = document.getElementById('from').value;"
    response.write "        const toDate = document.getElementById('to').value;"
    response.write "        const medicalOutcomes = Array.from(document.getElementById('medicalOutcome').selectedOptions).map(option => option.value).join(',');"
    response.write "        const baseUrl = 'http://172.19.0.36/hms/wpgPrtPrintLayoutAll.asp';"
    response.write "        const params = new URLSearchParams({"
    response.write "            PrintLayoutName: 'MortalityReport',"
    response.write "            PositionForTableName: 'WorkingDay',"
    response.write "            WorkingDayID: '',"
    response.write "            Dateperiod: fromDate + '||' + toDate,"
    response.write "            MedicalOutcomeID: medicalOutcomes"
    response.write "        });"
    response.write "        const newUrl = baseUrl + '?' + params.toString();"
    response.write "        window.location.href = newUrl;"
    response.write "        console.log(newUrl);"
    response.write "    }"
    response.write "</script>"
End Sub

http://172.19.0.36/hms/wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&VisitationID=V1250314191&WorkFlowNav=POP