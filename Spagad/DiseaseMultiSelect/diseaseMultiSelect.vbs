Sub DisplayDiseasesDropdown()
    Dim rst, sql

    Set rst = CreateObject("ADODB.RecordSet")

    ' SQL for populating the multiselect field
    sql = "SELECT DiseaseID, DiseaseName FROM Disease order by DiseaseName "

    response.write "<h3>Select Diseases</h3>"

    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            .MoveFirst
            
            response.write "<form id='mulSelForm'>"
            response.write "    <div class='multiselect-container'>"
            response.write "       <label for='disease'>Select Disease:</label><br>"
            response.write "       <input type='text' id='searchDisease' placeholder='Search disease...' onkeyup='filterDiseases()' class='search-input'><br>"
            response.write "       <select id='disease' name='disease' multiple class='mult-select-tag' onchange='updateSelectedNames()'>"
            
            Do Until .EOF
                response.write "<option value='" & .fields("DiseaseID") & "'>" & .fields("DiseaseName") & "</option>"
                .MoveNext
            Loop
            response.write "        </select>"
            response.write "    </div>"
            response.write "</form>"
        End If
        .Close
    End With

    response.write "<script src='https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js'></script>"
    response.write "<script>"
    response.write "    function filterDiseases() {"
    response.write "        var input, filter, select, options, option, i, txtValue;"
    response.write "        input = document.getElementById('searchDisease');"
    response.write "        filter = input.value.toUpperCase();"
    response.write "        select = document.getElementById('disease');"
    response.write "        options = select.getElementsByTagName('option');"
    response.write "        for (i = 0; i < options.length; i++) {"
    response.write "            option = options[i];"
    response.write "            txtValue = option.textContent || option.innerText;"
    response.write "            if (txtValue.toUpperCase().indexOf(filter) > -1) {"
    response.write "                option.style.display = '';"
    response.write "            } else {"
    response.write "                option.style.display = 'none';"
    response.write "            }"
    response.write "        }"
    response.write "    }"

    response.write "    function updateSelectedNames() {"
    response.write "        var select = document.getElementById('disease');"
    response.write "        var selectedNames = [];"
    response.write "        for (var i = 0; i < select.options.length; i++) {"
    response.write "            if (select.options[i].selected) {"
    response.write "                selectedNames.push(select.options[i].text);"
    response.write "            }"
    response.write "        }"
    response.write "        console.log('Selected Diseases:', selectedNames);"
    response.write "    }"

    response.write "</script>"
End Sub

Sub Styling()
    response.write "<style>"
    response.write "    body {"
    response.write "        font-family: Arial, sans-serif;"
    response.write "        background-color: #f2f2f2;"
    response.write "        padding: 20px;"
    response.write "    }"
    response.write "    .multiselect-container {"
    response.write "        max-width: 300px;"
    response.write "        margin: 10px auto;"
    response.write "        background-color: #fff;"
    response.write "        border: 1px solid #ccc;"
    response.write "        padding: 10px;"
    response.write "        border-radius: 5px;"
    response.write "        box-shadow: 0 0 10px rgba(0,0,0,0.1);"
    response.write "    }"
    response.write "    .search-input {"
    response.write "        width: 100%;"
    response.write "        padding: 8px;"
    response.write "        margin-bottom: 10px;"
    response.write "        border: 1px solid #ccc;"
    response.write "        border-radius: 3px;"
    response.write "        box-sizing: border-box;"
    response.write "        font-size: 14px;"
    response.write "    }"
    response.write "    .mult-select-tag {"
    response.write "        width: 100%;"
    response.write "        border: 1px solid #ccc;"
    response.write "        padding: 8px;"
    response.write "        height: 150px;"
    response.write "        overflow-y: auto;"
    response.write "        border-radius: 3px;"
    response.write "    }"
    response.write "</style>"
End Sub




' Call the subroutine to display diseases dropdown
DisplayDiseasesDropdown
Styling




