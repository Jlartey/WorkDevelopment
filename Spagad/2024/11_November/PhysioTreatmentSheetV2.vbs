'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
'Response.Write "Hello Joe"

Dim currentTime
currentTime = Now

Entries
DisplayTable
tableStyles
Sub Entries()
    
    response.write "  <body>"
    response.write "    <style>"
    response.write "      body {"
    response.write "        font-family: Arial, sans-serif;"
    response.write "      }"
    response.write "      .container {"
    response.write "        padding: 20px 20px;"
    response.write "        margin:  auto;"
    response.write "        width: 600px;"
    response.write "        line-height: 1.6;"
    response.write "      }"
    response.write "      .row {"
    response.write "        display: flex;"
    response.write "        margin-bottom: 10px;"
    response.write "      }"
    response.write "      label {"
    response.write "        width: 200px;"
    response.write "        margin-right: 20px;"
    response.write "        text-align: right;"
    response.write "      }"
    response.write "      .value {"
    response.write "        flex: 1;"
    response.write "      }"
    response.write "      button {"
    response.write "        width: 75px;"
    response.write "        margin: auto;"
    response.write "        margin-left: 160px;"
    response.write "        border-radius: 10px;"
    response.write "        background-color: blue;"
    response.write "        color: white;"
    response.write "        padding: 10px 10px;"
    response.write "        outline: none;"
    response.write "        cursor: pointer;"
    response.write "        font-family: Arial, sans-serif;"
    response.write "      }"
    response.write "    </style>"
    response.write "  </head>"
    response.write "  <body>"
    response.write "    <div class=""container"">"
    response.write "      <div class=""row"">"
    response.write "        <label for=""treatment-date"">Treatment Date :</label>"
    response.write "        " & currentTime & " "
    response.write "      </div>"
    response.write "      <div class=""row"">"
    response.write "        <label for=""treatment-type"">Type :</label>"
    response.write "        <select name=""treatment-type"" id=""treatment-type"">"
    response.write "          <option value=""Therapeutic Exercises"">Therapeutic Exercises</option>"
    response.write "          <option value=""Manual"">Manual</option>"
    response.write "          <option value=""Other"">Other</option>"
    response.write "        </select>"
    response.write "      </div>"
    response.write "      <div class=""row"">"
    response.write "        <label for=""intervention"">Intervention :</label>"
    response.write "        <input type=""text"" name=""intervention"" id=""intervention"" style=""width: 300px"" />"
    response.write "      </div>"
    response.write "      <div class=""row"">"
    response.write "        <label for=""treatment-value"">Value :</label>"
    response.write "        <input type=""text"" name=""treatment-value"" id=""treatment-value"" style=""width: 300px""/>"
    response.write "      </div>"
    response.write "      <button onclick=""updateTreatment()"">SAVE</button>"
    response.write "    </div>"
    
    response.write "<script>"
    response.write "        function updateTreatment(){"
    response.write "            const treatmentDate = new Date();"
    response.write "            const treatmentType = document.getElementById('treatment-type');"
    response.write "            const intervention = document.getElementById('intervention');"
    response.write "            const treatmentValue = document.getElementById('treatment-value');"
    response.write "            "
    response.write "            let url = 'wpgXMLHTTP.asp?procedurename=UpdatePhysioTreatment';"
    response.write "            if (treatmentDate && treatmentType?.value && intervention?.value && treatmentValue?.value) {"
    response.write "                url = url + '&treatmentDate=' + treatmentDate.toISOString().split('T')[0];"
    response.write "                url = url + '&treatmentType=' + encodeURIComponent(treatmentType.value);"
    response.write "                url = url + '&intervention=' + encodeURIComponent(intervention.value);"
    response.write "                url = url + '&treatmentValue=' + encodeURIComponent(treatmentValue.value);"
    response.write " fetch(url)"
    response.write "   .then((response) => response.json())"
    response.write "   .then((results) => {"
    response.write "     if (results.success) {"
    response.write "       const tableContainer = document.querySelector('.mytable>tbody');"
    response.write "       const tr = document.createElement('tr');"
    response.write "       tr.innerHTML = `<td class='mytd' id='count'>${"
    response.write "         tableContainer.querySelectorAll('tr').length + 1"
    response.write "       }</td>"
    response.write "                       <td class='mytd' >${"
    response.write "                         results.data.KeyPrefix.split('||')[0]"
    response.write "                       }</td>"
    response.write "                       <td class='mytd'>${results.data.performvarName}</td>"
    response.write "                       <td class='mytd'>${results.data.Description}</td>"
    response.write "                       <td class='mytd'>${"
    response.write "                         results.data.KeyPrefix.split('||')[1]"
    response.write "                       }</td>`;"
    response.write "       tableContainer.prepend(tr);"
    response.write "        const tds = tableContainer.querySelectorAll('[id=""count""]');"
    response.write "        tds.forEach((el,index) => {"
    response.write "            index++;"
    response.write "            el.innerText = index;"
    response.write "        });"
    response.write "     }"
    response.write "   })"
    response.write "   .catch((error) => console.error('Fetch error:', error));"
    response.write "            } else {"
    response.write "                alert('Please enter all values');"
    response.write "            }"
    response.write "        }"
    response.write "    </script>"
    response.write "  </body>"
End Sub

Sub DisplayTable()
    Dim count, sql, rst, treatmentDate, treatmentValue
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "Select * from PerformVar11"

    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<thead>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Date</th>"
            response.write "<th class='myth'>Type</th>"
            response.write "<th class='myth'>Intervention</th>"
            response.write "<th class='myth'>Value</th>"
            response.write "</tr class='mytr'>"
            response.write "</thead>"
            
            response.write "<tbody>"
            Do While Not .EOF
                treatmentDate = Split(.fields("KeyPrefix"), "||")(0)
                treatmentValue = Split(.fields("KeyPrefix"), "||")(1)
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd' id='count'>" & count & "</td>"
                response.write "<td class='mytd'>" & treatmentDate & "</td>"
                response.write "<td class='mytd'>" & .fields("PerformVar11Name") & "</td>"
                response.write "<td class='mytd'>" & .fields("Description") & "</td>"
                response.write "<td class='mytd'>" & treatmentValue & "</td>"
                response.write "</tr >"

                .MoveNext
                count = count + 1
            Loop
            response.write "</tbody>"
            response.write "</table>"
        Else
            response.write "<h1>No records found</h1>"
        End If
        
        .Close
    End With
    
    Set rst = Nothing
End Sub

Sub tableStyles()
    response.write "<style>"
        response.write ".mytable {"
        response.write "    width: 65vw;"
        response.write "    border-collapse: collapse;"
        response.write "    margin: 20px 0;"
        response.write "    font-size: 16px;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "}"
        response.write ".mytable, .myth, .mytd {"
        response.write "    border: 1px solid #dddddd;"
        response.write "}"
        response.write ".myth, .mytd {"
        response.write "    padding: 12px;"
        response.write "    text-align: left;"
        response.write "}"
        response.write ".myth {"
        response.write "    background-color: #f2f2f2;"
        response.write "    color: #333;"
        response.write "    font-weight: bold;"
        response.write "}"
        response.write ".mytr:nth-child(even) {"
        response.write "    background-color: #f9f9f9;"
        response.write "}"
        response.write ".mytr:hover {"
        response.write "    background-color: #f1f1f1;"
        response.write "}"
        response.write ".myth {"
        response.write "    text-transform: uppercase;"
        response.write "}"
        response.write "h1 {"
        response.write "    font-size: 18px;"
        response.write "    color: #555;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "    margin: 20px 0;"
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
