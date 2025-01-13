'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Styling
MultiSelectStyles
dispPatientVisitations

Sub dispPatientVisitations()

    Dim sql, periodStart, periodEnd, datePeriod, count, selectedWardIDs, idsArr, formattedIDs, id, dateArr
    Dim rstDropdown, rstMain
    Dim dropdownOptions, optionHTML

    ' Retrieve query parameters
    datePeriod = Trim(request.querystring("Dateperiod"))
    selectedWardIDs = Trim(request.querystring("WardID"))

    ' Parse date period
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If

    ' Format selected drug store IDs
    If selectedWardIDs <> "" Then
        idsArr = Split(selectedWardIDs, ",")
        For Each id In idsArr
            formattedIDs = formattedIDs & "'" & Trim(id) & "',"
        Next
        ' Remove the trailing comma
        formattedIDs = Left(formattedIDs, Len(formattedIDs) - 1)
    End If

    ' Construct SQL query for dropdown options (all pharmacies)
    sql = "Select WardID, Wardname from Ward"

    ' Initialize and open database connection for dropdown options
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.open sql, conn, 3, 4

    ' Populate dropdown options
    dropdownOptions = ""

    With rstDropdown
        If .recordCount > 0 Then
            .movefirst
            Do Until .EOF
                optionHTML = "<option value='" & .fields("WardID") & "'>" & .fields("Wardname") & "</option>"
                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With

    ' Close dropdown recordset
    rstDropdown.Close
    Set rstDropdown = Nothing

    ' Construct SQL query for main data
    sql = "select "
    sql = sql & "Patient.PatientID, "
    sql = sql & "Patient.Patientname, "
    sql = sql & "VisitType.VisitTypeName, "
    sql = sql & "convert(varchar(20), Visitation.VisitDate, 103) VisitDate, "
    sql = sql & "AdmissionStatus.AdmissionStatusName, "
    sql = sql & "convert(varchar(20), Admission.AdmissionDate, 103) AdmissionDate, "
    sql = sql & "convert(varchar(20), Admission.DischargeDate, 103) DischargeDate, "
    sql = sql & "Ward.WardName, "
    sql = sql & "Bed.BedName "
    sql = sql & "from "
    sql = sql & "Patient "
    sql = sql & "join "
    sql = sql & "Visitation "
    sql = sql & "on Visitation.PatientID = Patient.PatientID "
    sql = sql & "join "
    sql = sql & "VisitType on Visitation.VisitTypeID = VisitType.VisitTypeID "
    sql = sql & "join "
    sql = sql & "Admission on Admission.VisitationID = Visitation.VisitationID "
    sql = sql & "join "
    sql = sql & "AdmissionStatus on AdmissionStatus.AdmissionStatusID = Admission.AdmissionStatusID "
    sql = sql & "join "
    sql = sql & "Ward on Admission.WardID = Ward.WardID "
    sql = sql & "join "
    sql = sql & "Bed on Admission.BedID = Bed.BedID "
    sql = sql & "where "
     If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "convert(date, Visitation.VisitDate) between '" & periodStart & "' and '" & periodEnd & "' "
    Else
        sql = sql & "convert(date, Visitation.VisitDate) between '2018-01-01' and '2018-04-01' "
    End If
    
    If selectedWardIDs <> "" Then
        sql = sql & "and Ward.WardID IN (" & formattedIDs & ")"
    Else
        sql = sql & "and Ward.WardID IS NOT NULL "
    End If
    
    sql = sql & "order by "
    sql = sql & "convert(date, Admission.AdmissionDate) desc"


    'response.write sql
    
    ' Initialize and open database connection for main data
    Set rstMain = CreateObject("ADODB.Recordset")
    rstMain.open sql, conn, 3, 4

    ' Output dropdown
    response.write "<div>"
    response.write "        <label for='ward' class='font-style'>Select Ward:</label><br>"
    response.write "        <select id='ward' name='ward' multiple class='mult-select-tag'>"
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
        response.write "<h2 class='font-style'>SHOWING DATA FROM: 2018-01-01 TO: 2018-01-31</h2>"
    End If

    ' Reset recordset to the first record
    rstMain.movefirst

    ' Display table
    response.write "<table class='mytable'>"
    response.write "<tr>"
        response.write "<th class='myth'>No.</th>"
        response.write "<th class='myth'>Patient Name</th>"
        response.write "<th class='myth'>Visit Type</th>"
        response.write "<th class='myth'>Visit Date</th>"
        response.write "<th class='myth'>Admission Status</th>"
        response.write "<th class='myth'>Admission Date</th>"
        response.write "<th class='myth'>Discharge Date</th>"
        response.write "<th class='myth'>Ward</th>"
        response.write "<th class='myth'>Bed</th>"
    response.write "</tr>"

    count = 0

    With rstMain
        If .recordCount > 0 Then
            .movefirst
            Do While Not .EOF
                count = count + 1
                response.write "<tr>"
                response.write "<td class='mytd' align='center'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("Patientname") & "</td>"
                response.write "<td class='mytd' align='center'>" & .fields("VisitTypeName") & "</td>"
                response.write "<td class='mytd' align='center'>" & .fields("VisitDate") & "</td>"
                response.write "<td class='mytd' align='center'>" & .fields("AdmissionStatusName") & "</td>"
                response.write "<td class='mytd' align='center'>" & .fields("AdmissionDate") & "</td>"
                response.write "<td class='mytd' align='center'>" & .fields("DischargeDate") & "</td>"
                response.write "<td class='mytd' align='center'>" & .fields("WardName") & "</td>"
                response.write "<td class='mytd' align='center'>" & .fields("BedName") & "</td>"
                response.write "</tr>"
                response.flush
                .MoveNext
            Loop
        End If
    End With

    response.write "</table>"

    ' Close main data recordset
    rstMain.Close
    Set rstMain = Nothing

    ' Output scripts
    response.write "<script src='https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js'></script>"
    response.write "<script>"
        response.write "    new MultiSelectTag('ward', {"
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
        response.write "        const ward = Array.from(document.getElementById('ward').selectedOptions).map(option => option.value).join(',');"
        response.write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp';"
        response.write "        const params = new URLSearchParams({"
        response.write "            PrintLayoutName: 'dispPatientVisitations',"
        response.write "            PositionForTableName: 'WorkingDay',"
        response.write "            WorkingDayID: '',"
        response.write "            Dateperiod: `${fromDate}||${toDate}`,"
        response.write "            WardID: ward"
        response.write "        });"
        response.write "        const newUrl = baseUrl + '?' + params.toString();"
        response.write "        window.location.href = newUrl;"
        response.write "        console.log(newUrl);"
        response.write "    }"
    response.write "</script>"

End Sub

Sub Styling()
    response.write " <style>"
        response.write " .mytable {"
        response.write "     width: 85vw;"
        response.write "     font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
        response.write "     border-collapse: collapse;"
        response.write "     margin-top: 50px; "
        response.write "     border-radius: 10px;"
        response.write " }"
        
        
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
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
