'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
'Response.Write "Hello Joe"
'This is what Ihave done as of 29th August, 2024. I have been directed by Michael to do this using the PrintFilter

Dim periodStart, periodEnd, datePeriod, dateArr
Dim selectedDoctorIDs, idsArr, formattedIDs, id

datePeriod = Trim(Request.querystring("Dateperiod"))
selectedDoctorIDs = Trim(Request.querystring("DoctorID"))

If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    periodStart = dateArr(0)
    periodEnd = dateArr(1)
Else
    periodStart = "2024-08-20"
    periodEnd = "2024-08-28"
End If

 If selectedDoctorIDs <> "" Then
    idsArr = Split(selectedDoctorIDs, ",")
    For Each id In idsArr
        formattedIDs = formattedIDs & "'" & Trim(id) & "',"
    Next
    ' Remove the trailing comma
    formattedIDs = Left(formattedIDs, Len(formattedIDs) - 1)
End If


MultiSelectStyles
dispPatientProcedure

Sub dispPatientProcedure()
    Dim sql, count
    Dim rstDropdown, optionHTML
    
    sql = " SELECT MedicalStaffId, MedicalStaffName FROM MedicalStaff"
    
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.open sql, conn, 3, 4
    
    dropdownOptions = ""
    
    With rstDropdown
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                optionHTML = "<option value='" & .fields("MedicalStaffId") & "'>" & .fields("MedicalStaffName") & "</option>"
                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With
    
    ' Close dropdown recordset
    rstDropdown.Close
    Set rstDropdown = Nothing
    
    
    
    
    
    Set rst = CreateObject("ADODB.Recordset")
    
    'Construct SQL for main query
    sql = "SELECT "
    sql = sql & "TreatCharges.PatientId, "
    sql = sql & "Patient.PatientName, "
    sql = sql & "Patient.age, "
    sql = sql & "Gender.GenderName, "
    sql = sql & "TreatCharges.ConsultReviewDate, "
    sql = sql & "MedicalStaff.MedicalStaffName, "
    sql = sql & "TreatCategory.TreatCategoryName, "
    sql = sql & "Treatment.TreatmentName, "
    sql = sql & "format(TreatCharges.Qty, 'N2') Quantity, "
    sql = sql & "format(TreatCharges.UnitCost, 'N2') UnitCost, "
    sql = sql & "format(TreatCharges.FinalAmt, 'N2') FinalAmount "
    sql = sql & "FROM TreatCharges "
    sql = sql & "JOIN Gender ON TreatCharges.GenderID = Gender.GenderID "
    sql = sql & "JOIN Patient ON Patient.PatientID = TreatCharges.PatientID "
    sql = sql & "JOIN MedicalStaff ON MedicalStaff.MedicalStaffID = TreatCharges.MedicalStaffID "
    sql = sql & "JOIN TreatCategory ON TreatCategory.TreatCategoryID = TreatCharges.TreatCategoryID "
    sql = sql & "JOIN Treatment ON Treatment.TreatmentID = TreatCharges.TreatmentID "
    sql = sql & "WHERE convert(date, TreatCharges.ConsultReviewDate) BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
    
    If selectedDoctorIDs <> "" Then
        sql = sql & "AND MedicalStaff.MedicalStaffID IN (" & formattedIDs & ") "
    Else
        sql = sql & "AND MedicalStaff.MedicalStaffID IS NOT NULL "
    End If
    
    'response.write sql
    
    response.write periodStart
    response.write "<br> "
    response.write periodEnd
    
    ' Output dropdown
    response.write "<div style='display: flex;'> "
    response.write "<div>"
    response.write "        <label for='doctor' class='font-style'>Select Doctor:</label><br>"
    response.write "        <select id='doctor' name='doctor' multiple class='mult-select-tag'>"
    response.write dropdownOptions
    response.write "        </select>"
    response.write "</div>"
    
    'response.write "<form id='dateForm'> "
    response.write "    <div class='container' style='display: flex;'> "
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
    'response.write "</form> "
    response.write "</div> "
    
    response.write " <br />"

    response.write "<script src='https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js'></script>"
    response.write "<script>"
        response.write "    new MultiSelectTag('doctor', {"
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
        
    response.write "    function updateUrl() { "
    response.write "        const fromDate = document.getElementById('from').value; "
    response.write "        const toDate = document.getElementById('to').value; "
    response.write "        const doctor = Array.from(document.getElementById('doctor').selectedOptions).map(option => option.value).join(','); "
    response.write "        const baseUrl = 'http://172.2.2.31/hms/wpgPrtPrintLayoutAll.asp'; "
    response.write "        const params = new URLSearchParams({ "
    response.write "            PrintLayoutName: 'dispPatientProcedure', "
    response.write "            PositionForTableName: 'WorkingDay', "
    response.write "            WorkingDayID: '' ,"
    response.write "            Dateperiod: fromDate + '||' + toDate, "
    response.write "            DoctorID: doctor "
    response.write "        }); "
    response.write "        const newUrl = baseUrl + '?' + params.toString(); "
    response.write "        console.log(newUrl); "
    response.write "        window.location.href = newUrl; "
    response.write "    } "
    response.write "</script> "
    
    response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1'>"
    response.write "<tr>"
        response.write "<th class='myth'>No.</th>"
        response.write "<th class='myth'>PatientID</th>"
        response.write "<th class='myth'>PatientName</th>"
        response.write "<th class='myth'>Age</th>"
        response.write "<th class='myth'>Sex</th>"
        response.write "<th class='myth'>Date</th>"
        response.write "<th class='myth'>Doctor</th>"
        response.write "<th class='myth'>Type</th>"
        response.write "<th class='myth'>Item Name</th>"
        response.write "<th class='myth'>Quantity</th>"
        response.write "<th class='myth'>Unit Cost</th>"
        response.write "<th class='myth'>Final Amount</th>"
    response.write "</tr>"
    
    count = 0
    
    With rst
        .open sql, conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                count = count + 1
                response.write "<tr>"
                    response.write "<td>" & count & "</td>"
                    response.write "<td>" & .fields("PatientId") & "</td>"
                    response.write "<td>" & .fields("PatientName") & "</td>"
                    response.write "<td>" & .fields("age") & "</td>"
                    response.write "<td>" & .fields("GenderName") & "</td>"
                    response.write "<td>" & .fields("ConsultReviewDate") & "</td>"
                    response.write "<td>" & .fields("MedicalStaffName") & "</td>"
                    response.write "<td>" & .fields("TreatCategoryName") & "</td>"
                    response.write "<td>" & .fields("TreatmentName") & "</td>"
                    response.write "<td>" & .fields("Quantity") & "</td>"
                    response.write "<td>" & .fields("UnitCost") & "</td>"
                    response.write "<td>" & .fields("FinalAmount") & "</td>"
                response.write "</tr>"
              .MoveNext
            Loop
        End If
    End With
    
    response.write "</table>"
    
    rst.Close
    Set rst = Nothing
        
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
