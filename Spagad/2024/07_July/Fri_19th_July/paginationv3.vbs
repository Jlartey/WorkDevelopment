'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Response.Clear
conn.commandTimeOut = 7200

Dim periodStart, periodEnd, datePeriod, selectedWardIDs, selectedVisitTypeIDs, selectedAdmStatIDs
Dim idsArr, idsArr2, formattedIDs, formattedIDs2, id, id2, dateArr
Dim idsArr3, formattedIDs3, id3

' Retrieve query parameters
    datePeriod = Trim(Request.QueryString("Dateperiod"))
    selectedWardIDs = Trim(Request.QueryString("WardID"))
    selectedVisitTypeIDs = Trim(Request.QueryString("visitTypeID"))
    selectedAdmStatIDs = Trim(Request.QueryString("admissionTypeID"))
    
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

    If selectedVisitTypeIDs <> "" Then
        idsArr2 = Split(selectedVisitTypeIDs, ",")
        For Each id2 In idsArr2
            formattedIDs2 = formattedIDs2 & "'" & Trim(id2) & "',"
        Next
        ' Remove the trailing comma
        formattedIDs2 = Left(formattedIDs2, Len(formattedIDs2) - 1)
    End If
    
    If selectedAdmStatIDs <> "" Then
        idsArr3 = Split(selectedAdmStatIDs, ",")
        For Each id3 In idsArr3
            formattedIDs3 = formattedIDs3 & "'" & Trim(id3) & "',"
        Next
        ' Remove the trailing comma
        formattedIDs3 = Left(formattedIDs3, Len(formattedIDs3) - 1)
    End If
    
Styling
MultiSelectStyles

Response.Write "<!DOCTYPE html>"
Response.Write "<html lang='en'>"
Response.Write "<head>"
Response.Write "<meta charset='UTF-8'>"
Response.Write "<meta name='viewport' content='width=device-width, initial-scale=1.0'>"
Response.Write "<title>Patient Visitations</title>"

Response.Write "    <link href=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css"" rel=""stylesheet"""
Response.Write "        integrity=""sha384-9ndCyUaIbzAi2FUVXJi0CjmCapSmO7SnpJef0486qhLnuZ2cdeRhO02iuK6FUUVM"" crossorigin=""anonymous"">"
Response.Write "    <script src=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"""
Response.Write "        integrity=""sha384-geWF76RCwLtnZ8qwWowPQNguL3RmwHVBC9FhGdlKrxdiJJigb/j/68SIy3Te4Bkz"""
Response.Write "        crossorigin=""anonymous""></script>"
Response.Write " <link href=""https://cdn.datatables.net/v/bs5/jq-3.6.0/jszip-2.5.0/dt-1.13.5/af-2.6.0/b-2.4.0/b-colvis-2.4.0/b-html5-2.4.0/b-print-2.4.0/cr-1.7.0/date-1.5.0/fc-4.3.0/fh-3.4.0/kt-2.10.0/r-2.5.0/rg-1.4.0/rr-1.4.0/sc-2.2.0/sb-1.5.0/sp-2.2.0/sl-1.7.0/sr-1.3.0/datatables.min.css"" rel=""stylesheet""/>"
Response.Write " <script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js""></script>"
Response.Write " <script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js""></script>"
Response.Write " <script src=""https://cdn.datatables.net/v/bs5/jq-3.6.0/jszip-2.5.0/dt-1.13.5/af-2.6.0/b-2.4.0/b-colvis-2.4.0/b-html5-2.4.0/b-print-2.4.0/cr-1.7.0/date-1.5.0/fc-4.3.0/fh-3.4.0/kt-2.10.0/r-2.5.0/rg-1.4.0/rr-1.4.0/sc-2.2.0/sb-1.5.0/sp-2.2.0/sl-1.7.0/sr-1.3.0/datatables.min.js""></script>"
'response.write "<script src='https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js'></script>"

Response.Write "<style>"
Response.Write "  .chart-container {"
Response.Write "    display: flex;"
Response.Write "    justify-content: center;"
Response.Write "  }"
Response.Write "  .chart {"
Response.Write "    flex: 1;"
Response.Write "    margin: 10px;"
Response.Write "    width: 80%;"
Response.Write "  }"
Response.Write "  .tab-header {"
Response.Write "    display: flex;"
Response.Write "    justify-content: center;"
Response.Write "    background-color: #007bff;"
Response.Write "    border: 1px solid #ddd;"
Response.Write "    border-radius: 5px;"
Response.Write "  }"
Response.Write "  .tab-button {"
Response.Write "    flex: 1;"
Response.Write "    padding: 10px;"
Response.Write "    text-align: center;"
Response.Write "    cursor: pointer;"
Response.Write "    font-weight: bold;"
Response.Write "    color: #fff;"
Response.Write "    border-right: 1px solid #ddd;"
Response.Write "  }"
Response.Write "  .tab-button:last-child {"
Response.Write "    border-right: none;"
Response.Write "  }"
Response.Write "  .tab-button.active {"
Response.Write "    background-color: #0056b3;"
Response.Write "  }"
Response.Write "  .tab-content {"
Response.Write "    display: none;"
Response.Write "    padding: 20px;"
Response.Write "    border: 1px solid #ddd;"
Response.Write "    border-radius: 5px;"
Response.Write "    background-color: #f9f9f9;"
Response.Write "    margin-top: 10px;"
Response.Write "  }"
Response.Write "  .tab-content.active {"
Response.Write "    display: block;"
Response.Write "  }"
Response.Write "</style>"


Response.Write "</head>"
Response.Write "<body>"


  ' Construct SQL query for dropdown options (all pharmacies)
    sql = "Select WardID, Wardname from Ward where WardID IN ('001', '002', '003', 'W001', 'W002', 'W003', 'W005', 'W007')"

    ' Initialize and open database connection for dropdown options
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.open sql, conn, 3, 4

    ' Populate dropdown options
    dropdownOptions = ""

    With rstDropdown
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                optionHTML = "<option value='" & .Fields("WardID") & "'>" & .Fields("Wardname") & "</option>"
                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With

    ' Close dropdown recordset
    rstDropdown.Close
    Set rstDropdown = Nothing
    
    
    'Construct SQL for VisitType dropdown options
    sql = "select VisitTypeID, VisitTypeName from VisitType"
    
    Set rstVisitDropdown = CreateObject("ADODB.Recordset")
    rstVisitDropdown.open sql, conn, 3, 4
    
    visitTypeDropdown = ""
    
    With rstVisitDropdown
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                visitTypeOptions = "<option value='" & .Fields("VisitTypeID") & "'>" & .Fields("VisitTypeName") & "</option>"
                visitTypeDropdown = visitTypeDropdown & visitTypeOptions
                .MoveNext
            Loop
        End If
    End With
    
    ' Close dropdown recordset
    rstVisitDropdown.Close
    Set rstVisitDropdown = Nothing
    
    
    'Construct SQL for AdmissionStatus dropdown options
    sql = "select AdmissionStatusID, AdmissionStatusName from AdmissionStatus"
    
    Set rstAdmStatDropdown = CreateObject("ADODB.Recordset")
    rstAdmStatDropdown.open sql, conn, 3, 4
    
    admStatDropdown = ""
    
    With rstAdmStatDropdown
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                admStatOptions = "<option value='" & .Fields("AdmissionStatusID") & "'>" & .Fields("AdmissionStatusName") & "</option>"
                admStatDropdown = admStatDropdown & admStatOptions
                .MoveNext
            Loop
        End If
    End With
    
    ' Close dropdown recordset
    rstAdmStatDropdown.Close
    Set rstAdmStatDropdown = Nothing
    
    Response.Write "<div class='header'>"
' Output dropdown
    Response.Write "<div>"
    Response.Write "        <label for='ward' class='font-style'>Select Ward:</label><br>"
    Response.Write "        <select id='ward' name='ward' multiple class='mult-select-tag'>"
    Response.Write dropdownOptions
    Response.Write "        </select>"
    Response.Write "</div>"
    
    Response.Write "<div>"
    Response.Write "        <label for='visitType' class='font-style'>Select Visit Type:</label><br>"
    Response.Write "        <select id='visitType' name='visitType' multiple class='mult-select-tag'>"
    Response.Write visitTypeDropdown
    Response.Write "        </select>"
    Response.Write "</div>"
    
    Response.Write "<div>"
    Response.Write "        <label for='admStat' class='font-style'>Select Admission Type:</label><br>"
    Response.Write "        <select id='admStat' name='admStat' multiple class='mult-select-tag'>"
    Response.Write admStatDropdown
    Response.Write "        </select>"
    Response.Write "</div>"
    
    
    ' Output HTML Form for date selection
    Response.Write "    <form id='dateForm'>"
    Response.Write "    <div class='container' style='display: flex; align-items: center; justify-content: center'> "
    Response.Write "        <div> "
    Response.Write "            <label for='from'>From</label> "
    Response.Write "            <input type='date' name='from' id='from'> "
    Response.Write "        </div> "
    Response.Write "        <div> "
    Response.Write "            <label for='to' style='margin-left: 10px'>To</label> "
    Response.Write "            <input type='date' name='to' id='to'> "
    Response.Write "        </div> "
    Response.Write "        <div> "
    Response.Write "            <button type='button' onclick='updateUrl()'>Show Data</button> <br />"
    Response.Write "        </div>    "
    Response.Write "    </div> "
    Response.Write "   </form>"
    
    Response.Write "</div>"
    
     If (periodStart <> "" And periodEnd <> "") Then
        Response.Write "<h2 class='font-style'>SHOWING DATA FROM: " & periodStart & " TO: " & periodEnd & "</h2>"
    Else
        Response.Write "<h2 class='font-style'>SHOWING DATA FROM: 2018-01-01 TO: 2018-01-31</h2>"
    End If

Response.Write "<div id='yearlyTab' class='tab-content active'>"
Response.Write "  <div class='chart-container'>"
Response.Write "    <div id='yearlyChartDiv' class='chart'></div>"
Response.Write "  </div>"

' yearly table

  Response.Write "      <table style=""width:100%"" id=""yearlyTable"" class=""table table-striped table-bordered table-sm table-responsive pb-3"" width=""100%"">"
  Response.Write "      <thead class=""table-dark"">"
  Response.Write "        <tr>"
  Response.Write "             <th>S/No.</th>"
  Response.Write "             <th>Patient Name</th>"
  Response.Write "             <th>Visit Type</th>"
  Response.Write "             <th>Visit Date</th>"
  Response.Write "             <th>Admission Status </th>"
  Response.Write "             <th>Admission Date</th>"
  Response.Write "             <th>Discharge Date</th>"
  Response.Write "             <th>Ward</th>"
  Response.Write "             <th>Bed</th>"
  Response.Write "        </tr>"
  Response.Write "       </thead>"
  Response.Write "    </table>"
Response.Write "</div>"

'yearly tab end here

Response.Write "</body>"
Response.Write "</html>"


dispPatientVisitations

Sub dispPatientVisitations()

    Dim sql, count
    Dim dropdownOptions, optionHTML
    
' Construct SQL query for main data
    sql = "select "
    sql = sql & "Patient.PatientName, "
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
        sql = sql & "convert(date, Admission.AdmissionDate) between '" & periodStart & "' and '" & periodEnd & "' "
    Else
        sql = sql & "convert(date, Admission.AdmissionDate) between '2018-01-01' and '2018-04-01' "
    End If
    
    If selectedWardIDs <> "" Then
        sql = sql & "and Ward.WardID IN (" & formattedIDs & ")"
    Else
        sql = sql & "and Ward.WardID IS NOT NULL "
    End If
    
    If selectedVisitTypeIDs <> "" Then
        sql = sql & "and VisitType.VisitTypeID IN (" & formattedIDs2 & ") "
    Else
        sql = sql & "and VisitType.VisitTypeID IS NOT NULL "
    End If
    
    If selectedAdmStatIDs <> "" Then
        sql = sql & "and Admission.AdmissionStatusID IN (" & formattedIDs3 & ") "
    Else
        sql = sql & "and Admission.AdmissionStatusID IS NOT NULL "
    End If
    sql = sql & "order by "
    sql = sql & "convert(date, Admission.AdmissionDate) desc"

    'response.write sql
    
    ' Initialize and open database connection for main data
    Set rstMain = CreateObject("ADODB.Recordset")
    rstMain.open sql, conn, 3, 4
    
    Dim jsonData, counter
    counter = 1
    jsonData = "{""data"":["

    If rstMain.RecordCount > 0 Then
        rstMain.MoveFirst
        Do While Not rstMain.EOF
            jsonData = jsonData & "{"
            jsonData = jsonData & """counter"":""" & counter & ""","
            jsonData = jsonData & """PatientName"":""" & rstMain.Fields("PatientName").Value & ""","
            jsonData = jsonData & """VisitTypeName"":""" & rstMain.Fields("VisitTypeName").Value & ""","
            jsonData = jsonData & """VisitDate"":""" & rstMain.Fields("VisitDate").Value & ""","
            jsonData = jsonData & """AdmissionStatusName"":""" & rstMain.Fields("AdmissionStatusName").Value & ""","
            jsonData = jsonData & """AdmissionDate"":""" & rstMain.Fields("AdmissionDate").Value & ""","
            jsonData = jsonData & """DischargeDate"":""" & rstMain.Fields("DischargeDate").Value & ""","
            jsonData = jsonData & """WardName"":""" & rstMain.Fields("WardName").Value & ""","
            jsonData = jsonData & """BedName"":""" & rstMain.Fields("BedName").Value & """"
            jsonData = jsonData & "},"
            rstMain.MoveNext
            counter = counter + 1
        Loop
        jsonData = Left(jsonData, Len(jsonData) - 1) ' Remove the trailing comma
    End If

    jsonData = jsonData & "]}"

    rstMain.Close
    Set rstMain = Nothing
    
    ' Output scripts
    Response.Write "<script src='https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js'></script>"
    Response.Write "<script>"
        Response.Write "    new MultiSelectTag('ward', {"
        Response.Write "        rounded: true,"
        Response.Write "        shadow: true,"
        Response.Write "        placeholder: 'Search',"
        Response.Write "        tagColor: {"
        Response.Write "            textColor: '#327b2c',"
        Response.Write "            borderColor: '#92e681',"
        Response.Write "            bgColor: '#eaffe6',"
        Response.Write "        },"
        Response.Write "        onChange: function (values) {"
        Response.Write "            console.log(values);"
        Response.Write "        },"
        Response.Write "    });"
        
        
        Response.Write "    new MultiSelectTag('visitType', {"
        Response.Write "        rounded: true,"
        Response.Write "        shadow: true,"
        Response.Write "        placeholder: 'Search',"
        Response.Write "        tagColor: {"
        Response.Write "            textColor: '#327b2c',"
        Response.Write "            borderColor: '#92e681',"
        Response.Write "            bgColor: '#eaffe6',"
        Response.Write "        },"
        Response.Write "        onChange: function (values) {"
        Response.Write "            console.log(values);"
        Response.Write "        },"
        Response.Write "    });"
        
        Response.Write "    new MultiSelectTag('admStat', {"
        Response.Write "        rounded: true,"
        Response.Write "        shadow: true,"
        Response.Write "        placeholder: 'Search',"
        Response.Write "        tagColor: {"
        Response.Write "            textColor: '#327b2c',"
        Response.Write "            borderColor: '#92e681',"
        Response.Write "            bgColor: '#eaffe6',"
        Response.Write "        },"
        Response.Write "        onChange: function (values) {"
        Response.Write "            console.log(values);"
        Response.Write "        },"
        Response.Write "    });"
        
        Response.Write "    function updateUrl() {"
        Response.Write "        const fromDate = document.getElementById('from').value;"
        Response.Write "        const toDate = document.getElementById('to').value;"
        Response.Write "        const ward = Array.from(document.getElementById('ward').selectedOptions).map(option => option.value).join(',');"
        Response.Write "        const visitTypes = Array.from(document.getElementById('visitType').selectedOptions).map(option => option.value).join(',');"
        Response.Write "        const admissionTypes = Array.from(document.getElementById('admStat').selectedOptions).map(option => option.value).join(',');"
        Response.Write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp';"
        Response.Write "        const params = new URLSearchParams({"
        Response.Write "            PrintLayoutName: 'dispPagination',"
        Response.Write "            PositionForTableName: 'WorkingDay',"
        Response.Write "            WorkingDayID: '',"
        Response.Write "            Dateperiod: `${fromDate}||${toDate}`,"
        Response.Write "            WardID: ward,"
        Response.Write "            visitTypeID: visitTypes,"
        Response.Write "            admissionTypeID: admissionTypes"
        Response.Write "        });"
        Response.Write "        const newUrl = baseUrl + '?' + params.toString();"
        Response.Write "        window.location.href = newUrl;"
        Response.Write "        console.log(newUrl);"
        Response.Write "    }"
    Response.Write "</script>"
    
    'response.write jsonData
' ==========================================================================================================================


    ' DataTable Initialization
    Response.Write "<script>"
    Response.Write "var dbDataYearly = " & jsonData & ";"
    Response.Write "    new DataTable('#yearlyTable', {"
    Response.Write "        data: dbDataYearly.data,"
    Response.Write "        columns: ["
    Response.Write "            { data: 'counter' },"
    Response.Write "            { data: 'PatientName' },"
    Response.Write "            { data: 'VisitTypeName' },"
    Response.Write "            { data: 'VisitDate' },"
    Response.Write "            { data: 'AdmissionStatusName' },"
    Response.Write "            { data: 'AdmissionDate' },"
    Response.Write "            { data: 'DischargeDate' },"
    Response.Write "            { data: 'WardName' },"
    Response.Write "            { data: 'BedName' }"
    Response.Write "        ],"
        
    Response.Write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],"
    Response.Write "        dom: 'lBfrtip',"
    Response.Write "        buttons: ["
    Response.Write "            {"
    Response.Write "                extend: 'csv',"
    Response.Write "                text: 'CSV',"
    Response.Write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'excel',"
    Response.Write "                text: 'EXCEL',"
    Response.Write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'pdf',"
    Response.Write "                text: 'PDF',"
    Response.Write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            },"
    Response.Write "            {"
    Response.Write "                extend: 'print',"
    Response.Write "                text: 'PRINT',"
    Response.Write "                title: '" & brnchName & " Patient Visitations From: " & FormatDate(periodStart) & " To: " & FormatDate(periodEnd) & "'"
    Response.Write "            }"
    Response.Write "        ]"
    Response.Write "    });"
    Response.Write "</script>"
End Sub

Sub Styling()
    Response.Write " <style>"
        Response.Write " .mytable {"
        Response.Write "     width: 95vw;"
        Response.Write "     font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
        Response.Write "     border-collapse: collapse;"
        Response.Write "     margin-top: 50px; "
        Response.Write "     border-radius: 10px;"
        Response.Write " }"
        
        Response.Write " .header {"
        Response.Write "    display: flex;"
        Response.Write "    justify-content: center;"
        Response.Write "    align-items: center;"
        Response.Write " } "
        
        Response.Write " .font-style {"
        Response.Write "    text-align: center;"
        Response.Write " } "
        
        Response.Write " .container {"
        Response.Write "    display: flex"
        Response.Write "    margin-top: 50px !important;"
        Response.Write "    padding-top: 30px;"
        Response.Write " } "
        
        Response.Write " .myth, .mytd {"
        Response.Write "     border: 1px solid #ddd;"
        Response.Write "     padding: 10px;"
        Response.Write " }"
        
        Response.Write " .mytd {"
        Response.Write "     text-alig: 1px solid #ddd;"
        Response.Write "     padding: 8px;"
        Response.Write " }"
        
        Response.Write "  tr:nth-child(even) {"
        Response.Write "    background-color: rgba(249, 249, 249, 6);"
        Response.Write " } "
        
        Response.Write " .myth {"
        Response.Write "     background-color: #c2c2c2;"
        Response.Write "     color: black;"
        Response.Write "     text-align: center; "
        Response.Write "     text-transform: uppercase; "
        Response.Write "     font-size: 18px;"
        Response.Write " }"
        
        Response.Write "  button {"
        Response.Write "     background-color: #0236c4;"
        Response.Write "     border-radius: 5px;"
        Response.Write "     border: none;"
        Response.Write "     margin-left: 50px;"
        Response.Write "     padding: 5px 20px;"
        Response.Write "     color: white;"
        Response.Write "     cursor: pointer;"
        Response.Write "  }"
        
        Response.Write "  #to, #from {"
        Response.Write "    padding: 5px;"
        Response.Write "    border-radius: 5px;"
        Response.Write "    cursor: pointer;"
        Response.Write "  }"
        
        Response.Write " .pagination {"
        Response.Write "    text-align: center;"
        Response.Write "    margin: 20px 0;"
        Response.Write " }"
        
        Response.Write " .pagination a {"
        Response.Write "    margin: 0 5px;"
        Response.Write "    padding: 10px 15px;"
        Response.Write "    background-color: #f1f1f1;"
        Response.Write "    border: 1px solid #ccc;"
        Response.Write "    text-decoration: none;"
        Response.Write "    color: #333;"
        Response.Write " }"
        
        Response.Write " .pagination a:hover {"
        Response.Write "    background-color: #ddd;"
        Response.Write " }"
        
        Response.Write " .font-style {"
        Response.Write "    font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif;"
        Response.Write " }"
        
        Response.Write " #pharmacy {"
        Response.Write "    padding-bottom: 10px;"
        Response.Write " }"
        Response.Write " </style>"
        
End Sub

Sub MultiSelectStyles()
     Response.Write "    <style>" & vbCrLf
    Response.Write "        .mult-select-tag {" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            width: 300px;" & vbCrLf
    Response.Write "            flex-direction: column;" & vbCrLf
    Response.Write "            align-items: center;" & vbCrLf
    Response.Write "            position: relative;" & vbCrLf
    Response.Write "            --tw-shadow: 0 1px 3px 0 rgb(0 0 0 / 0.1), 0 1px 2px -1px rgb(0 0 0 / 0.1);" & vbCrLf
    Response.Write "            --tw-shadow-color: 0 1px 3px 0 var(--tw-shadow-color), 0 1px 2px -1px var(--tw-shadow-color);" & vbCrLf
    Response.Write "            --border-color: rgb(218, 221, 224);" & vbCrLf
    Response.Write "            font-family: Verdana, sans-serif;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .wrapper {" & vbCrLf
    Response.Write "            width: 100%;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .body {" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            border: 1px solid var(--border-color);" & vbCrLf
    Response.Write "            background: #fff;" & vbCrLf
    Response.Write "            min-height: 2.15rem;" & vbCrLf
    Response.Write "            width: 100%;" & vbCrLf
    Response.Write "            min-width: 14rem;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .input-container {" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            flex-wrap: wrap;" & vbCrLf
    Response.Write "            flex: 1 1 auto;" & vbCrLf
    Response.Write "            padding: 0.1rem;" & vbCrLf
    Response.Write "            align-items: center;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .input-body {" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            width: 100%;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .input {" & vbCrLf
    Response.Write "            flex: 1;" & vbCrLf
    Response.Write "            background: 0 0;" & vbCrLf
    Response.Write "            border-radius: 0.25rem;" & vbCrLf
    Response.Write "            padding: 0.45rem;" & vbCrLf
    Response.Write "            margin: 10px;" & vbCrLf
    Response.Write "            color: #2d3748;" & vbCrLf
    Response.Write "            outline: 0;" & vbCrLf
    Response.Write "            border: 1px solid var(--border-color);" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .btn-container {" & vbCrLf
    Response.Write "            color: #e2ebf0;" & vbCrLf
    Response.Write "            padding: 0.5rem;" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            border-left: 1px solid var(--border-color);" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag button {" & vbCrLf
    Response.Write "            cursor: pointer;" & vbCrLf
    Response.Write "            width: 100%;" & vbCrLf
    Response.Write "            color: #718096;" & vbCrLf
    Response.Write "            outline: 0;" & vbCrLf
    Response.Write "            height: 100%;" & vbCrLf
    Response.Write "            border: none;" & vbCrLf
    Response.Write "            padding: 0;" & vbCrLf
    Response.Write "            background: 0 0;" & vbCrLf
    Response.Write "            background-image: none;" & vbCrLf
    Response.Write "            -webkit-appearance: none;" & vbCrLf
    Response.Write "            text-transform: none;" & vbCrLf
    Response.Write "            margin: 0;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag button:first-child {" & vbCrLf
    Response.Write "            width: 1rem;" & vbCrLf
    Response.Write "            height: 90%;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .drawer {" & vbCrLf
    Response.Write "            position: absolute;" & vbCrLf
    Response.Write "            background: #fff;" & vbCrLf
    Response.Write "            max-height: 15rem;" & vbCrLf
    Response.Write "            z-index: 40;" & vbCrLf
    Response.Write "            top: 98%;" & vbCrLf
    Response.Write "            width: 100%;" & vbCrLf
    Response.Write "            overflow-y: scroll;" & vbCrLf
    Response.Write "            border: 1px solid var(--border-color);" & vbCrLf
    Response.Write "            border-radius: 0.25rem;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag ul {" & vbCrLf
    Response.Write "            list-style-type: none;" & vbCrLf
    Response.Write "            padding: 0.5rem;" & vbCrLf
    Response.Write "            margin: 0;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag ul li {" & vbCrLf
    Response.Write "            padding: 0.5rem;" & vbCrLf
    Response.Write "            border-radius: 0.25rem;" & vbCrLf
    Response.Write "            cursor: pointer;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag ul li:hover {" & vbCrLf
    Response.Write "            background: rgb(243 244 246);" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .item-container {" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            justify-content: center;" & vbCrLf
    Response.Write "            align-items: center;" & vbCrLf
    Response.Write "            padding: 0.2rem 0.4rem;" & vbCrLf
    Response.Write "            margin: 0.2rem;" & vbCrLf
    Response.Write "            font-weight: 500;" & vbCrLf
    Response.Write "            border: 1px solid;" & vbCrLf
    Response.Write "            border-radius: 9999px;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .item-label {" & vbCrLf
    Response.Write "            max-width: 100%;" & vbCrLf
    Response.Write "            line-height: 1;" & vbCrLf
    Response.Write "            font-size: 0.75rem;" & vbCrLf
    Response.Write "            font-weight: 400;" & vbCrLf
    Response.Write "            flex: 0 1 auto;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .item-close-container {" & vbCrLf
    Response.Write "            display: flex;" & vbCrLf
    Response.Write "            flex: 1 1 auto;" & vbCrLf
    Response.Write "            flex-direction: row-reverse;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .item-close-svg {" & vbCrLf
    Response.Write "            width: 1rem;" & vbCrLf
    Response.Write "            margin-left: 0.5rem;" & vbCrLf
    Response.Write "            height: 1rem;" & vbCrLf
    Response.Write "            cursor: pointer;" & vbCrLf
    Response.Write "            border-radius: 9999px;" & vbCrLf
    Response.Write "            display: block;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .hidden {" & vbCrLf
    Response.Write "            display: none;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .shadow {" & vbCrLf
    Response.Write "            box-shadow: var(--tw-ring-offset-shadow, 0 0 #0000), var(--tw-ring-shadow, 0 0 #0000), var(--tw-shadow);" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        .mult-select-tag .rounded {" & vbCrLf
    Response.Write "            border-radius: 0.375rem;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    </style>" & vbCrLf
End Sub


'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
