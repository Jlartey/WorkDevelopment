--Created on 26th June, 2024

'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Styling
displayPatientBilling

Sub displayPatientBilling()
    Dim sql, periodStart, periodEnd, datePeriod, count
    datePeriod = Trim(Request.QueryString("Dateperiod"))
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If
    
    Set rst = CreateObject("ADODB.RecordSet")

    sql = sql & "SELECT AdmissionID, "
    sql = sql & "Patient.PatientName, "
    sql = sql & "AdmissionName, "
    sql = sql & "convert(varchar(20), Admission.AdmissionDate, 103) AdmissionDate, "
    sql = sql & "convert(varchar(20), Admission.DischargeDate, 103) DischargeDate, "
    sql = sql & "ceiling(datediff(hour, AdmissionDate, DischargeDate)/24.0) NoOfDays"
    sql = sql & "FROM Admission "
    sql = sql & "JOIN AdmissionStatus "
    sql = sql & "ON AdmissionStatus.AdmissionStatusID = Admission.AdmissionStatusID "
    sql = sql & "JOIN Patient "
    sql = sql & "ON Patient.PatientID = Admission.PatientID "
    sql = sql & "WHERE Admission.AdmissionStatusID = 'A007' "
    
    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & "and Admission.AdmissionDate between '" & periodStart & "' and '" & periodEnd & "' "
    Else
        sql = sql & "and Admission.AdmissionDate between '2018-01-01' and '2019-01-31' "
    End If
    sql = sql & "order by Admission.AdmissionDate DESC "
    
    'response.write sql
    
    'Display the DatePicker
    response.write "<h2>Showing Data From " & periodStart & " To " & periodEnd & " </h2>"
    
    response.write "<form id='dateForm'> "
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
    response.write "</form> "

    response.write " <br />"
    response.write "<script> "
    response.write "    function updateUrl() { "
    response.write "        const fromDate = document.getElementById('from').value; "
    response.write "        const toDate = document.getElementById('to').value; "
    response.write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp'; "
    response.write "        const params = new URLSearchParams({ "
    response.write "            PrintLayoutName: 'displayPatientBilling', "
    response.write "            PositionForTableName: 'WorkingDay', "
    response.write "            WorkingDayID: '' ,"
    response.write "            Dateperiod: fromDate + '||' + toDate"
    response.write "        }); "
    response.write "        const newUrl = baseUrl + '?' + params.toString(); "
    response.write "        console.log(newUrl); "
    response.write "        window.location.href = newUrl; "
    response.write "    } "
    response.write "</script> "
    
    response.write " "
    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then

            .MoveFirst
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr>"
                response.write "<th class='myth'> No. </th>"
                response.write "<th class='myth'> AdmissionID </th>"
                response.write "<th class='myth'> Patient Name </th>"
                response.write "<th class='myth'>Admission Date</th>"
                response.write "<th class='myth'>Discharge Date</th>"
                response.write "<th class='myth'>No. Of Days</th>"
            response.write "</tr>"
            Do While Not .EOF
                count = count + 1
                response.write "<tr>"
                    response.write "<td class='mytd' align='center'>" & count & "</td>"
                    response.write "<td class='mytd'>" & .Fields("AdmissionID") & "</td>"
                    response.write "<td class='mytd'>" & .Fields("PatientName") & "</td>"
                    response.write "<td class='mytd' align='center'>" & .Fields("AdmissionDate") & "</td>"
                    response.write "<td class='mytd' align='center'>" & .Fields("DischargeDate") & "</td>"
                    response.write "<td class='mytd' align='center'>" & .Fields("NoOfDays") & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
            response.write "</table>"
        End If
    End With
End Sub

Sub Styling()
    response.write " <style>"
        response.write " table {"
        response.write "     width: 75vw;"
        response.write "     font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
        response.write "     border-collapse: collapse;"
       
        response.write " }"
        
        response.write " .myth, .mytd {"
        response.write "     border: 1px solid #ddd;"
        response.write "     padding: 10px;"
        response.write " }"
        
        response.write " .mytd {"
        response.write "     text-alig: 1px solid #ddd;"
        response.write "     padding: 8px;"
        response.write " }"
        
        response.write "  tr:nth-child(even) {"
        response.write "    background-color: #f9f9f9;"
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
    response.write " </style>"

End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
