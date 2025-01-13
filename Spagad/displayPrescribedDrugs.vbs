DatePicker
Styling
Sub DatePicker()
    Dim rst, sql, periodStart, periodEnd, datePeriod, count
    datePeriod = Trim(request.querystring("Dateperiod"))
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
        
    End If
    count = 0
    Set rst = CreateObject("ADODB.RecordSet")

    sql = "SELECT prescribedDrug, dispensedDrug, CONVERT(VARCHAR(30), prescriptiondate, 103) AS PrescriptionDate, CONVERT(VARCHAR(30), DispenseDate, 103) AS DispenseDate "
    If (periodStart <> "" And periodEnd <> "") Then
        sql = sql & " FROM fn_GetPrescribedDrugs('" & periodStart & "' , '" & periodEnd & "') "
    Else
        sql = sql & " FROM fn_GetPrescribedDrugs('2018-01-01' , '2024-12-31') "
    End If
  
'    response.write sql
'    response.write datePeriod
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
        response.write "        const baseUrl = 'http://192.168.5.11/rmchms01/wpgPrtPrintLayoutAll.asp'; "
        response.write "        const params = new URLSearchParams({ "
        response.write "            PrintLayoutName: 'displayPrescribedDrugs', "
        response.write "            PositionForTableName: 'WorkingDay', "
        response.write "            WorkingDayID: '' ,"
        response.write "            Dateperiod: fromDate + '||' + toDate"
        response.write "        }); "
        response.write "        const newUrl = baseUrl + '?' + params.toString(); "
        response.write "        console.log(newUrl); "
        response.write "        window.location.href = newUrl; "
        response.write "    } "
        response.write "</script> "
            
    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
           
            .MoveFirst
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr>"
                response.write "<th class='myth'> No.</th>"
                response.write "<th class='myth' > PRESCRIBED DRUG </th>"
                'response.write "<th class='myth'> DISPENSED DRUG  </th>"
                response.write "<th class='myth'>  PRESCRIPTION DATE  </th>"
               ' response.write "<th class='myth'> DISPENSE DATE</th>"
            response.write "</tr>"
            Do While Not .EOF
                count = count + 1
                response.write "<tr>"
                response.write "<td class='mytd' align= 'center'>" & count & "</td>"
                response.write "<td class='mytd' align= 'left'>" & .fields("prescribedDrug") & "</td>"
                'response.write "<td class='mytd'>" & .fields("dispensedDrug") & "</td>"
                response.write "<td class='mytd' align= 'center'>" & .fields("PrescriptionDate") & "</td>"
                'response.write "<td class='mytd'>" & .fields("DispenseDate") & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
            response.write "</table>"
        End If
        .Close
    End With
    
End Sub

Sub Styling()
    response.write " <style>"
        response.write " table {"
        response.write "     width: 65vw;"
        response.write "     font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
        response.write "     border-collapse: collapse;"
        response.write " }"
        
        response.write " .myth, .mytd {"
        response.write "     border: 1px solid #ddd;"
        response.write "     padding: 8px;"
        response.write " }"
        
        response.write " .mytd {"
        response.write "     text-alig: 1px solid #ddd;"
        response.write "     padding: 8px;"
        response.write " }"
        
        response.write "  tr:nth-child(even) {"
        response.write "    background-color: #f9f9f9;"
        response.write " } "
        
        response.write " .myth {"
        response.write "     background-color: #f2f2f2;"
        response.write "     color: black;"
        response.write "     text-align: center; "
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

        response.write "     .container {"
        response.write "         align-items: center;"
        response.write "         font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
        response.write "     }"
    response.write " </style>"

End Sub