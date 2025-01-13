'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

DatePicker
Styling
Sub DatePicker()
    Dim rst, sql, periodStart, periodEnd, datePeriod
    datePeriod = Trim(request.QueryString("Dateperiod"))
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If
    
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "WITH DispensedCTE"
    sql = sql & " AS"
    sql = sql & " ("
    sql = sql & " SELECT  Drug.Drugname [DispensedDrug],DrugSaleItems.DispenseDate, DrugsaleItems.VisitationID"
    sql = sql & " FROM DrugsaleItems"
    sql = sql & " JOIN Drug"
    sql = sql & " ON Drug.DrugID = DrugsaleItems.DrugID"
    sql = sql & " WHERE DispenseDate"
    If (periodStart <> "" And periodEnd <> "") Then
    sql = sql & " BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    Else
        sql = sql & " BETWEEN '2018-01-10'AND '2022-12-31'"
    End If
    sql = sql & " ),"
    sql = sql & " PrescriptionCTE"
    sql = sql & " AS"
    sql = sql & " ("
    sql = sql & " SELECT Drug.Drugname [PrescribedDrug],"
    sql = sql & " Prescription.PrescriptionDate,"
    sql = sql & " Prescription.VisitationID"
    sql = sql & " FROM Drug"
    sql = sql & " JOIN Prescription "
    sql = sql & " ON Drug.DrugID = Prescription.DrugID"
    sql = sql & " LEFT JOIN DrugsaleItems "
    sql = sql & " ON DrugsaleItems.VisitationID = Prescription.VisitationID "
    sql = sql & " WHERE PrescriptionDate"
    If (periodStart <> "" And periodEnd <> "") Then
    sql = sql & " BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    Else
        sql = sql & " BETWEEN '2018-01-10'AND '2022-12-31'"
    End If
    sql = sql & " ),"
    sql = sql & " NOTDISPENSEDCTE"
    sql = sql & " AS "
    sql = sql & " ("
    sql = sql & " SELECT PrescriptionCTE.PrescribedDrug, DispensedCTE.DispensedDrug, "
    sql = sql & " PrescriptionCTE.PrescriptionDate, DispensedCTE.DispenseDate, "
    sql = sql & " PrescriptionCTE.VisitationID [PresVstID], DispensedCTE.VisitationID [DisVstID]"
    sql = sql & " FROM PrescriptionCTE "
    sql = sql & " LEFT JOIN DispensedCTE"
    sql = sql & " ON PrescriptionCTE.VisitationID = DispensedCTE.VisitationID)"
    sql = sql & " SELECT Count(*) NumberOfPrescriptions,PrescribedDrug  FROM NOTDISPENSEDCTE"
    sql = sql & " WHERE DispensedDrug IS NULL"
    sql = sql & " GROUP BY PrescribedDrug"
    sql = sql & " ORDER BY NumberOfPrescriptions DESC, PrescribedDrug;"
    
    'response.write sql
    response.write datePeriod
    
    With rst
        .open sql, conn, 3, 4
            response.write "<form id='dateForm'> "
            response.write "    <div class='container' style='display: flex;'> "
            response.write "        <div> "
            response.write "            <label for='from'>From</label> "
            response.write "            <input type='date' name='from' id='from'> "
            response.write "        </div> "
            response.write "        <div> "
            response.write "            <label for='to'>To</label> "
            response.write "            <input type='date' name='to' id='to'> "
            response.write "        </div> "
            response.write "        <div> "
            response.write "            <button type='button' onclick='updateUrl()'>Show Data</button> "
            response.write "        </div>    "
            response.write "    </div> "
            response.write "</form> "
            response.write "<script> "
            response.write "    function updateUrl() { "
            response.write "        const fromDate = document.getElementById('from').value; "
            response.write "        const toDate = document.getElementById('to').value; "
            response.write "        const baseUrl = 'http://192.168.5.11/rmchms01/wpgPrtPrintLayoutAll.asp'; "
            response.write "        const params = new URLSearchParams({ "
            response.write "            PrintLayoutName: 'displayDrugs', "
            response.write "            PositionForTableName: 'WorkingDay', "
            response.write "            WorkingDayID: '' ,"
            response.write "            Dateperiod: fromDate + '||' + toDate"
            response.write "        }); "
            response.write "        const newUrl = baseUrl + '?' + params.toString(); "
            response.write "        console.log(newUrl); "
            response.write "        window.location.href = newUrl; "
            response.write "    } "
            response.write "</script> "
        If .RecordCount > 0 Then
           
            .MoveFirst
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr>"
                response.write "<th> Prescribed Drug </th>"
                response.write "<th> Number Of Prescriptions </th>"
                response.write "</tr>"
            Do While Not .EOF
                response.write "<tr>"
                response.write "<td>" & .fields("PrescribedDrug") & "</td>"
                response.write "<td>" & .fields("NumberOfPrescriptions") & "</td>"
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
        response.write "     border-collapse: collapse;"
        response.write " }"
        response.write " th, td {"
        response.write "     border: 1px solid #ddd;"
        response.write "     padding: 8px;"
        response.write " }"
        response.write " th {"
        response.write "     background-color: #f2f2f2;"
        response.write "     color: black;"
        response.write "     text-align: left;"
        response.write " }"
        response.write "     button {"
        response.write "         background-color: #0236c4;"
        response.write "         border-radius: 5px;"
        response.write "         border: none;"
        response.write "         margin-left: 50px;"
        response.write "         padding: 5px 20px;"
        response.write "         color: white;"
        response.write "     }"
        response.write "     #to, #from {"
        response.write "         margin: 20px;"
        response.write "     }"
        response.write "     .container {"
        response.write "         align-items: center;"
        response.write "     }"
        response.write "     button { "
        response.write "        cursor: pointer;"
        response.write "     } "
    response.write " </style>"

End Sub



'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
'
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
