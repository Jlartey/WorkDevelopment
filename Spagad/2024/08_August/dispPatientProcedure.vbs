'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

'Created on 28th August, 2024 by Joe Lartey

Dim periodStart, periodEnd, datePeriod, dateArr
datePeriod = Trim(request.querystring("Dateperiod"))
If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    periodStart = dateArr(0)
    periodEnd = dateArr(1)
Else
    periodStart = "2024-08-20"
    periodEnd = "2024-08-28"
End If

dispPatientProcedure

Sub dispPatientProcedure()
    Dim sql, count
     
    sql = " SELECT MedicalStaffId, MedicalStaffName FROM MedicalStaff"
    
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.open sql, conn, 3, 4
    
    dropdownOptions = ""
    
    With rstDropdown
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                optionHTML = "<option value='" & .fields("MedicalStaffId") & "'>" & .fields("MedicalStaffName")& "</option>"
            Loop
        End If
    End With
    
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

    'response.write sql
    
    response.write periodStart
    response.write "<br> "
    response.write periodEnd

    ' Output dropdown
    Response.Write "<div>"
    Response.Write "        <label for='doctor' class='font-style'>Select Doctor:</label><br>"
    Response.Write "        <select id='doctor' name='doctor' multiple class='mult-select-tag'>"
    Response.Write dropdownOptions
    Response.Write "        </select>"
    Response.Write "</div>"
    
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
    response.write "        const baseUrl = 'http://172.2.2.31/hms/wpgPrtPrintLayoutAll.asp'; "
    response.write "        const params = new URLSearchParams({ "
    response.write "            PrintLayoutName: 'dispPatientProcedure', "
    response.write "            PositionForTableName: 'WorkingDay', "
    response.write "            WorkingDayID: '' ,"
    response.write "            Dateperiod: fromDate + '||' + toDate"
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

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
