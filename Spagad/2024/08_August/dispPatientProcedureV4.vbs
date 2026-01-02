'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim patient, doctor, periodStart, periodEnd

patient = Trim(Request.querystring("PrintFilter"))
doctor = Trim(Request.querystring("PrintFilter1"))
datePeriod = Trim(Request.querystring("PrintFilter2"))

If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    periodStart = dateArr(0)
    periodEnd = dateArr(1)
End If

tableStyles
dispPatientProcedure

Sub dispPatientProcedure()
    Dim sql, count
    
    Set rst = CreateObject("ADODB.Recordset")
    
    ' Construct SQL for main query
    sql = "SELECT "
    sql = sql & "TreatCharges.PatientId, "
    sql = sql & "Patient.PatientName, "
    sql = sql & "Patient.age, "
    sql = sql & "Gender.GenderName, "
    sql = sql & "convert(varchar(20), TreatCharges.ConsultReviewDate, 106) ConsultReviewDate, "
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
    
    If doctor <> "" Then
        sql = sql & "AND MedicalStaff.MedicalStaffID = '" & doctor & "' "
    Else
        sql = sql & "AND MedicalStaff.MedicalStaffID IS NOT NULL "
    End If
    
    If patient <> "" Then
        sql = sql & "AND Patient.PatientID = '" & patient & "'"
    Else
        sql = sql & "AND Patient.PatientID  IS NOT NULL "
    End If
    
'    response.write sql
'    response.write periodStart
'    response.write "<br>"
'    response.write periodEnd
'    response.write "<br>"

    count = 0
    
    With rst
        .open sql, conn, 3, 4
        If .RecordCount > 0 Then
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
                response.write "<th class='myth'>No.</th>"
                response.write "<th class='myth'>Patient ID</th>"
                response.write "<th class='myth'>Patient Name</th>"
                response.write "<th class='myth'>Age</th>"
                response.write "<th class='myth'>Sex</th>"
                response.write "<th class='myth'>Date</th>"
                response.write "<th class='myth'>Doctor</th>"
                response.write "<th class='myth'>Type</th>"
                response.write "<th class='myth' style='width: 50px;'>Item Name</th>"
                response.write "<th class='myth'>Quantity</th>"
                response.write "<th class='myth'>Unit Cost</th>"
                response.write "<th class='myth'>Final Amount</th>"
            response.write "</tr class='mytr'>"
            
            .movefirst
            Do While Not .EOF
                count = count + 1
                response.write "<tr class='mytr'>"
                    response.write "<td class='mytd'>" & count & "</td>"
                    response.write "<td class='mytd'>" & .fields("PatientId") & "</td>"
                    response.write "<td class='mytd'>" & .fields("PatientName") & "</td>"
                    response.write "<td class='mytd'>" & .fields("age") & "</td>"
                    response.write "<td class='mytd'>" & .fields("GenderName") & "</td>"
                    response.write "<td class='mytd'>" & .fields("ConsultReviewDate") & "</td>"
                    response.write "<td class='mytd'>" & .fields("MedicalStaffName") & "</td>"
                    response.write "<td class='mytd'>" & .fields("TreatCategoryName") & "</td>"
                    response.write "<td class='mytd'>" & .fields("TreatmentName") & "</td>"
                    response.write "<td class='mytd'>" & .fields("Quantity") & "</td>"
                    response.write "<td class='mytd'>" & .fields("UnitCost") & "</td>"
                    response.write "<td class='mytd'>" & .fields("FinalAmount") & "</td>"
                response.write "</tr class='mytr'>"
                .MoveNext
            Loop
            response.write "</table>"
        Else
            response.write "<h1>No records found for the given filters.</h1>"
        End If
    End With
    
    rst.Close
    Set rst = Nothing
End Sub

Sub tableStyles()
    response.write "<style>"
        response.write ".mytable {"
        response.write "    width: 85vw;"
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
