'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
'Response.Write "Hello Joe"
Dim patient, doctor, periodStart, periodEnd

patient = Trim(Request.querystring("PrintFilter"))
doctor = Trim(Request.querystring("PrintFilter1"))
datePeriod = Trim(Request.querystring("PrintFilter2"))

response.write patient
response.write doctor
response.write datePeriod

If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    periodStart = dateArr(0)
    periodEnd = dateArr(1)
End If

dispPatientProcedure

Sub dispPatientProcedure()
    Dim sql, count
    
    
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
    sql = sql & "AND MedicalStaff.MedicalStaffID = '" & doctor & "' "
    sql = sql & "AND Patient.PatientID = '" & patient & "'"
    
    
    response.write sql
    
    response.write periodStart
    response.write "<br> "
    response.write periodEnd
    
    response.write " <br />"
    
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
            .movefirst
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
