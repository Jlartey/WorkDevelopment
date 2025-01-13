'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Response.write "Hello Joe"
dispPatientProcedure

Sub dispPatientProcedure()
    Dim sql, count
    
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
    sql = sql & "WHERE convert(date, TreatCharges.ConsultReviewDate) = convert(date, GETDATE())"
    
    Set rst = CreateObject("ADODB.Recordset")
    rst.open sql, conn, 3, 4
    
  
    Response.write "<table class='mytable'>"
    Response.write "<tr>"
        Response.write "<th class='myth'>No.</th>"
        Response.write "<th class='myth'>PatientID</th>"
        Response.write "<th class='myth'>PatientName</th>"
        Response.write "<th class='myth'>Age</th>"
        Response.write "<th class='myth'>Sex</th>"
        Response.write "<th class='myth'>Date</th>"
        Response.write "<th class='myth'>Doctor</th>"
        Response.write "<th class='myth'>Type</th>"
        Response.write "<th class='myth'>Item Name</th>"
        Response.write "<th class='myth'>Quantity</th>"
        Response.write "<th class='myth'>Unit Cost</th>"
        Response.write "<th class='myth'>Final Amount</th>"
    Response.write "</tr>"
    
    count = 0
    
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                count = count + 1
                Response.write "<tr>"
                    Response.write "<td>" & count & "</td>"
                    Response.write "<td>" & .fields("PatientId") & "</td>"
                    Response.write "<td>" & .fields("PatientName") & "</td>"
                    Response.write "<td>" & .fields("age") & "</td>"
                    Response.write "<td>" & .fields("GenderName") & "</td>"
                    Response.write "<td>" & .fields("ConsultReviewDate") & "</td>"
                    Response.write "<td>" & .fields("MedicalStaffName") & "</td>"
                    Response.write "<td>" & .fields("TreatCategoryName") & "</td>"
                    Response.write "<td>" & .fields("TreatmentName") & "</td>"
                    Response.write "<td>" & .fields("Quantity") & "</td>"
                    Response.write "<td>" & .fields("UnitCost") & "</td>"
                    Response.write "<td>" & .fields("FinalAmount") & "</td>"
                Response.write "</tr>"
              .MoveNext
            Loop
        End If
    End With
    
    Response.write "</table>"
    
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
