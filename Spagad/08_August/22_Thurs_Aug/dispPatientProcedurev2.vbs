'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Response.write "Hello Joe"
dispPatientProcedure

Sub dispPatientProcedure()
    Dim sql, count, rst
    
    ' Construct SQL for main query
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
    
    ' Open recordset
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open sql, conn, 3, 3  ' Use adOpenStatic (3) and adLockReadOnly (3) for read-only data
    
    If Not rst.EOF Then
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
        
        Do While Not rst.EOF
            count = count + 1
            Response.write "<tr>"
                Response.write "<td>" & count & "</td>"
                Response.write "<td>" & rst.Fields("PatientId").Value & "</td>"
                Response.write "<td>" & rst.Fields("PatientName").Value & "</td>"
                Response.write "<td>" & rst.Fields("age").Value & "</td>"
                Response.write "<td>" & rst.Fields("GenderName").Value & "</td>"
                Response.write "<td>" & rst.Fields("ConsultReviewDate").Value & "</td>"
                Response.write "<td>" & rst.Fields("MedicalStaffName").Value & "</td>"
                Response.write "<td>" & rst.Fields("TreatCategoryName").Value & "</td>"
                Response.write "<td>" & rst.Fields("TreatmentName").Value & "</td>"
                Response.write "<td>" & rst.Fields("Quantity").Value & "</td>"
                Response.write "<td>" & rst.Fields("UnitCost").Value & "</td>"
                Response.write "<td>" & rst.Fields("FinalAmount").Value & "</td>"
            Response.write "</tr>"
            rst.MoveNext  ' Move to the next record
        Loop
        
        Response.write "</table>"
    Else
        Response.write "No records found."
    End If
    
    ' Close recordset and cleanup
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
