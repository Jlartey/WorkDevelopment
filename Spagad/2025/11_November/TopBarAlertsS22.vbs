topBarAlertsS22

Sub topBarAlertsS22()
    Dim ot
    Set ot = CreateObject("System.Collections.ArrayList")
    
    Call GetNewPrescriptionsAll(60, ot)
    
    If IsObject(response) Then
        response.Clear
        response.ContentType = "application/json"
        response.write gUtils.JSONStringify(ot)
    End If
    
End Sub
Function GetNewPrescriptionsAll(nMinBack, ByRef ot)
    Dim cnt, vst, sql, rst
    
    Set rst = CreateObject("ADODB.RecordSet")
    If Not IsNumeric(nMinBack) Then
        nMinBack = 60
    End If
    
    sql = "select Prescription.PatientID, Patient.PatientName, Prescription.VisitationID, max(Prescription.PrescriptionDate) as [date]"
    sql = sql & "  , count(distinct Prescription.DrugID) as [total]"
    sql = sql & "  , max(Prescription.PrescribeAmt2) as [emerg]"
    sql = sql & "  , (" & nMinBack & " - datediff(minute, max(Prescription.PrescriptionDate), getdate()) ) as [nMins]"
    sql = sql & " from Prescription"
    sql = sql & " inner join Patient on Patient.PatientID=Prescription.PatientID "
    sql = sql & "   and datediff(minute, Prescription.PrescriptionDate, getdate())<=" & nMinBack & " "
    sql = sql & "   and Prescription.PrescriptionStatusID='P001' "
    sql = sql & " inner join Visitation on Visitation.VisitationID=Prescription.VisitationID"
    'sql = sql & "   and Visitation.MedicalServiceID in ('M001', 'M002')" 'outpatient
    sql = sql & " group by Prescription.PatientID, Patient.PatientName, Prescription.VisitationID "
    sql = sql & " order by max(PrescriptionDate) desc"
    
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation&FullScreen=No"
            href = href & "&VisitationID=" & rst.fields("VisitationID")
            
            Set tmp = CreateObject("Scripting.Dictionary")
            tmp.Add "url", href
            If rst.fields("emerg") > 0 Then
                tmp.Add "class", "danger"
            End If
            tmp.Add "title", "New Prescription for " & rst.fields("PatientName")
            tmp.Add "message", "This prescription may leave the notifications in " & rst.fields("nMins") & " mins."
            tmp.Add "count", rst.fields("total").value
            
            ot.Add tmp
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
End Function
