'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim rptGen, dateRange, args, consultationtype, medicalStaff, treatment, sponsor
dateRange = Split(Trim(Request.QueryString("PrintFilter0")), "||")

consultationtype = Trim(Request.QueryString("PrintFilter1"))
medicalStaff = Trim(Request.QueryString("PrintFilter2"))
treatment = Trim(Request.QueryString("PrintFilter3"))
sponsor = Trim(Request.QueryString("PrintFilter4"))

Set rptGen = New PRTGLO_RptGen

sql = GetProcedureGroups(dateRange)
args = "title=procedures done by type from " & dateRange(0) & " to " & dateRange(1) & ";"
rptGen.AddReport sql, args

sql = GetPatientProcedureLists(dateRange)
args = "title=procedures done from " & dateRange(0) & " to " & dateRange(1) & ";"
'args = args & "hiddenFields=patient name;"
rptGen.AddReport sql, args

rptGen.ShowReport

Function GetProcedureGroups(dateRange)
    Dim sql
    
    sql = "SELECT t.TreatmentID AS [Procedure Code]"
    sql = sql & " , Treatment.TreatmentName AS [Procedure Name]"
    sql = sql & " , TreatCategory.TreatCategoryName AS [Treatment Category]"
    sql = sql & " , COUNT(t.TreatmentID) AS [Number of cases]"
    sql = sql & " FROM TreatCharges AS t"
    sql = sql & " LEFT JOIN Visitation v ON t.VisitationID = v.VisitationID"
    sql = sql & " LEFT JOIN Treatment ON Treatment.TreatmentID=t.TreatmentID"
    sql = sql & " LEFT JOIN Jobschedule ON Jobschedule.JobscheduleID=t.JobscheduleID"
    sql = sql & " LEFT JOIN Department ON Department.DepartmentID=Jobschedule.DepartmentID"
    sql = sql & " LEFT JOIN TreatCategory ON TreatCategory.TreatCategoryID=t.TreatCategoryID "
    sql = sql & " WHERE 1=1 "
'    sql = sql & "   AND t.TreatmentID IN ('P177', 'OP045', 'CON107', 'CON106') "
    
    If IsArray(dateRange) Then
        If UBound(dateRange) > 0 Then
            sql = sql & "   AND t.ConsultReviewDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
        End If
    End If
    
    If consultationtype <> "" Then
        sql = sql & "AND v.SpecialistTypeID = '" & consultationtype & "' "
    Else
        sql = sql & " "
    End If
    
    If medicalStaff <> "" Then
        sql = sql & "AND t.MedicalStaffID = '" & medicalStaff & "' "
    Else
        sql = sql & " "
    End If
    
    If treatment <> "" Then
        sql = sql & " AND Treatment.TreatmentID = '" & treatment & "' "
    Else
        sql = sql & " "
    End If
    
    If sponsor <> "" Then
        sql = sql & "AND v.SponsorID = '" & sponsor & "' "
    Else
        sql = sql & " "
    End If
    
    sql = sql & " GROUP BY t.TreatmentID, Treatment.TreatmentName, TreatCategory.TreatCategoryName" ', Department.DepartmentName "
    sql = sql & " ORDER BY [Procedure Name]"
    GetProcedureGroups = sql
End Function

Function GetPatientProcedureLists(dateRange)
    Dim sql
    
    sql = "SELECT CAST(v.VisitDate AS DATE) AS [Date] "
    sql = sql & " , p.PatientName AS [Patient Name] "
    sql = sql & " , v.VisitInfo6 AS [Age] "
    sql = sql & " , gen.GenderName AS [Sex] "
    sql = sql & " , trt.TreatmentName AS [Type Of Procedure] "
    sql = sql & " , spn.SponsorName AS [Insurance Status] "
    sql = sql & " , Department.DepartmentName AS [Department] "
    sql = sql & " , TreatCategory.TreatCategoryName AS [Treatment Category] "
    sql = sql & " , md.MedicalStaffName AS [Name of Surgeon] "
    sql = sql & " FROM TreatCharges AS t INNER JOIN Visitation AS v ON v.VisitationID=t.VisitationID "
    sql = sql & "       LEFT JOIN Patient AS p ON p.PatientID=v.PatientID "
    sql = sql & "       LEFT JOIN Gender AS gen ON gen.GenderID=p.GenderID "
    sql = sql & "       LEFT JOIN Treatment AS trt ON trt.TreatmentID=t.TreatmentID "
    sql = sql & "       LEFT JOIN Sponsor AS spn ON spn.SponsorID=v.SponsorID "
    sql = sql & "       LEFT JOIN MedicalStaff AS md ON md.MedicalStaffID = t.MedicalStaffID "
    sql = sql & "       LEFT JOIN Jobschedule ON Jobschedule.JobscheduleID=t.JobscheduleID "
    sql = sql & "       LEFT JOIN Department ON Department.DepartmentID=Jobschedule.DepartmentID "
    sql = sql & "       LEFT JOIN TreatCategory ON TreatCategory.TreatCategoryID=trt.TreatCategoryID "
    sql = sql & " WHERE 1=1 "
'    sql = sql & "   AND t.TreatmentID IN ('P177', 'OP045', 'CON107', 'CON106') "
    sql = sql & "   AND t.ConsultReviewDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    
    If consultationtype <> "" Then
        sql = sql & "AND v.SpecialistTypeID = '" & consultationtype & "' "
    Else
        sql = sql & " "
    End If
    
    If medicalStaff <> "" Then
        sql = sql & "AND md.MedicalStaffID = '" & medicalStaff & "' "
    Else
        sql = sql & " "
    End If
    
    If treatment <> "" Then
        sql = sql & " AND trt.TreatmentID = '" & treatment & "' "
    Else
        sql = sql & " "
    End If
    
    If sponsor <> "" Then
        sql = sql & "AND spn.SponsorID = '" & sponsor & "' "
    Else
        sql = sql & " "
    End If
        
    sql = sql & " ORDER BY [Date] DESC "
    
    GetPatientProcedureLists = sql
End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
