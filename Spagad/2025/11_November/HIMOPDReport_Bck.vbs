'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim rGenObj, dateRange
Dim args, sql, ageGroup, spType

ShowReport
Sub ShowReport()
    Set rGenObj = New PRTGLO_RptGen
    dateRange = Split(Trim(Request.QueryString("PrintFilter0")), "||")
    ageGroup = Trim(Request.QueryString("PrintFilter1"))
    spType = Trim(Request.QueryString("PrintFilter2"))
    sql = ConsultationReportSQL(dateRange, ageGroup, spType)
    args = "title=consultations from " & dateRange(0) & " to " & dateRange(1)
    args = args & ";fieldFunctions=diagnosis:getpatientdiagnosis"
    args = args & "|investigations Requested:getpatientInvestigations|medications:getPatientMedications;"
    'args = args & ";hidenFields=patient name"
    rGenObj.AddReport sql, args
    
    'sql = MorbidityReportSQL(dateRange)
    'args = "title=morbidity;"
    'rGenObj.AddReport sql, args
    '
    'sql = ConsultationReportSQL(dateRange)
    'args = "title=statement of out-patients;"
    'rGenObj.AddReport sql, args
    
    sql = ReferralInReportSQL(dateRange)
    args = "title=referral in from " & dateRange(0) & " to " & dateRange(1)
    args = args & ";fieldfunctions=type of referral:HandleReferral"
    args = args & "|referral from:HandleReferral"
    args = args & "|referring diagnosis:HandleReferral"
    args = args & "|other diagnosis:HandleReferral"
    args = args & "|specialty referring to:HandleReferral"
    'args = args & ";hidenFields=patient name"
    rGenObj.AddReport sql, args
    
    sql = GenericOPDReport(dateRange, ageGroup, spType)
    args = "title=Report (NEW) from " & dateRange(0) & " to " & dateRange(1)
    args = args & ";fieldFunctions=Cashier:GetCashiers|Nurses:GetNurses|Lab staff:GetLabStaff|Radiology staff:GetRadStaff|Pharmacy staff:GetPharmStaff|Dialysis staff:GetDialStaff|Mortuary staff:GetMortStaff"
    rGenObj.AddReport sql, args
    
    rGenObj.ShowReport
End Sub
Function ConsultationReportSQL(dateRange, ageGroup, spType)
    Dim sql
    
   sql = "WITH DoctorConsults AS ( "
    sql = sql & "SELECT DISTINCT SystemUserID, VisitationID "
    sql = sql & "From EMRRequestItems "
    sql = sql & "WHERE EMRDataID IN ('TH060', 'IM051') "
    sql = sql & ") "
    sql = sql & "SELECT DISTINCT sp.SpecialistTypeName AS [Consultation Type], "
    sql = sql & "CAST(CONVERT(DATE, vst.VisitDate) AS VARCHAR) AS [Date of Consultation], "
    sql = sql & "Patient.PatientID AS [Patient ID], "
    sql = sql & "Patient.PatientName AS [Patient Name], "
    sql = sql & "vst.PatientAge AS [Age], gen.GenderName AS [Sex], "
    sql = sql & "VisitType.VisitTypeName AS [First/Follow Up], "
    sql = sql & "vst.VisitationID AS [Diagnosis], "
    sql = sql & "(SELECT TOP 1 dg.DiagnosisType "
    sql = sql & " FROM Diagnosis AS dg "
    sql = sql & " WHERE dg.VisitationID = vst.VisitationID "
    sql = sql & " ORDER BY dg.ConsultReviewID) AS [Diagnosis Type], "
    sql = sql & "(SELECT TOP 1 temp.temperature "  '
    sql = sql & " FROM ( "
    sql = sql & "    SELECT DISTINCT visitationid, CONVERT(NVARCHAR(MAX), Column3) AS Temperature "
    sql = sql & "    FROM EMRResults AS emrres "
    sql = sql & "    JOIN EMRRequest emrreq ON emrres.EMRRequestID = emrreq.EMRRequestID "
    sql = sql & "    WHERE emrdataid = 'EMR034' AND EMRComponentID = 'EMR03410' "
    sql = sql & " ) AS temp "
    sql = sql & " WHERE temp.visitationid = vst.VisitationID "
    sql = sql & " ORDER BY temp.visitationid) AS [Temperature], "
    sql = sql & "(SELECT TOP 1 wg.weight12 "
    sql = sql & " FROM ( "
    sql = sql & "    SELECT DISTINCT visitationid, CONVERT(NVARCHAR(MAX), Column2) AS Weight12 "
    sql = sql & "    FROM EMRResults AS emrres "
    sql = sql & "    JOIN EMRRequest emrreq ON emrres.EMRRequestID = emrreq.EMRRequestID "
    sql = sql & "    WHERE emrdataid = 'EMR034' AND EMRComponentID = 'EMR05003' "
    sql = sql & " ) AS wg "
    sql = sql & " WHERE wg.visitationid = vst.VisitationID "
    sql = sql & " ORDER BY wg.visitationid) AS [Weight], "
    sql = sql & "(SELECT TOP 1 ht.height "
    sql = sql & " FROM ( "
    sql = sql & "    SELECT DISTINCT visitationid, CONVERT(NVARCHAR(MAX), Column4) AS height "
    sql = sql & "    FROM EMRResults AS emrres "
    sql = sql & "    JOIN EMRRequest emrreq ON emrres.EMRRequestID = emrreq.EMRRequestID "
    sql = sql & "    WHERE emrdataid = 'EMR034' AND EMRComponentID = 'EMR05004' "
    sql = sql & " ) AS ht "
    sql = sql & " WHERE ht.visitationid = vst.VisitationID "
    sql = sql & " ORDER BY ht.visitationid) AS [Height], "
    sql = sql & "vst.VisitationID AS [Investigations Requested], "
    sql = sql & "vst.VisitationID AS [Medications], mo.MedicalOutcomeName AS [Outcome], "
    sql = sql & "spn.SponsorName AS [Insurance Status], "
    sql = sql & " (SELECT  TOP 1 Staff.StaffName"
    sql = sql & " FROM DoctorConsults dc"
    sql = sql & " JOIN SystemUser ON SystemUser.SystemUserID = dc.SystemUserID"
    sql = sql & " JOIN Staff ON Staff.StaffID = SystemUser.StaffID"
    sql = sql & " WHERE dc.VisitationID = vst.VisitationID) AS [Attending Doctor],"
    sql = sql & "Patient.ResidencePhone AS [Patient Tel.] "
    sql = sql & "FROM Visitation AS vst "
    sql = sql & "LEFT JOIN Patient ON Patient.PatientID = vst.PatientID "
    sql = sql & "LEFT JOIN Gender AS gen ON gen.GenderID = vst.GenderID "
    sql = sql & "LEFT JOIN VisitType ON vst.VisitTypeID = VisitType.VisitTypeID "
    sql = sql & "LEFT JOIN SpecialistType AS sp ON sp.SpecialistTypeID = vst.SpecialistTypeID "
    sql = sql & "LEFT JOIN MedicalOutcome AS mo ON mo.MedicalOutcomeID = vst.MedicalOutcomeID "
    sql = sql & "LEFT JOIN SystemUser AS su ON su.SystemUserID = vst.SpecialistID "
    sql = sql & "LEFT JOIN Sponsor AS spn ON spn.SponsorID = vst.SponsorID "
    sql = sql & "WHERE vst.VisitDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "    AND vst.PatientID <> 'P3' "
    If ageGroup <> "" Then sql = sql & "    AND vst.AgeGroupID='" & ageGroup & "' "
    If spType <> "" Then sql = sql & "    AND vst.SpecialistTypeID='" & spType & "' "
    sql = sql & "ORDER BY CAST(CONVERT(DATE, vst.VisitDate) AS VARCHAR) ASC, Patient.PatientName ASC "

    ConsultationReportSQL = sql
End Function
Function MorbidityReportSQL(dateRange)
    Dim sql
    sql = "SELECT sp.SpecialistType AS [Consultation Type], vst.VisitDate AS [Date of Consultation]"
    sql = sql & " , vst.VisitInfo6 AS [Age], gen.GenderName AS [Sex], vst.VisitationID AS [Diagnosis]"
    sql = sql & " , vst.VisitationID AS [Investigations Requested]"
    sql = sql & " , vst.VisitationID AS [Medications], mo.MedicalOutcome AS [Outcome]"
    sql = sql & " , spn.SponsorName AS [Insurance Status]"
    sql = sql & " , s.SpecialistName AS [Attending Doctor]"
    
    sql = sql & " FROM Visitation AS vst LEFT JOIN Patient AS p ON p.PatientID=vst.PatientID"
    sql = sql & "   LEFT JOIN Gender ON gen.GenderID=vst.GenderID "
    sql = sql & "   LEFT JOIN SpecialistType AS sp ON sp.SpecialistTypeID=vst.SpecialistTypeID"
    sql = sql & "   LEFT JOIN MedicalOutcome AS mo ON mo.MedicalOutcomeID=vst.MedicalOutcomeID"
    sql = sql & "   LEFT JOIN SystemUser AS su ON su.SystemUserID=vst.SpecialistID"
    sql = sql & "   LEFT JOIN Specialist AS s ON s.SpecialistID=vst.SpecialistID"
    sql = sql & "   LEFT JOIN Sponsor AS spn ON spn.SponsorID=vst.SponsorID "
    
    sql = sql & " WHERE vst.VisitDate BETWEEN '" & dateRange(1) & "' AND '" & dateRange(0) & "' "
    sql = sql & "   AND vst.PatientID <> 'P3' "
    sql = sql & " ORDER BY vst.VisitDate ASC "
    MorbidityReportSQL = sql
End Function
Function ReferralInReportSQL(dateRange)
    Dim sql
    sql = "SELECT CAST(vst.VisitDate AS VARCHAR) AS [Date of visit]"
    sql = sql & " , p.PatientName AS [Patient Name]"
    sql = sql & " , r.Column1 AS [Type of Referral]"
    sql = sql & " , r.Column1 AS [Referral From]"
    sql = sql & " , r.Column1 AS [Referring Diagnosis]"
    sql = sql & " , r.Column1 AS [Other Diagnosis]"
    sql = sql & " , r.Column1 AS [Specialty referring to]"
    sql = sql & " FROM EMRRequestItems AS e"
    sql = sql & " LEFT JOIN EMRResults AS r ON r.EMRRequestID=e.EMRRequestID AND r.EMRComponentID='EMR05006'"
    sql = sql & " LEFT JOIN Visitation AS vst ON vst.VisitationID=e.VisitationID "
    sql = sql & " LEFT JOIN Patient AS p ON p.PatientID=vst.PatientID"
    'sql = sql & " LEFT JOIN PatientReferral AS pr ON pr.VisitationID=vst.VisitationID "
    sql = sql & " WHERE vst.VisitDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "'"
    sql = sql & "   AND LEN(CAST(r.Column1 AS VARCHAR)) > 0 "
    sql = sql & "   AND vst.PatientID <> 'P3' "
    sql = sql & " ORDER BY vst.VisitDate ASC "
    'sql = sql & " UNION "
    'sql = sql & "   "
    ReferralInReportSQL = sql
End Function
Function HandleReferral(rowObj, fieldNAme)
    Dim ot, val, cmp, rst
    ot = ""

    val = rowObj(fieldNAme)
    
    If UCase(fieldNAme) = "TYPE OF REFERRAL" Then
        cmp = "EMRVar2B-EMR050060100101"
        ot = ExtractCheckDetCmp(val, cmp, "Column2", 0)
        ot = GetComboName("EMRVar3B", Trim(ot))
    ElseIf UCase(fieldNAme) = "REFERRING DIAGNOSIS" Then
        cmp = "EMRVar2B-EMR050060100102"
        ot = ExtractCheckDetCmp(val, cmp, "Column2", 0)
        arr = Split(ot, "~~")
        ot = ""
        For Each val In arr
            kv = Split(val, "||")
            If ot <> "" Then ot = ot & ", "
            ot = ot & GetComboName("Disease", kv(0))
        Next
    ElseIf UCase(fieldNAme) = "REFERRAL FROM" Then
        cmp = "EMRVar2B-EMR050060100101"
        ot = ExtractCheckDetCmp(val, cmp, "Column5", 0)
    ElseIf UCase(fieldNAme) = "OTHER DIAGNOSIS" Then
        cmp = "EMRVar2B-EMR050060100104"
        ot = ExtractCheckDetCmp(val, cmp, "Column2", 0)
    ElseIf UCase(fieldNAme) = "SPECIALTY REFERRING TO" Then
        cmp = "EMRVar2B-EMR050060100103"
        ot = ExtractCheckDetCmp(val, cmp, "Column2", 0)
        ot = GetComboName("SpecialistGroup", Trim(ot))
        ot2 = Trim(ExtractCheckDetCmp(val, cmp, "Column5", 0))
        If ot2 <> "" Then
            If ot <> "" Then ot = ot & ", "
            ot = ot & ot2
        End If
    End If
    HandleReferral = ot
End Function
Function GenericOPDReport(dateRange, ageGroup, spType)
    Dim sql
    
    sql = "select Visitation.VisitDate as [Date/time], Patient.PatientName as [Patient Name], SpecialistType.SpecialistTypeName as [Scheduled visit]"
    sql = sql & " , Staff.StaffName as [Records Dpt. Staff], Visitation.VisitationID as [Cashier], Visitation.VisitationID as [Nurses]"
    sql = sql & " , Specialist.SpecialistName as [Assigned doctor/consulting room], Visitation.VisitationID as [Lab staff]"
    sql = sql & " , Visitation.VisitationID as [Radiology staff], Visitation.VisitationID as [Pharmacy staff]"
    sql = sql & " , Visitation.VisitationID as [Dialysis staff], Visitation.VisitationID as [Mortuary staff]"
    sql = sql & " from Visitation"
    sql = sql & " left join Patient on Patient.PatientID=Visitation.PatientID"
    sql = sql & " left join SpecialistType on SpecialistType.SpecialistTypeID=Visitation.SpecialistTypeID"
    sql = sql & " left join Specialist on Specialist.SpecialistID=Visitation.SpecialistID"
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=Visitation.SystemUserID"
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID"
    sql = sql & " where 1=1"
    sql = sql & "  and Visitation.VisitDate between '" & dateRange(0) & "' and '" & dateRange(1) & "' "
    If ageGroup <> "" Then
        sql = sql & " and Visitation.AgeGroupID='" & ageGroup & "' "
    End If
    If spType <> "" Then
        sql = sql & " and Visitation.SpecialistTypeID='" & spType & "' "
    End If
    
    GenericOPDReport = sql
End Function
Function GetCashiers(RECOBJ, fieldNAme)
    Dim sql, rst, vst, ot
    ot = ""
    vst = RECOBJ(fieldNAme)
    
    sql = "select distinct Staff.StaffName from PatientReceipt2"
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=PatientReceipt2.SystemUserID"
    sql = sql & " left join Receipt on Receipt.ReceiptID=PatientReceipt2.ReceiptID"
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID"
    sql = sql & " where PatientReceipt2.VisitationID='" & vst & "' "
    Set rst = CreateObject("ADODB.RecordSet")
    
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If ot <> "" Then ot = ot & ", "
            ot = ot & rst.fields("StaffName")
            rst.MoveNext
        Loop
        rst.Close
    End If
    Set rst = Nothing
    GetCashiers = ot
End Function
Function GetNurses(RECOBJ, fieldNAme)
    Dim sql, rst, vst, ot
    ot = ""
    vst = RECOBJ(fieldNAme)
    
    sql = "select distinct Staff.StaffName from EMRRequestItems"
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=EMRRequestItems.SystemUserID"
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID"
    sql = sql & " where VisitationID='" & vst & "' and EMRRequestItems.JobscheduleID<> 'M0325' "
    Set rst = CreateObject("ADODB.RecordSet")
    
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If ot <> "" Then ot = ot & ", "
            ot = ot & rst.fields("StaffName")
            rst.MoveNext
        Loop
        rst.Close
    End If
    Set rst = Nothing
    GetNurses = ot
End Function
Function GetLabStaff(RECOBJ, fieldNAme)
    Dim sql, rst, vst, ot
    
    ot = ""
    vst = RECOBJ(fieldNAme)
    
    sql = "select Staff.StaffName from Investigation "
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=Investigation.LabTechID"
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID"
    sql = sql & " where VisitationID='" & vst & "' and Investigation.JobScheduleID in ('s13', 'dpt005')"
    sql = sql & " union "
    sql = sql & " select Staff.StaffName from Investigation2 "
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=Investigation2.LabTechID"
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID"
    sql = sql & " where VisitationID='" & vst & "' and Investigation2.JobScheduleID in ('s13', 'dpt005')"
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If ot <> "" Then ot = ot & ", "
            ot = ot & rst.fields("StaffName")
            rst.MoveNext
        Loop
        rst.Close
    End If
    GetLabStaff = ot
End Function
Function GetRadStaff(RECOBJ, fieldNAme)
    Dim sql, rst, vst, ot
    
    ot = ""
    vst = RECOBJ(fieldNAme)
    
    sql = "select Staff.StaffName from Investigation "
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=Investigation.LabTechID"
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID"
    sql = sql & " where VisitationID='" & vst & "' and Investigation.JobScheduleID in ('s19', 'dpt011')"
    sql = sql & " union "
    sql = sql & " select Staff.StaffName from Investigation2 "
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=Investigation2.LabTechID"
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID"
    sql = sql & " where VisitationID='" & vst & "' and Investigation2.JobScheduleID in ('s19', 'dpt011')"
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If ot <> "" Then ot = ot & ", "
            ot = ot & rst.fields("StaffName")
            rst.MoveNext
        Loop
        rst.Close
    End If
    GetRadStaff = ot
End Function
Function GetPharmStaff(RECOBJ, fieldNAme)
    Dim sql, rst, vst, ot
    
    ot = ""
    vst = RECOBJ(fieldNAme)
    
    sql = "select Staff.StaffName from DrugSale "
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=DrugSale.SystemUserID"
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID"
    sql = sql & " where VisitationID='" & vst & "' and DrugSale.JobScheduleID in ('s22', 'm0601', 'm0602', 'm0603') "
    sql = sql & " union "
    sql = sql & " select Staff.StaffName from DrugReturn "
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=DrugReturn.SystemUserID"
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID"
    sql = sql & " where VisitationID='" & vst & "' and DrugReturn.JobScheduleID in ('s22', 'm0601', 'm0602', 'm0603') "
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If ot <> "" Then ot = ot & ", "
            ot = ot & rst.fields("StaffName")
            rst.MoveNext
        Loop
        rst.Close
    End If
    GetPharmStaff = ot
End Function
Function GetDialStaff(RECOBJ, fieldNAme)
    Dim sql, rst, vst, ot
    ot = ""
    vst = RECOBJ(fieldNAme)
    
    sql = "select distinct Staff.StaffName from EMRRequestItems"
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=EMRRequestItems.SystemUserID"
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID"
    sql = sql & " where VisitationID='" & vst & "' and EMRRequestItems.JobscheduleID= 'M0325' "
    Set rst = CreateObject("ADODB.RecordSet")
    
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If ot <> "" Then ot = ot & ", "
            ot = ot & rst.fields("StaffName")
            rst.MoveNext
        Loop
        rst.Close
    End If
    Set rst = Nothing
    GetDialStaff = ot
End Function
Function GetMortStaff(RECOBJ, fieldNAme)
    Dim sql, rst, vst, ot
    ot = ""
    vst = RECOBJ(fieldNAme)
    
    sql = "select distinct Staff.StaffName from Mortuary"
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=Mortuary.SystemUserID"
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID"
    sql = sql & " where VisitationID='" & vst & "' "
    Set rst = CreateObject("ADODB.RecordSet")
    
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If ot <> "" Then ot = ot & ", "
            ot = ot & rst.fields("StaffName")
            rst.MoveNext
        Loop
        rst.Close
    End If
    Set rst = Nothing
    GetMortStaff = ot
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
