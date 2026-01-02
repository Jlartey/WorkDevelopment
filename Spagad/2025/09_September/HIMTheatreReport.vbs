'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim sql, args, rptGen, dateRange, docID

dateRange = Split(Trim(Request.QueryString("PrintFilter0")), "||")
docID = Trim(Request.QueryString("PrintFilter1"))
Set rptGen = New PRTGLO_RptGen

sql = GetDoctorsOperationSummary(dateRange, docID)
args = "title=theatre report from " & dateRange(0) & " to " & dateRange(1) & ""
args = args & ";fieldFunctions=Operation Performed:getPerformedOperations"
'args = args & ";hiddenFields=findings|Patient Name"
rptGen.AddReport sql, args

rptGen.ShowReport

Function GetDoctorsOperationSummary(dateRange, docID)
    Dim sql
    
    sql = "SELECT CAST(v.VisitDate AS VARCHAR) AS [Date]"
    sql = sql & "   , p.PatientName AS [Patient Name] "
    sql = sql & "   , v.VisitInfo6 AS [Age] "
    sql = sql & "   , gen.GenderName AS [Sex] "
    sql = sql & "   , r2.Column2 AS [Operation Performed] "
    sql = sql & "   , (SELECT EMRVar3BName FROM EMRVar3B WHERE EMRVar3BID=CAST(r1.Column2 AS VARCHAR)) AS [Classification of Operation]"
    sql = sql & "   , (CAST(ISNULL(r3.Column4,'') AS NVARCHAR(MAX)) + ' ' + CAST(ISNULL(r4.Column1, '') AS NVARCHAR(MAX))) AS [Findings]"
    sql = sql & "   , (CAST(ISNULL(r5.Column6,'') AS NVARCHAR) + ' ' + CAST(ISNULL(EMRVar3B.EMRVar3BName, '') AS NVARCHAR )) AS [Anaesthesia given]"
    sql = sql & "   , wd.WardName AS [Ward transferred to]"
    sql = sql & "   , (SELECT StaffName FROM Staff WHERE Staff.StaffID=(SELECT StaffID FROM SystemUser WHERE e.SystemUserID=SystemUser.SystemUserID)) AS [Name Of Surgeon] "
    sql = sql & "   , (SELECT SponsorName FROM Sponsor WHERE Sponsor.SponsorID=v.SponsorID) AS [Insurance Status] "
    sql = sql & "   FROM "
    sql = sql & "       Visitation AS v INNER JOIN Patient AS p ON p.PatientID=v.PatientID "
    sql = sql & "       LEFT JOIN Gender AS gen ON gen.GenderID=p.GenderID "
    sql = sql & "       INNER JOIN EMRRequestItems AS e ON e.VisitationID=v.VisitationID "
    sql = sql & "           AND e.EMRDataID IN ('IM053', 'TH018' )"
    sql = sql & "    AND e.EMRDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "       LEFT JOIN EMRResults AS r1 ON r1.EMRRequestID=e.EMRRequestID AND r1.EMRComponentID IN ('IM05306') "
    sql = sql & "       LEFT JOIN EMRResults AS r2 ON r2.EMRRequestID=e.EMRRequestID AND r2.EMRComponentID IN ('IM05301', 'TH01808' ) "
    sql = sql & "       LEFT JOIN EMRResults AS r3 ON r3.EMRRequestID=e.EMRRequestID AND r3.EMRComponentID IN ('IM05301') "
    sql = sql & "       LEFT JOIN EMRResults AS r4 ON r4.EMRRequestID=e.EMRRequestID AND r4.EMRComponentID IN ('TH01810' ) "
    sql = sql & "       LEFT JOIN EMRResults AS r5 ON r5.EMRRequestID=e.EMRRequestID AND r5.EMRComponentID IN ('IM05301' ) "
    sql = sql & "       LEFT JOIN EMRResults AS r6 ON r6.EMRRequestID=e.EMRRequestID AND r6.EMRComponentID IN ('TH01804' ) "
    sql = sql & "       LEFT JOIN EMRVar3B ON EMRVar3B.EMRVar3BID=CAST(r6.Column4 AS VARCHAR) "
    sql = sql & "       LEFT JOIN EMRResults AS w ON w.EMRRequestID=e.EMRRequestID AND w.EMRComponentID IN ('IM05302' ) "
    sql = sql & "       LEFT JOIN Ward AS wd ON wd.WardID=CAST(w.Column2 AS VARCHAR) "
    
    sql = sql & " ORDER BY v.VisitDate ASC "
    
    GetDoctorsOperationSummary = sql
End Function
Function GetPerformedOperations(RECOBJ, fieldNAme)
    Dim ot, res, sql, catID
    
    res = Array()
    For Each cat In Split(RECOBJ(fieldNAme), "||")
        If catID <> "" Then catID = catID & ", "
        If Trim(cat) <> "" Then catID = catID & " '" & Replace(cat, "'", "''") & "' "
    Next
    If catID <> "" Then
        sql = "SELECT AppointmentCatTypeName FROM AppointmentCatType WHERE AppointmentCatTypeID IN ( " & catID & ")  "
        res = GetQueryResultsArray(sql)
        For Each oper In res
            If ot <> "" Then ot = ot & ", "
            ot = ot & oper
        Next
    End If
    
    If UBound(res) < 0 Then ot = RECOBJ(fieldNAme)
    GetPerformedOperations = ot
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
