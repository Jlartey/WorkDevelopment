'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim thispage, rptDict, selrpt, dateRange, reportTitle, PRLTGLO_K
Set PRLTGLO_K = New PRTGLO_kenGlobal
reportTitle = "Pharmacy Reports"
'''
'''
''' INITIALIZATION
InitScript
ProcessPage
'''
'''
''' UTILITY FUNCTIONS
'''
Sub InitScript()
    Set rptDict = CreateObject("Scripting.Dictionary")
    selrpt = Trim(Request.QueryString("selRpt"))
    
'    add your report title, and a function to be called when your report is selected here
'    rptdict.Add "Report Title", GetRef("FunctionOrSubName")
    rptDict.Add "Staff Activity Report", GetRef("StaffActivity")
    rptDict.Add "Dispenses By Sponsor Report", GetRef("DispenseBySponsor")
    rptDict.Add "Dispenses By Service Type Report", GetRef("DispenseByMedicalService")
    If UCase(jSchd) = UCase(uname) Then
        rptDict.Add "Drug Sale/Returns Report", GetRef("SaleReturn")
    End If
    rptDict.Add "Top OPD Pharm. Drugs Sold", GetRef("TopDispenseReportOPD")
    rptDict.Add "Top IPD Pharm. Drugs Sold", GetRef("TopDispenseReportIPD")
    rptDict.Add "Top Drugs Returned", GetRef("TopReturnsReport")
    rptDict.Add "Top Drugs Sold All Pharm.", GetRef("TopDispenseReport")
'   rptDict.Add "Monthly Returns Surgical Operations", GetRef("SurgicalOperations")
End Sub 'InitScript
'''
'''
Sub ProcessPage()
    Dim dateRange, rptArray
    dateRange = PRLTGLO_K.MakeDateRange(Request.QueryString("PrintFilter0"))
    rptArray = rptDict.Keys()
    response.write PRLTGLO_K.GetReportHeader(dateRange, "", reportTitle)
    response.write PRLTGLO_K.ReportNavBar(rptArray, dateRange, selrpt)
    
    ''''' your sub/function will be called when it's it turn
    
    If rptDict.Exists(Replace(selrpt, "_", " ")) Then
        Set func = rptDict(Replace(selrpt, "_", " "))
        dateRange = PRLTGLO_K.MakeDateRange(Request.QueryString("PrintFilter0"))
        func dateRange
    Else
        ''
    End If
    
End Sub 'ProcessPage
Function GetAllStaff(dept)
    Dim sql, str, rst
    str = ""
    
    sql = "SELECT StaffName, StaffID FROM Staff WHERE StaffID "
    sql = sql & " IN (SELECT StaffID FROM SystemUser WHERE SystemUser.JobScheduleID "
    sql = sql & " IN ( SELECT JobScheduleID FROM JobSchedule WHERE DepartmentID='" & dept & "' ) "
    sql = sql & " AND SystemUser.SystemUserID IN (SELECT SystemUserID FROM DrugSale UNION SELECT SystemUserID FROM DrugReturn) "
    sql = sql & " ) AND StaffID<>'STF001'"
    sql = sql & " ORDER BY StaffName "
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        Do While Not rst.EOF
            If str = "" Then
                str = rst.fields("StaffID") & "||" & rst.fields("StaffName")
            Else
                str = str & "|*|" & rst.fields("StaffID") & "||" & rst.fields("StaffName")
            End If
            rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing
        GetAllStaff = Split(str, "|*|")
    Else
        GetAllStaff = Array()
    End If
End Function 'GetAllStaff
Function SelectStaff(staffArr, dateRange, selrpt)
    Dim str, cnt
    cnt = 0
    str = str & vbCrLf & "<table class='nav'>"
        str = str & vbCrLf & "<tr><td>Select Staff</td><td><select id='sel-staff-list'>"
        str = str & vbCrLf & "<option value></option>"
        For Each staff In staffArr
            staff = Split(staff, "||")
            str = str & vbCrLf & "<option value='" & Replace(staff(0), " ", "_") & "'"
            If UCase(staff(0)) = UCase(Replace(selrpt, "_", " ")) Then
                str = str & vbCrLf & "   selected"
            End If
            str = str & vbCrLf & ">" & UCase(staff(1)) & "</option>"
            cnt = cnt + 1
        Next
        str = str & vbCrLf & "</select></td> "
        str = str & vbCrLf & "<td><input type='button' onclick='open_staff_report()' value='Go'></td>"
    str = str & vbCrLf & "</table> <br/>"
    str = str & vbCrLf & "<style>"
        str = str & vbCrLf & ".nav td {text-transform: uppercase;}"
    str = str & vbCrLf & "</style>"
    str = str & vbCrLf & "<script> "
        str = str & vbCrLf & " var cnt = " & cnt & ";"
        str = str & vbCrLf & " if(cnt== 0 ){ "
        str = str & vbCrLf & "    var sel = document.getElementById('sel-staff-list'); " & vbCrLf
        str = str & vbCrLf & "    if(sel) { "
        str = str & vbCrLf & "       sel.disabled = true; "
        str = str & vbCrLf & "    }"
        str = str & vbCrLf & " }"
        str = str & vbCrLf & " if(cnt== 1 ){ "
        str = str & vbCrLf & "    var sel = document.getElementById('sel-staff-list'); "
        str = str & vbCrLf & "    if(sel) { "
        str = str & vbCrLf & "       sel.selectedIndex = 1; "
        str = str & vbCrLf & "       if ('" & selrpt & "'.toLowerCase()!=sel.options[sel.selectedIndex].value.toLowerCase()) open_staff_report();"
        str = str & vbCrLf & "    }"
        str = str & vbCrLf & " }"
        str = str & vbCrLf & " function open_staff_report(){ "
        str = str & vbCrLf & "   var sel = document.getElementById('sel-staff-list'); "
        str = str & vbCrLf & "   if(sel){ "
        str = str & vbCrLf & "       var opt = sel.options[sel.selectedIndex].value; "
        str = str & vbCrLf & "       window.location = CreatePageURL(['staffID=' + opt, 'PrintFilter0=" & dateRange(0) & "||" & dateRange(1) & "']); "
        str = str & vbCrLf & "   } "
        str = str & vbCrLf & " } "
    str = str & vbCrLf & "</script> "
    SelectStaff = str
End Function 'SelectStaff
Function MakeDropDownLink(title, queryParam, listArr, dateRange, selrpt)
    Dim str, cnt
    cnt = 0
    str = str & vbCrLf & "<table class='nav'>"
        str = str & vbCrLf & "<tr><td>" & title & "</td><td><select id='sel-staff-list'>"
        str = str & vbCrLf & "<option value></option>"
        For Each lst In listArr
            lst = Split(lst, "||")
            str = str & vbCrLf & "<option value='" & Replace(lst(0), " ", "_") & "'"
            If UCase(lst(0)) = UCase(Replace(selrpt, "_", " ")) Then
                str = str & vbCrLf & "   selected"
            End If
            str = str & vbCrLf & ">" & UCase(lst(1)) & "</option>"
            cnt = cnt + 1
        Next
        str = str & vbCrLf & "</select></td> "
        str = str & vbCrLf & "<td><input type='button' onclick='open_staff_report()' value='Go'></td>"
    str = str & vbCrLf & "</table> <br/>"
    str = str & vbCrLf & "<style>"
        str = str & vbCrLf & ".nav td {text-transform: uppercase;}"
    str = str & vbCrLf & "</style>"
    str = str & vbCrLf & "<script> "
        str = str & vbCrLf & " var cnt = " & cnt & ";"
        str = str & vbCrLf & " if(cnt== 0 ){ "
        str = str & vbCrLf & "    var sel = document.getElementById('sel-staff-list'); " & vbCrLf
        str = str & vbCrLf & "    if(sel) { "
        str = str & vbCrLf & "       sel.disabled = true; "
        str = str & vbCrLf & "    }"
        str = str & vbCrLf & " }"
        str = str & vbCrLf & " if(cnt== 1 ){ "
        str = str & vbCrLf & "    var sel = document.getElementById('sel-staff-list'); "
        str = str & vbCrLf & "    if(sel) { "
        str = str & vbCrLf & "       sel.selectedIndex = 1; "
        str = str & vbCrLf & "       if ('" & selrpt & "'.toLowerCase()!=sel.options[sel.selectedIndex].value.toLowerCase()) open_staff_report();"
        str = str & vbCrLf & "    }"
        str = str & vbCrLf & " }"
        str = str & vbCrLf & " function open_staff_report(){ "
        str = str & vbCrLf & "   var sel = document.getElementById('sel-staff-list'); "
        str = str & vbCrLf & "   if(sel){ "
        str = str & vbCrLf & "       var opt = sel.options[sel.selectedIndex].value; "
        str = str & vbCrLf & "       window.location = CreatePageURL(['" & queryParam & "=' + opt, 'PrintFilter0=" & dateRange(0) & "||" & dateRange(1) & "']); "
        str = str & vbCrLf & "   } "
        str = str & vbCrLf & " } "
    str = str & vbCrLf & "</script> "
    MakeDropDownLink = str
End Function 'MakeDropDownLink
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub StaffActivity(dateRange)
    Dim sql, staffID, extParams
    staffID = Request.QueryString("StaffID")
    Set extParams = CreateObject("Scripting.Dictionary")
    
    extParams.Add "heading", "Staff Transactions/ Activity Report"
    extParams.Add "FormatNumber", True
    extParams.Add "ShowRowTotal", False
    extParams.Add "IgnoreFromComputations", "Drug ID|Unit Cost"
    response.write SelectStaff(GetAllStaff("DPT002"), dateRange, staffID)

    sql = " SELECT * FROM ("
    sql = sql & "SELECT "
    sql = sql & "IIF( [Dispenses].DispenseDate='1/1/1900', [Dispenses].ReturnDate"
    sql = sql & " , [Dispenses].DispenseDate ) AS [Date / Time]"
    sql = sql & " , [Dispenses].DrugID AS [Drug ID], Drug.DrugName AS [Drug Name], Patient.PatientName AS [Patient Name]"
    sql = sql & " , [Type] AS [Transaction]"
    sql = sql & " , [Dispenses].UnitCost, [Dispenses].Qty AS [Sold Qty]"
    sql = sql & " , [Dispenses].ReturnQty AS [Return Qty]"
    sql = sql & " , ( CASE WHEN [Dispenses].[Type]='SALE' THEN [Dispenses].FinalAmt ELSE -[Dispenses].ReturnAmt END) AS [Final Cost]"
    sql = sql & " , Staff.StaffName AS [By] "
    sql = sql & " FROM ( "
    sql = sql & "   SELECT PatientID, DrugID, DispenseDate, '' AS ReturnDate "
    sql = sql & "       , Qty, UnitCost, FinalAmt"
    sql = sql & "       , 0 AS ReturnQty, 0 AS ReturnAmt"
    sql = sql & "       , SystemUserID, 'SALE' AS [Type]"
    sql = sql & "       FROM DrugSaleItems WHERE 1=1 AND DrugCategoryID<>'D002' "
    sql = sql & "           AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT PatientID, DrugID, DispenseDate, '' AS ReturnDate "
    sql = sql & "       , DispenseAmt1 AS Qty, UnitCost, DispenseAmt2 AS FinalAmt"
    sql = sql & "       , 0 AS ReturnQty, 0 AS ReturnAmt"
    sql = sql & "       , SystemUserID, 'SALE' AS [Type]"
    sql = sql & "       FROM DrugSaleItems2 WHERE 1=1 AND DrugCategoryID<>'D002' "
    sql = sql & "           AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT PatientID, DrugID, '' AS DispenseDate,  ReturnDate "
    sql = sql & "       , 0 AS Qty, UnitCost, 0 AS FinalAmt"
    sql = sql & "       , ReturnQty, FinalAmt AS ReturnAmt"
    sql = sql & "       , SystemUserID, 'RETURN' AS [Type]"
    sql = sql & "       FROM DrugReturnItems WHERE 1=1 AND DrugCategoryID<>'D002'"
    sql = sql & "           AND ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT PatientID, DrugID, '' AS DispenseDate, ReturnDate "
    sql = sql & "       , 0 AS Qty, UnitCost, 0 As FinalAmt"
    sql = sql & "       , ReturnQty, MainItemValue1 AS ReturnAmt"
    sql = sql & "       , SystemUserID, 'RETURN' AS [Type]"
    sql = sql & "       FROM DrugReturnItems2 WHERE 1=1 AND DrugCategoryID<>'D002'"
    sql = sql & "           AND ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & " )"
    sql = sql & " AS [Dispenses]"
    sql = sql & " LEFT JOIN Drug ON Drug.DrugID=[Dispenses].DrugID"
    sql = sql & " LEFT JOIN Patient ON Patient.PatientID=[Dispenses].PatientID"
    sql = sql & " INNER JOIN SystemUser ON SystemUser.SystemUserID=[Dispenses].SystemUserID AND SystemUser.JobScheduleID IN (SELECT JobScheduleID FROM JobSchedule WHERE JobSchedule.DepartmentID='DPT002')"
    sql = sql & " INNER JOIN Staff ON Staff.StaffID=SystemUser.StaffID "
    If staffID <> "" Then
        sql = sql & " AND Staff.StaffID ='" & staffID & "'"
    End If
    sql = sql & ") AS [Report] ORDER BY [Date / Time] ASC, [Patient Name] ASC "
    PRLTGLO_K.PrintSQLReport sql, Nothing, extParams
    
    Set extParams = Nothing
End Sub 'StaffActivity
Sub DispenseBySponsor(dateRange)
    Dim sql, SponsorID, extParams, lstArr, spnName, spnArr
    SponsorID = Request.QueryString("SponsorID")
    Set extParams = CreateObject("Scripting.Dictionary")
    
    lstArr = PRLTGLO_K.GetQueryResultsArray("SELECT SponsorID, SponsorName FROM Sponsor WHERE SponsorID IN (SELECT DISTINCT SponsorID FROM DrugSale WHERE DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "') ORDER BY SponsorName")
    response.write MakeDropDownLink("Select Sponsor", "SponsorID", lstArr, dateRange, SponsorID)
    
    If SponsorID <> "" Then
        spnArr = Array(SponsorID)
    Else
        spnArr = PRLTGLO_K.GetQueryResultsArray("SELECT SponsorID From Sponsor WHERE SponsorID IN (SELECT DISTINCT SponsorID FROM DrugSale WHERE DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "')")
    End If
    
    extParams.Add "FormatNumber", True
    extParams.Add "ShowRowTotal", False
    extParams.Add "IgnoreFromComputations", "Drug ID|Sold Qty|Avg. Cost|Return Qty"
        
    For Each SponsorID In spnArr
        sql = " SELECT * FROM ("
        sql = sql & "SELECT "
        sql = sql & " [Dispenses].DrugID AS [Drug ID], Drug.DrugName AS [Drug Name] "
        sql = sql & " , CAST(AVG([Dispenses].UnitCost) AS Money) AS [Avg. Cost], SUM([Dispenses].Qty) AS [Sold Qty] "
        sql = sql & " , SUM([Dispenses].ReturnQty) AS [Return Qty]"
        sql = sql & " , CAST(SUM([Dispenses].FinalAmt - [Dispenses].ReturnAmt) AS Money) AS [Tot. Sales]"
        sql = sql & " FROM ( "
        sql = sql & "   SELECT PatientID, DrugID, DispenseDate, '' AS ReturnDate "
        sql = sql & "       , Qty, UnitCost, FinalAmt"
        sql = sql & "       , 0 AS ReturnQty, 0 AS ReturnAmt"
        sql = sql & "       , SponsorID, 'SALE' AS [Type], JobScheduleID "
        sql = sql & "       FROM DrugSaleItems WHERE 1=1 AND DrugCategoryID<>'D002' "
        sql = sql & "           AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
        sql = sql & "   UNION ALL "
        sql = sql & "   SELECT PatientID, DrugID, DispenseDate, '' AS ReturnDate "
        sql = sql & "       , DispenseAmt1 AS Qty, UnitCost, DispenseAmt2 AS FinalAmt"
        sql = sql & "       , 0 AS ReturnQty, 0 AS ReturnAmt"
        sql = sql & "       , SponsorID, 'SALE' AS [Type], JobScheduleID "
        sql = sql & "       FROM DrugSaleItems2 WHERE 1=1 AND DrugCategoryID<>'D002' "
        sql = sql & "           AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
        sql = sql & "   UNION ALL "
        sql = sql & "   SELECT PatientID, DrugID, '' AS DispenseDate,  ReturnDate "
        sql = sql & "       , 0 AS Qty, UnitCost, 0 AS FinalAmt"
        sql = sql & "       , ReturnQty, FinalAmt AS ReturnAmt"
        sql = sql & "       , SponsorID, 'RETURN' AS [Type], JobScheduleID "
        sql = sql & "       FROM DrugReturnItems WHERE 1=1 AND DrugCategoryID<>'D002'"
        sql = sql & "           AND ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
        sql = sql & "   UNION ALL "
        sql = sql & "   SELECT PatientID, DrugID, '' AS DispenseDate, ReturnDate "
        sql = sql & "       , 0 AS Qty, UnitCost, 0 As FinalAmt"
        sql = sql & "       , ReturnQty, MainItemValue1 AS ReturnAmt"
        sql = sql & "       , SponsorID, 'RETURN' AS [Type], JobScheduleID "
        sql = sql & "       FROM DrugReturnItems2 WHERE 1=1 AND DrugCategoryID<>'D002' "
        sql = sql & "           AND ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
        sql = sql & " )"
        sql = sql & " AS [Dispenses]"
        sql = sql & " INNER JOIN Drug ON Drug.DrugID=[Dispenses].DrugID"
        sql = sql & " AND SponsorID ='" & SponsorID & "'"
        sql = sql & " INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=[Dispenses].JobScheduleID AND JobSchedule.DepartmentID='DPT002' "
        sql = sql & " GROUP BY Drug.DrugName, [Dispenses].DrugID"
        sql = sql & ") AS [Report]  "
        sql = sql & " ORDER BY [Drug Name] ASC"
        spnName = PRLTGLO_K.GetQueryResultStr("SELECT SponsorName FROM Sponsor WHERE SponsorID='" & SponsorID & "' ")
        extParams.Add "heading", "Drug Dispense Report [For " & spnName & "]"
        PRLTGLO_K.PrintSQLReport sql, Nothing, extParams
        extParams.Remove "heading"
    Next
    Set extParams = Nothing
    
    SponsorID = Request.QueryString("SponsorID")
    If SponsorID = "" Then
        PrintSponsorSummary (dateRange)
    End If
End Sub 'DispenseBySponsor
Sub PrintSponsorSummary(dateRange)
    Dim sql
    
    sql = " SELECT * FROM ("
    sql = sql & "SELECT "
    sql = sql & " spn.SponsorName AS [Sponsor Name] "
    sql = sql & " , CAST(SUM([Dispenses].FinalAmt - [Dispenses].ReturnAmt) AS Money) AS [Tot. Sales]"
    sql = sql & " FROM ( "
    sql = sql & "   SELECT PatientID, DrugID, DispenseDate, '' AS ReturnDate "
    sql = sql & "       , Qty, UnitCost, FinalAmt"
    sql = sql & "       , 0 AS ReturnQty, 0 AS ReturnAmt"
    sql = sql & "       , dsi.SponsorID, 'SALE' AS [Type], JobScheduleID "
    sql = sql & "       FROM DrugSaleItems AS dsi WHERE 1=1 AND DrugCategoryID<>'D002' "
    sql = sql & "           AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT PatientID, DrugID, DispenseDate, '' AS ReturnDate "
    sql = sql & "       , DispenseAmt1 AS Qty, UnitCost, DispenseAmt2 AS FinalAmt"
    sql = sql & "       , 0 AS ReturnQty, 0 AS ReturnAmt"
    sql = sql & "       , dsi.SponsorID, 'SALE' AS [Type], JobScheduleID"
    sql = sql & "       FROM DrugSaleItems2 AS dsi WHERE 1=1 AND DrugCategoryID<>'D002' "
    sql = sql & "           AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT PatientID, DrugID, '' AS DispenseDate,  ReturnDate "
    sql = sql & "       , 0 AS Qty, UnitCost, 0 AS FinalAmt"
    sql = sql & "       , ReturnQty, FinalAmt AS ReturnAmt"
    sql = sql & "       , dri.SponsorID, 'RETURN' AS [Type], JobScheduleID "
    sql = sql & "       FROM DrugReturnItems AS dri WHERE 1=1 AND DrugCategoryID<>'D002' "
    sql = sql & "           AND ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT PatientID, DrugID, '' AS DispenseDate, ReturnDate "
    sql = sql & "       , 0 AS Qty, UnitCost, 0 As FinalAmt"
    sql = sql & "       , ReturnQty, MainItemValue1 AS ReturnAmt"
    sql = sql & "       , dri.SponsorID, 'RETURN' AS [Type], JobScheduleID "
    sql = sql & "       FROM DrugReturnItems2 AS dri WHERE 1=1 AND DrugCategoryID<>'D002' "
    sql = sql & "           AND ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & " ) "
    sql = sql & " AS [Dispenses]"
    sql = sql & " INNER JOIN Drug ON Drug.DrugID=[Dispenses].DrugID"
    sql = sql & " INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=[Dispenses].JobScheduleID AND JobSchedule.DepartmentID='DPT002' "
    sql = sql & " INNER JOIN Sponsor AS spn ON spn.SponsorID=[Dispenses].SponsorID "
    sql = sql & " GROUP BY spn.SponsorName "
    sql = sql & ") AS [Report]  "
    sql = sql & " ORDER BY [Sponsor Name] ASC"
    
    PRLTGLO_K.PrintSQLReport sql, Nothing, Nothing
    
End Sub
Sub DispenseByMedicalService(dateRange)
    Dim sql, MedicalServiceID, extParams, lstArr, medicalServiceName, medArr
    MedicalServiceID = Request.QueryString("medicalServiceID")
    Set extParams = CreateObject("Scripting.Dictionary")
    
    lstArr = PRLTGLO_K.GetQueryResultsArray("SELECT MedicalServiceID, MedicalServiceName FROM MedicalService WHERE MedicalServiceID IN (SELECT DISTINCT MedicalServiceID FROM DrugSale WHERE DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "') ORDER BY MedicalServiceName ")
    response.write MakeDropDownLink("Select Service Type", "MedicalServiceID", lstArr, dateRange, MedicalServiceID)
    
    If MedicalServiceID <> "" Then
        medArr = Array(MedicalServiceID)
    Else
        medArr = PRLTGLO_K.GetQueryResultsArray("SELECT MedicalServiceID From MedicalService WHERE MedicalServiceID IN (SELECT DISTINCT MedicalServiceID FROM DrugSale WHERE DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "')")
    End If
    
    extParams.Add "FormatNumber", True
    extParams.Add "ShowRowTotal", False
    extParams.Add "IgnoreFromComputations", "Drug ID|Sold Qty|Avg. Cost|Return Qty"
        
    For Each MedicalServiceID In medArr
        sql = " SELECT * FROM ("
        sql = sql & "SELECT "
        sql = sql & " [Dispenses].DrugID AS [Drug ID], Drug.DrugName AS [Drug Name] "
        sql = sql & " , CAST(AVG([Dispenses].UnitCost) AS Money) AS [Avg. Cost], SUM([Dispenses].Qty) AS [Sold Qty] "
        sql = sql & " , SUM([Dispenses].ReturnQty) AS [Return Qty]"
        sql = sql & " , CAST(SUM([Dispenses].FinalAmt - [Dispenses].ReturnAmt) AS Money) AS [Tot. Sales]"
        sql = sql & " FROM ( "
        sql = sql & "   SELECT PatientID, DrugID, DispenseDate, '' AS ReturnDate "
        sql = sql & "       , Qty, UnitCost, FinalAmt"
        sql = sql & "       , 0 AS ReturnQty, 0 AS ReturnAmt"
        sql = sql & "       , MedicalServiceID, 'SALE' AS [Type], JobScheduleID"
        sql = sql & "       FROM DrugSaleItems WHERE 1=1 AND DrugCategoryID<>'D002' "
        sql = sql & "           AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
        sql = sql & "   UNION ALL "
        sql = sql & "   SELECT PatientID, DrugID, DispenseDate, '' AS ReturnDate "
        sql = sql & "       , DispenseAmt1 AS Qty, UnitCost, DispenseAmt2 AS FinalAmt"
        sql = sql & "       , 0 AS ReturnQty, 0 AS ReturnAmt"
        sql = sql & "       , (SELECT MedicalServiceID FROM Visitation WHERE Visitation.VisitationID=DrugSaleItems2.VisitationID) AS MedicalServiceID, 'SALE' AS [Type], JobScheduleID"
        sql = sql & "       FROM DrugSaleItems2 WHERE 1=1 AND DrugCategoryID<>'D002' "
        sql = sql & "           AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
        sql = sql & "   UNION ALL "
        sql = sql & "   SELECT PatientID, DrugID, '' AS DispenseDate,  ReturnDate "
        sql = sql & "       , 0 AS Qty, UnitCost, 0 AS FinalAmt"
        sql = sql & "       , ReturnQty, FinalAmt AS ReturnAmt"
        sql = sql & "       , MedicalServiceID, 'RETURN' AS [Type], JobScheduleID"
        sql = sql & "       FROM DrugReturnItems WHERE 1=1 AND DrugCategoryID<>'D002' "
        sql = sql & "           AND ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
        sql = sql & "   UNION ALL "
        sql = sql & "   SELECT PatientID, DrugID, '' AS DispenseDate, ReturnDate "
        sql = sql & "       , 0 AS Qty, UnitCost, 0 As FinalAmt"
        sql = sql & "       , ReturnQty, MainItemValue1 AS ReturnAmt"
        sql = sql & "       , MedicalServiceID, 'RETURN' AS [Type], JobScheduleID "
        sql = sql & "       FROM DrugReturnItems2 WHERE 1=1 AND DrugCategoryID<>'D002' "
        sql = sql & "           AND ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
        sql = sql & " )"
        sql = sql & " AS [Dispenses]"
        sql = sql & " INNER JOIN Drug ON Drug.DrugID=[Dispenses].DrugID"
        sql = sql & " INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=[Dispenses].JobScheduleID AND JobSchedule.DepartmentID='DPT002' "
        sql = sql & " AND MedicalServiceID ='" & MedicalServiceID & "'"
        sql = sql & " GROUP BY Drug.DrugName, [Dispenses].DrugID"
        sql = sql & ") AS [Report]  "
        sql = sql & " ORDER BY [Drug Name] ASC"
        
        spnName = PRLTGLO_K.GetQueryResultStr("SELECT MedicalServiceName FROM MedicalService WHERE MedicalServiceID='" & MedicalServiceID & "' ")
        extParams.Add "heading", "Drug Dispense Report [For " & spnName & "]"
        PRLTGLO_K.PrintSQLReport sql, Nothing, extParams
        extParams.Remove "heading"
    Next
    Set extParams = Nothing
    
    
    MedicalServiceID = Request.QueryString("MedicalServiceID")
    If MedicalServiceID = "" Then
        PrintMedicalServiceSummary (dateRange)
    End If
    
End Sub 'DispenseByMedicalService
Sub PrintMedicalServiceSummary(dateRange)
    Dim sql
    
    sql = " SELECT * FROM ("
    sql = sql & "SELECT "
    sql = sql & " md.MedicalServiceName AS [Service Type] "
    sql = sql & " , CAST(SUM([Dispenses].FinalAmt - [Dispenses].ReturnAmt) AS Money) AS [Tot. Sales] "
    sql = sql & " FROM ( "
    sql = sql & "   SELECT PatientID, DrugID, DispenseDate, '' AS ReturnDate "
    sql = sql & "       , Qty, UnitCost, FinalAmt"
    sql = sql & "       , 0 AS ReturnQty, 0 AS ReturnAmt"
    sql = sql & "       , MedicalServiceID, 'SALE' AS [Type], JobScheduleID"
    sql = sql & "       FROM DrugSaleItems WHERE 1=1 AND DrugCategoryID<>'D002'"
    sql = sql & "           AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT PatientID, DrugID, DispenseDate, '' AS ReturnDate "
    sql = sql & "       , DispenseAmt1 AS Qty, UnitCost, DispenseAmt2 AS FinalAmt"
    sql = sql & "       , 0 AS ReturnQty, 0 AS ReturnAmt"
    sql = sql & "       , (SELECT MedicalServiceID FROM Visitation WHERE Visitation.VisitationID=DrugSaleItems2.VisitationID) AS MedicalServiceID, 'SALE' AS [Type], JobScheduleID"
    sql = sql & "       FROM DrugSaleItems2 WHERE 1=1 AND DrugCategoryID<>'D002' "
    sql = sql & "           AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT PatientID, DrugID, '' AS DispenseDate,  ReturnDate "
    sql = sql & "       , 0 AS Qty, UnitCost, 0 AS FinalAmt"
    sql = sql & "       , ReturnQty, FinalAmt AS ReturnAmt"
    sql = sql & "       , MedicalServiceID, 'RETURN' AS [Type], JobScheduleID"
    sql = sql & "       FROM DrugReturnItems WHERE 1=1 AND DrugCategoryID<>'D002' "
    sql = sql & "           AND ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT PatientID, DrugID, '' AS DispenseDate, ReturnDate "
    sql = sql & "       , 0 AS Qty, UnitCost, 0 As FinalAmt"
    sql = sql & "       , ReturnQty, MainItemValue1 AS ReturnAmt"
    sql = sql & "       , MedicalServiceID, 'RETURN' AS [Type], JobScheduleID "
    sql = sql & "       FROM DrugReturnItems2 WHERE 1=1 AND DrugCategoryID<>'D002' "
    sql = sql & "           AND ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & " )"
    sql = sql & " AS [Dispenses]"
    sql = sql & " INNER JOIN Drug ON Drug.DrugID=[Dispenses].DrugID"
    sql = sql & " INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=[Dispenses].JobScheduleID AND JobSchedule.DepartmentID='DPT002' "
    sql = sql & " INNER JOIN MedicalService AS md ON md.MedicalServiceID=[Dispenses].MedicalServiceID "
    sql = sql & " GROUP BY MedicalServiceName"
    sql = sql & ") AS [Report]  "
    sql = sql & " ORDER BY [Service Type] ASC"
    
    PRLTGLO_K.PrintSQLReport sql, Nothing, Nothing
    
End Sub
Sub SaleReturn(dateRange)
    Dim sql, SponsorID, extParams, lstArr
    SponsorID = Request.QueryString("SponsorID")
    Set extParams = CreateObject("Scripting.Dictionary")
    
    extParams.Add "heading", "Drug Sale / Returns Report"
    extParams.Add "FormatNumber", True
    extParams.Add "ShowRowTotal", False
    extParams.Add "IgnoreFromComputations", "Drug ID"
    
'   lstArr = PRLTGLO_K.GetQueryResultsArray("SELECT SponsorID, SponsorName FROM Sponsor ")
'   Response.Write MakeDropDownLink ("Select Sponsor", "SponsorID", lstarr, dateRange, sponsorID)

'    sql = " SELECT * FROM ("
'    sql = sql & "SELECT "
'    sql = sql & "IIF( [Dispenses].DispenseDate='1/1/1900', [Dispenses].ReturnDate"
'    sql = sql & " , [Dispenses].DispenseDate ) AS [Date / Time]"
'    sql = sql & " ,[Dispenses].DrugID AS [Drug ID], Drug.DrugName AS [Drug Name], Patient.PatientName AS [Patient Name]"
'    sql = sql & " , [Dispenses].UnitCost, [Dispenses].Qty AS [Sold Qty]"
'    sql = sql & " , [Dispenses].ReturnQty AS [Return Qty]"
'    sql = sql & " , ( CASE WHEN [Dispenses].[Type]='SALE' THEN [Dispenses].FinalAmt ELSE -[Dispenses].ReturnAmt END) AS [Final Cost]"
'    sql = sql & " , Staff.StaffName AS [By]"
'    sql = sql & " FROM ( "
'    sql = sql & "   SELECT PatientID, DrugID, DispenseDate, '' AS ReturnDate "
'    sql = sql & "       , Qty, UnitCost, FinalAmt"
'    sql = sql & "       , 0 AS ReturnQty, 0 AS ReturnAmt"
'    sql = sql & "       , SystemUserID, 'SALE' AS [Type]"
'    sql = sql & "       FROM DrugSaleItems WHERE 1=1"
'    sql = sql & "           AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
'    sql = sql & "   UNION ALL "
'    sql = sql & "   SELECT PatientID, DrugID, DispenseDate, '' AS ReturnDate "
'    sql = sql & "       , DispenseAmt1 AS Qty, UnitCost, DispenseAmt2 AS FinalAmt"
'    sql = sql & "       , 0 AS ReturnQty, 0 AS ReturnAmt"
'    sql = sql & "       , SystemUserID, 'SALE' AS [Type]"
'    sql = sql & "       FROM DrugSaleItems2 WHERE 1=1 "
'    sql = sql & "           AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
'    sql = sql & "   UNION ALL "
'    sql = sql & "   SELECT PatientID, DrugID, '' AS DispenseDate,  ReturnDate "
'    sql = sql & "       , 0 AS Qty, UnitCost, 0 AS FinalAmt"
'    sql = sql & "       , ReturnQty, FinalAmt AS ReturnAmt"
'    sql = sql & "       , SystemUserID, 'RETURN' AS [Type]"
'    sql = sql & "       FROM DrugReturnItems WHERE 1=1"
'    sql = sql & "           AND ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
'    sql = sql & "   UNION ALL "
'    sql = sql & "   SELECT PatientID, DrugID, '' AS DispenseDate, ReturnDate "
'    sql = sql & "       , 0 AS Qty, UnitCost, 0 As FinalAmt"
'    sql = sql & "       , ReturnQty, MainItemValue1 AS ReturnAmt"
'    sql = sql & "       , SystemUserID, 'RETURN' AS [Type]"
'    sql = sql & "       FROM DrugReturnItems2 WHERE 1=1"
'    sql = sql & "           AND ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
'    sql = sql & " )"
'    sql = sql & " AS [Dispenses]"
'    sql = sql & " LEFT JOIN Drug ON Drug.DrugID=[Dispenses].DrugID"
'    sql = sql & " LEFT JOIN Patient ON Patient.PatientID=[Dispenses].PatientID"
'    sql = sql & " LEFT JOIN SystemUser ON SystemUser.SystemUserID=[Dispenses].SystemUserID "
'    sql = sql & " INNER JOIN Staff ON Staff.StaffID=SystemUser.StaffID "
'    If StaffID <> "" Then
'        sql = sql & " AND Staff.StaffID ='" & StaffID & "'"
'    End If
'    sql = sql & ") AS [Report] ORDER BY [Date / Time] ASC, [Patient Name] ASC "

    sql = "SELECT  "
    sql = sql & " [Drug ID], [Drug Name] "
    sql = sql & " , CAST(AVG([Unit Cost]) AS Money) AS [Avg. Cost] "
    sql = sql & " , SUM([Sold Qty]) AS [Qty Sold] "
    sql = sql & " , SUM([Return Qty]) AS [Qty Ret.] "
    sql = sql & " , CAST(SUM([Final Amt]) AS Money) AS [Final Amt] "
    sql = sql & " FROM ( "
    sql = sql & " SELECT  dsi.DrugID AS [Drug ID], dg.DrugName AS [Drug Name], dsi.UnitCost AS [Unit Cost]"
    sql = sql & "   ,dsi.Qty AS [Sold Qty], dri.ReturnQty AS [Return Qty]"
    sql = sql & "   , (dsi.FinalAmt - ISNULL(dri.FinalAmt, 0)) AS [Final Amt], '1' AS [Type] "
    sql = sql & "   FROM DrugSaleItems AS dsi "
    sql = sql & "   INNER JOIN Drug AS dg ON dsi.DrugID=dg.DrugID AND dsi.DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' AND dsi.DrugCategoryID<>'D002' "
    sql = sql & "   INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=dsi.JobScheduleID AND JobSchedule.DepartmentID='DPT002' "
    sql = sql & "   LEFT JOIN DrugReturnItems AS dri ON dri.DrugSaleID=dsi.DrugSaleID AND dri.DrugID=dsi.DrugID AND dri.ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' AND dri.DrugCategoryID<>'D002' "
    sql = sql & " "
    sql = sql & " UNION ALL "
    sql = sql & " SELECT  dsi.DrugID AS [Drug ID], dg.DrugName AS [Drug Name], dsi.UnitCost AS [Unit Cost] "
    sql = sql & "   , dsi.DispenseAmt1 AS [Sold Qty], dri.ReturnQty AS [Return Qty] "
    sql = sql & "   , (dsi.DispenseAmt2 - ISNULL(dri.MainItemValue1, 0)) AS [Final Amt], '2' AS [Type]  "
    sql = sql & "   FROM DrugSaleItems2 AS dsi "
    sql = sql & "   INNER JOIN Drug AS dg ON dsi.DrugID=dg.DrugID AND dsi.DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' AND dsi.DrugCategoryID<>'D002' "
    sql = sql & "   INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=dsi.JobScheduleID AND JobSchedule.DepartmentID='DPT002' "
    sql = sql & "   LEFT JOIN DrugReturnItems2 AS dri ON dri.DrugSaleID=dsi.DrugSaleID AND dri.DrugID=dsi.DrugID AND dri.ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' AND dri.DrugCategoryID<>'D002' "
    sql = sql & " "
    sql = sql & " ) AS [Report] "
    sql = sql & "   GROUP BY [Drug ID], [Drug Name]"
    sql = sql & "   ORDER BY [Drug Name]"
    
    PRLTGLO_K.PrintSQLReport sql, Nothing, extParams
    
    Set extParams = Nothing
    
End Sub 'SaleReturn
Sub TopDispenseReport(dateRange)
    Dim extParams, sql
    
    Set extParams = CreateObject("Scripting.Dictionary")
    
    extParams.Add "heading", "TOP Drug Dispenses All Pharm(s) "
    extParams.Add "ShowRowTotal", False
    extParams.Add "IgnoreFromComputations", "Drug ID"
    
    sql = "SELECT DrugID AS [Drug ID], DrugName AS [Drug Name], COUNT(DrugID) AS [Number Of Dispenses], SUM(Qty) AS [Total Quantity] "
    sql = sql & " FROM ("
    sql = sql & "   SELECT dsi.DrugID, drg.DrugName, dsi.Qty AS Qty FROM DrugSaleItems AS dsi INNER JOIN Drug AS drg ON drg.DrugID=dsi.DrugID AND dsi.DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "       INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=dsi.JobScheduleID AND JobSchedule.DepartmentID='DPT002' AND dsi.DrugCategoryID<>'D002' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT dsi.DrugID, drg.DrugName, dsi.DispenseAmt1 AS Qty FROM DrugSaleItems2 AS dsi INNER JOIN Drug AS drg ON drg.DrugID=dsi.DrugID AND dsi.DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
    sql = sql & "      INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=dsi.JobScheduleID AND JobSchedule.DepartmentID='DPT002' AND dsi.DrugCategoryID<>'D002' "
    sql = sql & ""
    sql = sql & ") "
    sql = sql & " AS [Report]"
    sql = sql & " GROUP BY DrugID, DrugName"
    sql = sql & " ORDER BY COUNT(DrugID) DESC, DrugName"
    
    PRLTGLO_K.PrintSQLReport sql, Nothing, extParams
    
    Set extParams = Nothing
End Sub 'TopDispenseReport
Sub TopDispenseReportOPD(dateRange)
    Dim extParams, sql
    
    Set extParams = CreateObject("Scripting.Dictionary")
    
    extParams.Add "heading", "TOP OPD Pharm Drug Dispenses "
    extParams.Add "ShowRowTotal", False
    extParams.Add "IgnoreFromComputations", "Drug ID"
    
    sql = "SELECT DrugID AS [Drug ID], DrugName AS [Drug Name], COUNT(DrugID) AS [Number Of Dispenses], SUM(Qty) AS [Total Quantity] "
    sql = sql & " FROM ("
    sql = sql & "   SELECT dsi.DrugID, drg.DrugName, dsi.Qty AS Qty FROM DrugSaleItems AS dsi INNER JOIN Drug AS drg ON drg.DrugID=dsi.DrugID AND dsi.DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' AND dsi.JobScheduleID IN ('S22', 'M0601', 'M0602') AND dsi.DrugCategoryID<>'D002' "
    'sql = sql & "       INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=dsi.JobScheduleID AND JobSchedule.DepartmentID='DPT002' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT dsi.DrugID, drg.DrugName, dsi.DispenseAmt1 AS Qty FROM DrugSaleItems2 AS dsi INNER JOIN Drug AS drg ON drg.DrugID=dsi.DrugID AND dsi.DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' AND dsi.JobScheduleID IN ('S22', 'M0601', 'M0602') AND dsi.DrugCategoryID<>'D002' "
    'sql = sql & "      INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=dsi.JobScheduleID AND JobSchedule.DepartmentID='DPT002' "
    sql = sql & ""
    sql = sql & ") "
    sql = sql & " AS [Report]"
    sql = sql & " GROUP BY DrugID, DrugName"
    sql = sql & " ORDER BY COUNT(DrugID) DESC, DrugName"
    
    PRLTGLO_K.PrintSQLReport sql, Nothing, extParams
    
    Set extParams = Nothing
End Sub 'TopDispenseReport
Sub TopDispenseReportIPD(dateRange)
    Dim extParams, sql
    
    Set extParams = CreateObject("Scripting.Dictionary")
    
    extParams.Add "heading", "TOP In-Patient Pharm Drug Dispenses "
    extParams.Add "ShowRowTotal", False
    extParams.Add "IgnoreFromComputations", "Drug ID"
    
    sql = "SELECT DrugID AS [Drug ID], DrugName AS [Drug Name], COUNT(DrugID) AS [Number Of Dispenses], SUM(Qty) AS [Total Quantity] "
    sql = sql & " FROM ("
    sql = sql & "   SELECT dsi.DrugID, drg.DrugName, dsi.Qty AS Qty FROM DrugSaleItems AS dsi INNER JOIN Drug AS drg ON drg.DrugID=dsi.DrugID AND dsi.DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' AND dsi.JobScheduleID IN ('S22A', 'M0601A', 'M0602A') AND dsi.DrugCategoryID<>'D002' "
    'sql = sql & "       INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=dsi.JobScheduleID AND JobSchedule.DepartmentID='DPT002' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT dsi.DrugID, drg.DrugName, dsi.DispenseAmt1 AS Qty FROM DrugSaleItems2 AS dsi INNER JOIN Drug AS drg ON drg.DrugID=dsi.DrugID AND dsi.DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' AND dsi.JobScheduleID IN ('S22A', 'M0601A', 'M0602A') AND dsi.DrugCategoryID<>'D002' "
    'sql = sql & "      INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=dsi.JobScheduleID AND JobSchedule.DepartmentID='DPT002' "
    sql = sql & ""
    sql = sql & ") "
    sql = sql & " AS [Report]"
    sql = sql & " GROUP BY DrugID, DrugName"
    sql = sql & " ORDER BY COUNT(DrugID) DESC, DrugName"
    
    PRLTGLO_K.PrintSQLReport sql, Nothing, extParams
    
    Set extParams = Nothing
End Sub 'TopDispenseReport
Sub TopReturnsReport(dateRange)
    Dim extParams, sql
    
    Set extParams = CreateObject("Scripting.Dictionary")
    
    extParams.Add "heading", "TOP Drug Returns "
    extParams.Add "IgnoreFromComputations", "Drug ID"
    extParams.Add "ShowRowTotal", False
    
    sql = "SELECT DrugID AS [Drug ID], DrugName AS [Drug Name], COUNT(DrugID) AS [Number Of Dispenses], SUM(Qty) AS [Total Return] "
    sql = sql & " FROM ("
    sql = sql & "   SELECT dri.DrugID, drg.DrugName, dri.ReturnQty AS Qty FROM DrugReturnItems AS dri INNER JOIN Drug AS drg ON drg.DrugID=dri.DrugID AND dri.ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' AND dri.DrugCategoryID<>'D002' "
    sql = sql & "       INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=dri.JobScheduleID AND JobSchedule.DepartmentID='DPT002' "
    sql = sql & "   UNION ALL "
    sql = sql & "   SELECT dri.DrugID, drg.DrugName, dri.ReturnQty AS Qty FROM DrugReturnItems2 AS dri INNER JOIN Drug AS drg ON drg.DrugID=dri.DrugID AND dri.ReturnDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' AND dri.DrugCategoryID<>'D002' "
    sql = sql & "       INNER JOIN JobSchedule ON JobSchedule.JobScheduleID=dri.JobScheduleID AND JobSchedule.DepartmentID='DPT002' "
    sql = sql & ""
    sql = sql & ") "
    sql = sql & " AS [Report]"
    sql = sql & " GROUP BY DrugID, DrugName"
    sql = sql & " ORDER BY COUNT(DrugID) DESC, DrugName"
    
    PRLTGLO_K.PrintSQLReport sql, Nothing, extParams
    
    Set extParams = Nothing
End Sub 'TopReturnsReport

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
