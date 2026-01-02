'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim sql, rptGen, args, dateRange

dateRange = Split(Trim(Request.QueryString("Printfilter0")), "||")


Set rptGen = New PRTGLO_RptGen

    sql = "SELECT Drug.DrugName AS [Drug Name], DrugIssueItems.DrugRequest2ID AS [ID], "
    sql = sql & " DrugIssueItems.RequestedQty AS [Requested Quantity], staff.staffName AS [Requested By], "
    sql = sql & " DrugIssueItems.IssuedQty AS [Issued Qty], Drug.RetailUnitCost AS [RetailUnitCost], "
    sql = sql & " DrugIssueItems.IssuedQty*Drug.RetailUnitCost AS [Total Amt],DrugStore.DrugStoreName AS [Drug Store],"
    
    sql = sql & " DrugRequestStore.DrugRequestStoreName AS [Store], "
    'sql = sql & " ItemIssueItems.systemUserID AS [IssuedBy], "
    sql = sql & " convert(varchar, DrugIssueItems.RequestDate, 106) AS [RequestDate],"
    sql = sql & " convert(varchar, DrugIssueItems.IssuedDate1, 106) AS [IssueDate] "
    sql = sql & " FROM DrugIssueItems "
    sql = sql & " LEFT JOIN Drug ON Drug.DrugID = DrugIssueItems.DrugID "
    sql = sql & " LEFT JOIN staff ON Staff.StaffID = DrugIssueItems.RequestedBy"
    sql = sql & " LEFT JOIN UnitOfMeasure ON UnitOfMeasure.UnitOfMeasureID = DrugIssueItems.UnitOfMeasureid"
    sql = sql & " LEFT JOIN DrugStore On DrugStore.DrugStoreID = DrugIssueItems.DrugStoreid "
    sql = sql & " LEFT JOIN DrugRequestStore ON DrugRequestStore.DrugRequestStoreID = DrugIssueItems.DrugRequestStoreID"
    
    sql = sql & " WHERE DrugIssueItems.IssuedDate1 BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' AND DrugIssueItems.DrugCategoryID<>'D002'  "
    sql = sql & " AND DrugRequestStore.DrugRequestStoreID = 'S22' "
    sql = sql & " ORDER BY Drug.DrugName ASC"
    
    
args = "showcolumntotal=yes;SubGroupFields=Drug Name;hiddenfields=Store|ID;formatmoneyfields=RetailUnitCost|Total Amt;"
args = args & " ignoreFromComputations=Drug Name|ID|Requested Quantity|Requested By|RetailUnitCost|Drug Store|"
args = args & " |Store|IssuedBy|RequestDate|IssueDate; "
args = args & " Title=Issued Drugs between " & dateRange(0) & " to " & dateRange(1) & " by Pharmacy ;"
rptGen.PrintSQLReport sql, args

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
