'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim rst, sql, drange, sDate, edate
Set rst = CreateObject("ADODB.RecordSet")
drange = Split(Trim(Request("PrintFilter0")), "||")
sDate = drange(0)
If sDate = "" Then
    sDate = DateAdd("d", -1, Now)
End If
If UBound(drange) = 0 Then
    edate = Now
Else
    edate = drange(1)
End If


filter1 = Trim(Request("filter"))

sql = ""
sql = sql & "SELECT CAST(DrugSale.DispenseDate as Date) [Date],DrugSale.DrugSaleID [DrugSaleID],Drug.DrugID [DrugID],Drug.DrugName [Drug Name] "
sql = sql & ",DrugSaleItems.Qty [Quantity],DrugSaleItems.FinalAmt [Unit Cost],DrugSaleItems.FinalAmt [Final Amount] "
sql = sql & "FROM DrugSale  "
sql = sql & "INNER JOIN DrugSaleItems ON DrugSaleItems.DrugSaleID = DrugSale.DrugSaleID   "
sql = sql & "INNER JOIN Drug ON Drug.DrugID = DrugSaleItems.DrugID AND Drug.DrugCategoryID = 'D002' "
sql = sql & "INNER JOIN SystemUser ON SystemUser.SystemUserID = DrugSale.SystemUserID "
sql = sql & "INNER JOIN Staff ON Staff.StaffID = SystemUser.StaffID "
sql = sql & " WHERE DrugSale.DispenseDate BETWEEN '" & sDate & "' AND '" & edate & "' "

args = "title=Mini Mart Sales from " & sDate & " to " & edate & ""
If filter1 <> "" Then
    sql = sql & " AND DrugSale.SystemUserID = '" & filter1 & "' "
    args = args & " By " & GetComboName("Staff", GetComboNameFld("SystemUser", filter1, "StaffID"))
End If

sql = sql & " ORDER BY DrugSale.DispenseDate DESC "


response.write vbCrLf & "<span>Filter:</span>"
response.write vbCrLf & "<select onchange=""userSelected(this)"">"
response.write vbCrLf & "    <option value="""">ALL</option>"
GetSystemUsers "s22c"
response.write vbCrLf & "</select>"
response.write vbCrLf & "<script>"
response.write vbCrLf & "function userSelected(el){"
response.write vbCrLf & "    const url = new URL(window.location);"
response.write vbCrLf & "    url.searchParams.set('filter',el.value);"
response.write vbCrLf & "    window.location.href = url;"
response.write vbCrLf & "}"
response.write vbCrLf & "</script>"

Set rptGen = New PRTGLO_RptGen2
args = args & ";showColumnTotal=YES"
args = args & ";FormatMoneyFields=Final Amount|Unit Cost"
args = args & ";IgnoreFromComputations=DrugSaleID|DrugID|Unit Cost|Date"
args = args & ";IncludeInComputations=Final Amount"

rptGen.PrintSQLReport sql, args

Sub GetSystemUsers(jb)
    Dim rst, activeUser, sql
    activeUser = Trim(Request("filter"))
    Set rst = CreateObject("ADODB.RecordSet")
    sql = ""
    sql = sql & "SELECT SystemUser.SystemUserID,Staff.StaffName "
    sql = sql & "from SystemUser "
    sql = sql & "INNER JOIN Staff on Staff.StaffID = SystemUser.StaffID "
    sql = sql & "where SystemUser.JobScheduleID = '" & jb & "' "
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            response.write "<option value='" & rst.fields("systemuserid") & "' "
            If UCase(activeUser) = UCase(rst.fields("systemuserid")) Then
                response.write " selected "
            End If
            response.write ">" & rst.fields("StaffName") & "</option>"
            rst.MoveNext
            response.Flush
        Loop
    End If
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
