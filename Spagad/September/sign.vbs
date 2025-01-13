'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim rst, sql, cnt, tot, qty, drg, ucst, amt
Dim firstApprovBy, secondApprovBy, thirdApprovBy, fourthApprovBy
Dim firstApprovDt, secondApprovDt, thirdApprovDt, fourthApprovDt

Set rst = CreateObject("ADODB.Recordset")

response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""0"" cellpadding=""0"">"
response.write "<tr>"
response.write "<td><img src=""images/letterhead5.jpg""></td>"
'AddReportHeader

response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"
'response.write "<tr>"
'response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'response.write "</tr>"
response.write "<tr>"
response.write "<td height=""10"" align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:15pt"">"
response.write "LOCAL PURCHASE ORDER&nbsp;&nbsp;&nbsp;[ No :&nbsp;" & GetRecordField("DrugPurOrderID") & " ]</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td height=""0"" align=""right"" bgcolor=""#FFFFFF"" style=""font-weight:bold;font-family: Arial;font-size:9pt"">"
response.write "Date&nbsp;:&nbsp;&nbsp;" & GetRecordField("WorkingDayName") & "</td>"
response.write "</tr>"
'response.write "<tr>"
'response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'response.write "</tr>"
response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table border=""1"" width=""" & (PrintWidth) & """ cellspacing=""0"" cellpadding=""5"" bgcolor=""White"" style=""border-collapse:collapse;font-size: 9pt; font-family: Arial"">"
response.write "<tr>"
'Address
response.write "<td width=""50%"">"
response.write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""3"" bgcolor=""White"" style=""border-collapse:collapse;font-size: 10pt; font-family: Arial"">"
response.write "<tr>"
response.write "<td><b>To :</b></td>"
response.write "<td align=""left"">THE SALES MANAGER</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td><b>Address :</b></td>"
response.write "<td align=""left"">&nbsp;" & GetRecordField("SupplierName") & "</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td colspan=""2"" align=""left"">&nbsp;" & GetComboNameFld("Supplier", GetRecordField("SupplierID"), "Address") & "</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td colspan=""2"" align=""left"">&nbsp;" & GetComboNameFld("Supplier", GetRecordField("SupplierID"), "City") & "</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
'Divider
'response.write "<td width=""4%""></td>"
'Delivery Info
response.write "<td width=""50%"">"
response.write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""3"" bgcolor=""White"" style=""border-collapse:collapse;font-size: 10pt; font-family: Arial"">"
response.write "<tr>"
response.write "<td><b>Delivery Instructions :</b></td>"
response.write "<td><b>1.</b> Attach original copy of LPO</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td colspan=""2""><b>2.</b> Invoice & Waybill&nbsp;&nbsp;&nbsp;<b>3.</b> Social Security Clearance</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td colspan=""2""><b>*</b>&nbsp;" & GetRecordField("Remarks") & "</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td colspan=""2""><b>Quote LPO No. on all your bills <br>LPO valid for 30 days</b></td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"

response.write "</table>"
response.write "</td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td align=""left"">Please supply the following and in accordance with the specifications</td>"
response.write "</tr>"

'If Len(GetRecordField("Remarks")) > 0 Then
'    response.write "<tr>"
'    response.write "<td align=""left""><div style=""font-weight:bold;display:inline-block;vertical-align:top;padding:5px;"">REMARKS</div><div style=""display:inline-block;vertical-align:top;padding:5px;"">" & Replace(GetRecordField("Remarks"), vbCrLf, "<br>") & "</div></td>"
'    response.write "</tr>"
'End If

response.write "<tr>"
response.write "<td valign=""bottom"" height=""30"" align=""left""><b>Your reference</b> ...............................................................................</b></td>"
response.write "</tr>"

'response.write "<tr>"
'response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'response.write "</tr>"

'List Order Items
response.write "<tr>"
response.write "<td height=""200"" valign=""top"" align=""center"">"
response.write "<table width=""100%"" border=""1"" cellspacing=""0"" cellpadding=""2"" style=""border-collapse:collapse; font-size: 9pt; font-family: Arial"">"
response.write "<tr>"
response.write "<td align=""center""><b>No.</b></td>"
response.write "<td align=""left""><b>Code</b></td>"
response.write "<td align=""left""><b>Full Description</b></td>"
response.write "<td align=""right""><b>Qty</b></td>"
response.write "<td align=""right""><b>Price (GHC)</b></td>"
response.write "<td align=""right""><b>Amount (GHC)</b></td>"
response.write "</tr>"

cnt = 0
tot = 0

sql = "select * from drugPurOrderItems where drugpurorderid='" & Trim(GetRecordField("DrugPurOrderid")) & "' order by drugid"
With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
      cnt = cnt + 1
      drg = .fields("drugid")
      qty = .fields("orderquantity")
      ucst = .fields("orderamount1")
      amt = .fields("orderamount2")
      tot = tot + amt
      response.write "<tr>"
      response.write "<td align=""center""><b>" & CStr(cnt) & ".</b></td>"
      response.write "<td>" & UCase(drg) & "</td>"
      response.write "<td>" & GetComboName("Drug", drg) & "</td>"
      response.write "<td align=""right"">" & CStr(qty) & "</td>"
      response.write "<td align=""right"">" & FormatNumber(ucst, 4, , , -1) & "</td>"
      response.write "<td align=""right"">" & FormatNumber(amt, 2, , , -1) & "</td>"
      response.write "</tr>"
      .MoveNext
    Loop
  End If
  .Close
End With
'By Tender
sql = "select * from drugPurOrderItems2 where drugpurorderid='" & Trim(GetRecordField("DrugPurOrderid")) & "' order by drugid"
With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
      cnt = cnt + 1
      drg = .fields("drugid")
      qty = .fields("orderquantity")
      ucst = .fields("orderamount1")
      amt = .fields("orderamount2")
      tot = tot + amt
      response.write "<tr>"
      response.write "<td align=""right""><b>" & CStr(cnt) & "*</b></td>"
      response.write "<td>" & UCase(drg) & "</td>"
      response.write "<td>" & GetComboName("Drug", drg) & "</td>"
      response.write "<td align=""right"">" & CStr(qty) & "</td>"
      response.write "<td align=""right"">" & FormatNumber(ucst, 4, , , -1) & "</td>"
      response.write "<td align=""right"">" & FormatNumber(amt, 2, , , -1) & "</td>"
      response.write "</tr>"
      .MoveNext
    Loop
    
  End If
  .Close
End With

response.write "<tr>"
response.write "<td align=""right"">-</td>"
response.write "<td align=""center"" colspan=""4""><b>TOTALS</b></td>"
response.write "<td align=""right""><b>" & FormatNumber(tot, 2, , , -1) & "</b></td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"


response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""7"" bgcolor=""White"" style=""margin-top:10px;border-collapse:collapse;font-weight:bold; font-size: 9pt; font-family: Arial"">"
response.write "<tr>"
response.write "<td valign=""bottom"" colspan=""4"">Order Status: <span style=""border-bottom:1px dotted black;"">"
response.write GetComboName("TransProcessVal", GetRecotmp = GetDrugPurOrderProUser(GetRecordField("DrugPurOrderID"), "DrugPurOrderPro-T002") 'supply chain
If UBound(tmp) > 0 Then
    firstApprovBy = tmp(0)
    firstApprovDt = tmp(1)
End If

tmp = GetDrugPurOrderProUser(GetRecordField("DrugPurOrderID"), "DrugPurOrderPro-T006") 'finance
If UBound(tmp) > 0 Then
    secondApprovBy = tmp(0)
    secondApprovDt = tmp(1)
Else
    tmp = GetDrugPurOrderProUser(GetRecordField("DrugPurOrderID"), "DrugPurOrderPro-T011") 'finance
    If UBound(tmp) > 0 Then
        secondApprovBy = tmp(0)
        secondApprovDt = tmp(1)
    End If
End If

tmp = GetDrugPurOrderProUser(GetRecordField("DrugPurOrderID"), "DrugPurOrderPro-T003") 'admin
If UBound(tmp) > 0 Then
    thirdApprovBy = tmp(0)
    thirdApprovDt = tmp(1)
End If

tmp = GetDrugPurOrderProUser(GetRecordField("DrugPurOrderID"), "DrugPurOrderPro-T004") 'ceo
If UBound(tmp) > 0 Then
    fourthApprovBy = tmp(0)
    fourthApprovDt = tmp(1)
End If
rdField("TransProcessValID"))
tmp = GetDrugPurOrderProUser(GetRecordField("DrugPurOrderID"), GetRecordField("TransProcessValID"))
If UBound(tmp) > 0 Then response.write " (by " & GetStaffName(tmp(0)) & ", " & FormatDateDetail(tmp(1)) & ")"
response.write "</span></td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td valign=""top"" align=""center"" width=""25%"">Prepared by:</td>"
response.write "<td valign=""top"" align=""center"" width=""25%"">Approved by Supply Chain Manager</td>"
response.write "<td valign=""top"" align=""center"" width=""25%"">Approved by Finance</td>"

If Len(fourthApprovBy) > 0 Then
    response.write "<td valign=""top"" align=""center"" width=""25%"">Approved by CEO</td>"
Else
    response.write "<td valign=""top"" align=""center"" width=""25%"">Approved by Administrator</td>"
End If

'response.write "<td valign=""top"" align=""center"" width=""33%"">Counter Signed</td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td valign=""bottom"" height=""50"" width=""25%"" ><img src=""" & GetSignaturePath(GetRecordField("SystemUserID")) & """ style=""border-bottom:1px dotted black;width:150px;height:130px;""/></td>"
response.write "<td valign=""bottom"" height=""50"" width=""25%"" ><img src=""" & GetSignaturePath(firstApprovBy) & """ style=""border-bottom:1px dotted black;width:150px;height:130px;""</td>"
response.write "<td valign=""bottom"" height=""50"" width=""25%"" ><img src=""" & GetSignaturePath(secondApprovBy) & """ style=""border-bottom:1px dotted black;width:150px;height:130px;""</td>"
If Len(fourthApprovBy) > 0 Then
    response.write "<td valign=""bottom"" height=""50"" width=""25%"" ><img src=""" & GetSignaturePath(fourthApprovBy) & """ style=""border-bottom:1px dotted black;width:150px;height:130px;""</td>"
Else
    response.write "<td valign=""bottom"" height=""50"" width=""25%"" ><img src=""" & GetSignaturePath(thirdApprovBy) & """ style=""border-bottom:1px dotted black;width:150px;height:130px;""</td>"
End If

'Response.write "<td valign=""bottom"" height=""50"" width=""33%"">...........................................................</td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td valign=""top"" align=""center"" width=""25%"">" & GetStaffName(GetRecordField("SystemUserID")) & "</td>"
response.write "<td valign=""top"" align=""center"" width=""25%"">" & GetStaffName(firstApprovBy) & "</td>"
response.write "<td valign=""top"" align=""center"" width=""25%"">" & GetStaffName(secondApprovBy) & "</td>"
If Len(fourthApprovBy) > 0 Then
    response.write "<td valign=""top"" align=""center"" width=""25%"">" & GetStaffName(fourthApprovBy) & "</td>"
Else
    response.write "<td valign=""top"" align=""center"" width=""25%"">" & GetStaffName(thirdApprovBy) & "</td>"
End If

'response.write "<td valign=""top"" align=""center"" width=""33%"">Counter Signed</td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td valign=""bottom"" width=""25%"">Date : <span style=""border-bottom:1px dotted black;"">" & FormatDateDetail(GetRecordField("PurchaseOrderDate")) & "</span></td>"
response.write "<td valign=""bottom"" width=""25%"">Date : <span style=""border-bottom:1px dotted black;"">" & FormatDateDetail(firstApprovDt) & "</td>"
response.write "<td valign=""bottom"" width=""25%"">Date : <span style=""border-bottom:1px dotted black;"">" & FormatDateDetail(secondApprovDt) & "</td>"

If IsDate(fourthApprovDt) Then
    response.write "<td valign=""bottom"" width=""25%"">Date : <span style=""border-bottom:1px dotted black;"">" & FormatDateDetail(fourthApprovDt) & "</td>"
Else
    response.write "<td valign=""bottom"" width=""25%"">Date : <span style=""border-bottom:1px dotted black;"">" & FormatDateDetail(thirdApprovDt) & "</td>"
End If

'response.write "<td valign=""bottom"" width=""33%"">Date : ...........................................................</td>"
response.write "</tr>"

response.write "</table>"
response.write "</td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""3"" bgcolor=""White"" style=""border-collapse:collapse;font-size: 9pt; font-family: Arial"">"
response.write "<tr>"

response.write "<td width=""50%"">"
response.write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""3"" bgcolor=""White"" style=""border-collapse:collapse;font-size: 10pt; font-family: Arial"">"
response.write "<tr>"
response.write "<td valign=""top"">I CERTIFY that the above mentioned supplies have been received and taken on charge.</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td height=""30"" valign=""bottom""><b>Storekeeper Sign :</b> _________________________</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td height=""30"" valign=""bottom""><b>Date :</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;_____________________________</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"

response.write "<td width=""50%"" valign=""top"">"
response.write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""3"" bgcolor=""White"" style=""border-collapse:collapse;font-size: 10pt; font-family: Arial"">"
response.write "<tr>"
response.write "<td align=""right""><b>&nbsp;</b></td>"
response.write "</tr></table>"
response.write "</td>"
response.write "</tr>"

response.write "</table>"
response.write "</td>"
response.write "</tr>"
'NOTICE
response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""3"" bgcolor=""White"" style=""border-collapse:collapse;font-size: 9pt; font-family: Arial"">"
response.write "<tr>"
response.write "<td width=""100%"" valign=""top"">"
response.write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""3"" bgcolor=""White"" style=""border-collapse:collapse;font-size: 10pt; font-family: Arial"">"
response.write "<tr>"
response.write "<td align=""center""><b><u>NOTICE TO SUPPLIERS</u></b></td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""left""><b>1.</b> The hospital does not accept liability for any Order signed by an unauthorised person. "
response.write " Suppliers are therefore advised to satisfy themselves that the Order has been signed by a proper person.</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""left""><b>2.</b> A genuine Order should always be presented in Original and Duplicate. "
response.write " Suppliers are required to return the Original copy with a priced invoice together with the materials supplied.</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""center"">* To be filled in after supplies have been received</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"

response.write "</table>"
response.write "</td>"
response.write "</tr>"

Set rst = Nothing


Function GetSignaturePath(UserName)
    Dim ot, staffid
    
    staffid = GetComboNameFld("SystemUser", UserName, "StaffID")
    
    ot = "images/signatures/" & staffid & ".png"
    GetSignaturePath = ot
End Function
Function GetStaffName(UserName)
    GetStaffName = GetComboName("Staff", GetComboNameFld("SystemUser", UserName, "StaffID"))
End Function
Function GetDrugPurOrderProUser(drgPur, stg)
    Dim ot, sql, rst
    sql = "select top 1 * from DrugPurOrderPro where DrugPurOrderID='" & drgPur & "' and TransProcessVal2ID='" & stg & "' order by TransProcessDate1 desc "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        ot = Array(rst.fields("SystemUserID").value, rst.fields("TransProcessDate1").value)
    Else
        ot = Array()
    End If
    
    GetDrugPurOrderProUser = ot
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

response.write "<tr>"
response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table id=""tblHiddenFields"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table id=""tblFooter"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
response.write "<tr>"
response.write "<td colspan=""7"" bgcolor=""#FFFFFF"" height=""10"" style=""font-size: 8pt"" align=""right"">"
response.write "" & GetComboName("Branch", brnch) & ". Copyright@" & Year(Now) & "</td>" ' <br>Software by : Spagad Technologies</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
