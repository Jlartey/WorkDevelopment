'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim sql, rstPrn1, rstPrn2, cat, catNm, tot, vst, hdr, pos, gTot, gPaid, gUsed, hrf2, hrf3, hrf4
Dim admDDt, admID, hrf, pat, coRec, coAmt, coPay, copaySponsor, gTotSpon
Dim htStr, dt, sp, holdAge, allowCopay, amtAllowCopay, typ
server.scripttimeout = 1800
vst = Trim(GetRecordField("VisitationID"))
pat = Trim(GetRecordField("PatientID"))
dispTyp = Trim(Request.queryString("typ"))
addCss
addJS
tot = 0
gTot = 0
gTotSpon = 0
gPaid = 0
gUsed = 0
coPay = 0
hrf = "wpgConsultReview.asp?PageMode=AddNew&MedicalStaffID=012&PullupData=VisitationID||" & vst
hrf2 = "wpgVisitation.asp?PageMode=ProcessSelect&VisitationID=" & vst
hrf4 = "wpgVisitationPro.asp?PageMode=AddNew&TransactType=Cancel&PullupData=VisitationID||" & vst
hrf3 = "wpgPrtPrintLayoutAll.asp?PositionForTableName=Patient&PrintLayoutName=PatientMedicalHistory&PatientID=" & pat
admID = Trim(GetAdmissionID(vst))
admDDt = ""
If Len(admID) > 0 Then
  admDDt = GetComboNameFld("Admission", admID, "DischargeDate")
End If
allowCopay = False
amtAllowCopay = 0

'Client Script
htStr = ""
htStr = htStr & "<script id=""scptPrintLayoutExtraScript"" LANGUAGE=""javascript"">"
  htStr = htStr & vbCrLf
  htStr = htStr & "function PLExtraScriptOnLoad(){" & vbCrLf
  htStr = htStr & "}" & vbCrLf

htStr = htStr & "function ApproveClaim(vst){" & vbCrLf
htStr = htStr & "window.open ('wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&PrintLayoutName=ApproveClaim&VisitationID=" & vst & "');" & vbCrLf
htStr = htStr & "window.close();" & vbCrLf
htStr = htStr & "}" & vbCrLf

htStr = htStr & "function SelectClaim(vst){" & vbCrLf
htStr = htStr & "window.open ('wpgVisitation.asp?PageMode=ProcessSelect&VisitationID=" & vst & "', '_Blank');" & vbCrLf
htStr = htStr & "window.close();" & vbCrLf
htStr = htStr & "}" & vbCrLf

htStr = htStr & "function CancelClaim(vst){" & vbCrLf
htStr = htStr & "window.open ('wpgVisitationPro.asp?PageMode=AddNew&PullupData=VisitationID||" & vst & "');" & vbCrLf
htStr = htStr & "window.close();" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "</script>"
response.write htStr



''@bless - 22 Dec 2023 //Process copay for DrugSale and LabRequest items
If (Request.Form.count) >= 6 Then
    For Each inpId In Request.Form
        If UCase(inpId) = UCase("ProcessCoPay") Then
            If UCase(Request.Form(inpId)) = UCase("YES") Then
                allowCopay = True
            End If
        ElseIf UCase(inpId) = UCase("ProcessCoPayAmount") Then
            If IsNumeric(Request.Form(inpId)) Then
                amtAllowCopay = (Trim(Request.Form(inpId)))
            End If
        End If
    Next
Else
    ' response.write "count < 6"
End If
If allowCopay And IsNumeric(amtAllowCopay) And CDbl(amtAllowCopay) >= 0 Then
    tbl = Trim(Request.queryString("TableID"))
    recKy = Trim(Request.queryString("ItemID"))
    vst = Trim(Request.queryString("VisitationID"))
    If Len(tbl) > 1 And Len(recKy) > 0 And UCase(GetRecordField("VisitationID")) = UCase(vst) Then
        Glob_UpdateCopayTables tbl, recKy, vst
    Else
        ' response tbl & " :: " & recKy & " :: " & vst
    End If
Else
    ' response.write allowCopay & " :: " & amtAllowCopay
End If

' If UCase(Trim(Request.QueryString("ProcessCoPay"))) = UCase("Yes") Then
'     If IsNumeric(Trim(Request.QueryString("ProcessCoPayAmount"))) Then
'         amtAllowCopay = Trim(Request.QueryString("ProcessCoPayAmount"))
'         If CDbl(amtAllowCopay) > 0 Then
'             ' If UCase(jSchd)=UCase("SystemAdmin") Or UCase(jSchd)=UCase(uName) Or UCase(GetComboNameFld("SystemUser", uName, "StaffID"))=UCase("STF001") Then
'                 allowCopay = True
'                 tbl = Trim(Request.QueryString("TableID"))
'                 recKy = Trim(Request.QueryString("ItemID"))
'                 vst = Trim(Request.QueryString("VisitationID"))
'                 If Len(tbl) > 1 And Len(recKy) > 0 And UCase(GetRecordField("VisitationID")) = UCase(vst) Then
'                     Glob_UpdateCopayTables tbl, recKy, vst
'                 End If
'             ' End If
'         End If
'     End If
' End If


'Start Claim Form
response.write "<tr>"
response.write "<td  width=""" & (PrintWidth) & """>"
response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""page-break-after:always"">"

'response.write "<tr>"
'response.write "<td align=""center"">"
'response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""0"" cellpadding=""0"">"
'response.write "<tr>"
''AddReportHeader
'response.write "</tr>"
'response.write "</table>"
'response.write "</td>"
'response.write "</tr>"

'response.write "<tr>"
'response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'response.write "</tr>"

' response.write "<tr>"
' response.write "<td align=""center"" valign=""bottom"">" & CStr(PrintOutRecCnt) & "</td>"
' response.write "</tr>"

response.write "<tr>"
response.write "<td align=""center"" valign=""bottom"" class=""hidden-on-print"">CORPORATE/INSURANCE CLAIM VETTING FORM</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""center"" valign=""bottom"" class=""hidden-on-print""><b>CLAIM STATUS: " & UCase(GetRecordField("VisitModeName")) & "</b></td>"
response.write "</tr>"
response.write "<tr class='imgcenter'>"
response.write "<td class=""shown-on-print""><img border=""0"" src=""images/logo.jpg"" align=""center""></td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td>"

response.write "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" style=""font-size:11pt"">"
response.write "<tr height=""40"">"
response.write "<td align=""right"">" & GetComboName("Branch", brnch) & "&emsp;|&emsp;</td>"
response.write "<td align=""left"">" & UCase(GetRecordField("SponsorName")) & "</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td title=""Click to Approve Claim"" align=""center"" style=""font-weight:bold; color:#11ff11; cursor:hand; border:1px solid;"" onclick=""ApproveClaim('" & vst & "')"" class=""hidden-on-print"">APPROVE CLAIM</td>"
lnk = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationClaimB&PositionForTableName=Visitation&VisitationID=" & vst & "&typ=spn"
response.write "<td title=""Click to Get Sponsor Bill"" align=""center"" style=""font-weight:bold; color:Blue; cursor:hand; border:1px solid;""class=""hidden-on-print""><a data-href='" & lnk & "' onclick=""openPopup(this)"">SPONSOR BILL</a></td>"
response.write "<td ></td>"
response.write "<td title=""Click to Cancel Claim"" align=""center"" style=""color:#ff2222; cursor:hand; border:1px solid;"" onclick=""CancelClaim('" & vst & "')"" class=""hidden-on-print"">CANCEL CLAIM</td>"
response.write "<td ></td>"
response.write "<td ></td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""2"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 9pt; font-family: Arial"">"
response.write "<tr>"
response.write "<td name=""tdLabelInpInsuranceNo"" id=""tdLabelInpInsuranceNo"" style=""font-weight: bold"">Insurance No.</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpInsuranceNo"" id=""tdInputInpInsuranceNo"">" & (GetRecordField("InsuranceNo")) & "</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdLabelInpVisitationID"" id=""tdLabelInpVisitationID"" style=""font-weight: bold"">Claim No.</td>"
response.write "<td width=""20""></td>"
response.write "<td title=""Click to Edit Claim"" style=""color:#2222ff; cursor:hand"" onclick=""SelectClaim('" & vst & "')"">" & (GetRecordField("VisitationID")) & "&nbsp;[EDIT]</td>"
response.write "<td width=""20""></td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td name=""tdLabelInpVisitTypeID"" id=""tdLabelInpVisitTypeID"" style=""font-weight: bold"">Patient Name</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpVisitTypeID"" id=""tdInputInpID""><a target=""_Blank"" href=""" & hrf3 & """>" & (GetRecordField("PatientName")) & "</a></td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdLabelInpPatientID"" id=""tdLabelInpPatientID"" style=""font-weight: bold"">Patient No.</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpPatientID"" id=""tdInputInpPatientID"">" & (GetRecordField("PatientID")) & "</td>"
response.write "<td width=""20""></td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td name=""tdLabelInpPatientAge"" id=""tdLabelInpPatientAge"" style=""font-weight: bold"">Age</td>"
response.write "<td width=""20""></td>"
holdAge = Int(CInt(GetRecordField("PatientAge")))
response.write "<td name=""tdInputInpPatientAge"" id=""tdInputInpPatientAge"">" & CStr(Int(CInt(GetRecordField("PatientAge")))) & " [" & GetRecordField("VisitInfo6") & "] [" & FormatDate(GetComboNameFld("Patient", GetRecordField("PatientID"), "BirthDate")) & "]</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdLabelInpGenderID"" id=""tdLabelInpGenderID"" style=""font-weight: bold"">Gender</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & (GetRecordField("GenderName")) & "</td>"
response.write "<td width=""20""></td>"
response.write "</tr>"


response.write "<tr>"
response.write "<td name=""tdLabelInpVisitDate"" id=""tdLabelInpVisitDate"" style=""font-weight: bold"">Date</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpVisitDate"" id=""tdInputInpVisitDate"">" & (FormatDate(GetRecordField("VisitDate"))) & "</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdLabelInpInsuranceNo"" id=""tdLabelInpInsuranceNo"" style=""font-weight: bold"">Service Type</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpInsuranceNo"" id=""tdInputInpInsuranceNo"">" & (GetRecordField("MedicalServiceName")) & "</td>"
response.write "<td width=""20""></td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td name=""tdLabelInpInsuranceNo"" id=""tdLabelInpInsuranceNo"" style=""font-weight: bold"">Contact No.</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpInsuranceNo"" id=""tdInputInpInsuranceNo"">" & GetComboNameFld("Patient", GetRecordField("PatientID"), "ResidencePhone") & "</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdLabelInpVisitDate"" id=""tdLabelInpVisitDate"" style=""font-weight: bold"">Attending Doctor</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpVisitDate"" id=""tdInputInpVisitDate"">" & GetRecordField("SpecialistName") & "</td>"
response.write "<td width=""20""></td>"
response.write "</tr>"

DisplayAdmitDate vst

response.write "</table>"
response.write "</td>"
response.write "</tr>"



'Bill Details
response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""2"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 9pt; font-family: Arial"">"

response.write "<tr>"
response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1"" height=""10""></td>"
response.write "</tr>"

'Diagnosis Href
response.write "<tr>"
'response.write "<td></td>"
response.write "<td title=""Click to Add/Edit Diagnosis"" colspan=""8"" style=""font-size:9pt"" class='show-td'><b><a target=""_Blank"" href=""" & hrf & """>DIAGNOSIS :</a>  </b>" & (GetDiagnosis(vst)) & "</td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1"" height=""10""></td>"
response.write "</tr>"

If Len(Trim(GetRecordField("VisitInfo5"))) > 0 Then
response.write "<tr>"
response.write "<td colspan=""5""><b>COMMENTS :  </b>" & Trim(GetRecordField("VisitInfo5")) & "</td>"
response.write "</tr>"
End If
DisplayProcedure vst

If UCase(GetRecordField("VisitModeID")) <> "V006" Then ''On Admission
  response.write "<tr class=""hidden-on-print"">"
  response.write "<td style=""" & cellStyT & """ colspan='100'>"
  AddClaimVettingLinks
  response.write "</td>"
  response.write "</tr>"
End If

response.write "<tr>"
response.write "<td style=""font-weight: bold""><u>NO.</u></td>"
response.write "<td style=""font-weight: bold""><u>SERVICE DESCRIPTION</u></td>"
response.write "<td align=""right"" style=""font-weight: bold""><u>QTY  </u></td>"
response.write "<td align=""right"" style=""font-weight: bold""><u>UNIT  </u></td>"
response.write "<td align=""right"" style=""font-weight: bold""><u>  TOTAL</u></td>"
response.write "<td align=""right"" style=""font-weight: bold"" class=""hidden-on-print""><u> COPAY CASH </u></td>"
response.write "<td align=""right"" style=""font-weight: bold"" class=""hidden-on-print""><u> GRATIS </u></td>"
response.write "<td align=""right"" style=""font-weight: bold"" class=""hidden-on-print""><u> UNIT COST </u></td>"
response.write "<td align=""right"" style=""font-weight: bold"" class=""hidden-on-print""><u> COPAY SPONSOR </u></td>"
response.write "</tr>"

'Consultation
hdr = (GetRecordField("SpecialistTypeName")) & " [Consultation]"

If GetRecordField("VisitTypeID") = "V001" Then
  hdr = (GetRecordField("SpecialistTypeName")) & " [Registration & Consultation]"
  If UCase(GetRecordField("SponsorID")) = UCase("GCAA") Then
    hdr = (GetRecordField("SpecialistTypeName")) & " [Registration, File Charge, Documentation and Consultation]"
  End If
ElseIf GetRecordField("VisitTypeID") = "V002" Then
  hdr = (GetRecordField("SpecialistTypeName")) & " [Subsequent Consultation]"
End If

response.write "<tr>"
response.write "<td colspan=""8"" style=""font-weight: bold"" height=""10"" valign=""bottom"">    Consultation</td>"
response.write "</tr>"

gTot = gTot + GetRecordField("Visitcost")

  otDisc = GetRecordField("Visitcost2") ''Discount
  copayCash = GetRecordField("Visitcost3")
  copaySponsor = GetRecordField("Visitcost") - copayCash ''@bless - not implemented yet
  ' CoPaySponsor = GetRecordField("Visitcost")
  gTotSpon = gTotSpon + copaySponsor
response.write "<tr>"
response.write "<td>1.</td>"
response.write "<td>" & hdr & "</td>"
response.write "<td align=""right"">1</td>"
response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(GetRecordField("Visitcost")), 2, , , -1)) & "</td>"
response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(GetRecordField("Visitcost")), 2, , , -1)) & "</td>"
If IsNumeric(otDisc) And CDbl(otDisc) > 0 Then
  response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(otDisc), 2, , , -1)) & "</td>"
Else
  response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
End If
If IsNumeric(copayCash) And CDbl(copayCash) > 0 Then
  response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(copayCash), 2, , , -1)) & "</td>"
Else
  response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
End If
response.write "<td align=""right"">1</td>"
response.write "<td align=""right"">" & (FormatNumber(CStr(copaySponsor), 2, , , -1)) & "</td>"
response.write "</tr>"

pos = 1
AddAdmission vst
' AddDrug vst
  AddDrugCopay vst ''@bless - 15 Jan 2024
AddNonDrug vst
AddLab vst
'AddXray vst
AddTreat vst

response.write "<tr>"
response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
response.write "</tr>"

  ''@bless - 15 Jan 2024 //Process copay post values
  ' response.write "<tr class=""no-print"" style=""display:none;"">"
  response.write "<tr class=""no-print"" style=""display:--none--;"">"
  response.write "<td class=""no-print"" colspan=""5"" align=""center"">"
  response.write "<input type=""text"" style=""display:none;"" class=""no-print"" id=""ProcessCoPay""  name=""ProcessCoPay"" value=""YES"">"
  response.write "<input type=""number"" style=""display:none;"" class=""no-print"" id=""ProcessCoPayAmount""  name=""ProcessCoPayAmount"" value=""" & (gTotSpon - coPay) & """>"
  response.write "<input class=""btn btn-danger"" type=""submit"" onclick=""cmdSaveOnClick()"" id=""cmdSave""  name=""cmdSave"" value=""Process Co-Pay"">"
  response.write "</td>"
  response.write "</tr>"

'Grand Total
response.write "<tr>"
response.write "<td></td>"
response.write "<td colspan=""3"" align=""left"">TOTAL BILL</td>"
response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(gTot), 2, , , -1)) & "</td>"
response.write "<td colspan=""3"" align=""left"" class=""hidden-on-print""></td>"
response.write "<td align=""right""><b>" & (FormatNumber(CStr(gTotSpon), 2, , , -1)) & "</b></td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
response.write "</tr>"

' 'Add Co-Payment
' AddCoPayment
AddCoPaymentNew

''Payments
'AddPayments GetRecordField("PatientID")

''UsedPayments
'AddUsedPayments GetRecordField("PatientID")

If coPay > 0 Then
  'Grand Total
  response.write "<tr>"
  response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
  response.write "</tr>"

  response.write "<tr>"
  response.write "<td></td>"
  response.write "<td colspan=""3"" align=""left""><b>OUTSTANDING BILL</b></td>"
  response.write "<td align=""right"" class=""hidden-on-print""><b>" & (FormatNumber(CStr(gTot - coPay), 2, , , -1)) & "</b></td>"
  response.write "<td colspan=""3"" align=""left""></td>"
  response.write "<td align=""right""><b>" & (FormatNumber(CStr(gTotSpon - coPay), 2, , , -1)) & "</b></td>"
  response.write "</tr>"


  response.write "<tr>"
  response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
  response.write "</tr>"
End If

response.write "</table>"
response.write "</td>"
response.write "</tr>"

'End Claim Form

'====== Frank Give HR Option to send Client Bills to them Via SMS ==========
If (jSchd = "m23" Or jSchd = "M23") Then


response.write "<tr>"
response.write "<center>"
response.write "<td align=""center"">"
response.write "<br>"
    'Clickable Url Link
    tot = FormatNumber(CStr(gTotSpon), 2, , , -1)
    lnkCnt = lnkCnt + 501
    lnkID = "lnk" & CStr(lnkCnt)
    lnkText = "<b>&nbsp;&nbsp;Click To Send Bill To Client (sms) </b>"
    lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=LabSmsInterface&PositionForTableName=WorkingDay&WorkingDayID= &vst=" & vst & " &typ=claim&amt=" & tot & ""
    navPop = "POP"
    inout = "IN"
    fntSize = ""
    fntColor = "  #0000ff"
    bgColor = ""
    wdth = ""
    lnkCnt = lnkCnt + 1
    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
    response.write "</td>"
response.write "</center>"
response.write "</tr>"
End If



response.write "</table>"
response.write "</td>"
response.write "</tr>"

'AddCoPayment
Sub AddCoPayment()
  Dim coRec, coAmt, coRecVal, rTyp
  rTyp = Trim(GetRecordField("ReceiptTypeID"))
  coRec = Trim(GetRecordField("VisitInfo1"))
  coAmt = Trim(GetRecordField("VisitValue1"))
  coRecVal = GetReceiptValue(coRec)
  If UCase(Trim(rTyp)) = "R002" Then 'Credit
    If IsNumeric(coAmt) And IsNumeric(coRecVal) Then
      If (Round(CDbl(coAmt), 2) > 0) And (Round(CDbl(coRecVal), 2) >= Round(CDbl(coAmt), 2)) Then
        coPay = CDbl(coAmt) ''@bless - 04 Jan 2024
        response.write "<tr>"
        response.write "<td style=""font-weight: bold""> </td>"
        response.write "<td style=""font-weight: bold""><u>PAYMENT DESCRIPTION</u></td>"
        response.write "<td align=""right"" style=""font-weight: bold"">  </td>"
        response.write "<td align=""right"" style=""font-weight: bold"">  </td>"
        response.write "<td align=""right"" style=""font-weight: bold""><u>   AMT</u></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td> </td>"
        response.write "<td> Co-Payment</td>"
        response.write "<td align=""right"">  </td>"
        response.write "<td align=""right"">  </td>"
        response.write "<td align=""right"">-" & FormatNumber(coAmt, 2, , , -1) & "</td>"
        response.write "</tr>"
      End If
    End If
  End If
End Sub
Sub AddCoPaymentNew()
  Dim coRec, coAmt, coRecVal, rTyp
  rTyp = Trim(GetRecordField("ReceiptTypeID"))
  coRec = Trim(GetRecordField("VisitInfo1"))
  coAmt = Trim(GetRecordField("VisitValue1"))
  coRecVal = GetReceiptValue(coRec)
  If UCase(Trim(rTyp)) = "R002" Then 'Credit
    If IsNumeric(coAmt) And IsNumeric(coRecVal) Then
      ' If (Round(CDbl(coAmt), 2) > 0) And (Round(CDbl(coRecVal), 2) >= Round(CDbl(coAmt), 2)) Then
      If (Round(CDbl(coAmt), 2) > 0) Then
      '   ' coPay = CDbl(coAmt) ''@bless - 04 Jan 2024
        response.write "<tr class=""no-print"">"
        response.write "<td style=""font-weight: bold""> </td>"
        response.write "<td style=""font-weight: bold""><u>PAYMENT DESCRIPTION</u></td>"
        response.write "<td align=""right"" style=""font-weight: bold"">  </td>"
        response.write "<td align=""right"" style=""font-weight: bold"">  </td>"
        response.write "<td align=""right"" style=""font-weight: bold""><u>   AMT</u></td>"
        response.write "</tr>"

        response.write "<tr class=""no-print"">"
        response.write "<td> </td>"
        response.write "<td> Co-Payment</td>"
        response.write "<td align=""right"">  </td>"
        response.write "<td align=""right"">  </td>"
        response.write "<td align=""right"">(" & FormatNumber(coAmt, 2, , , -1) & ")</td>"
        response.write "</tr>"
      End If
    End If
  End If
End Sub
Function GetReceiptValue(coRec)
  Dim arr, ul, num, ot, rec, bal
  ot = 0
  arr = Split(coRec, ",")
  ul = UBound(arr)
  For num = 0 To ul
    rec = Trim(arr(num))
    bal = GetComboNameFld("Receipt", rec, "ReceiptAmount3")
    If IsNumeric(bal) Then
      ot = ot + CDbl(bal)
    End If
  Next
  GetReceiptValue = ot
End Function

'DisplayAdmitDate
Sub DisplayAdmitDate(vst)
  Dim rst, sql, ot, cnt, hdr, adm, chg, dys, aDt, dDt
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select * from admission where visitationid='" & vst & "' and admissionstatusid<>'A003' order by admissiondate"
  With rst
  .Open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  .MoveFirst
  'Admission
  cnt = 0
  Do While Not .EOF
  cnt = cnt + 1
  aDt = ""
  dDt = ""
  If Not IsNull(.fields("AdmissionDate")) Then
   If IsDate(.fields("AdmissionDate")) Then
     aDt = FormatDate(.fields("AdmissionDate"))
   End If
  End If

  If Not IsNull(.fields("DischargeDate")) Then
   If IsDate(.fields("DischargeDate")) Then
     dDt = FormatDate(.fields("DischargeDate"))
   End If
  End If

  response.write "<tr>"
  response.write "<td name=""tdLabelInpInsuranceNo"" id=""tdLabelInpInsuranceNo"" style=""font-weight: bold"">" & CStr(cnt) & ".  Admission Date</td>"
  response.write "<td width=""20""></td>"
  response.write "<td name=""tdInputInpInsuranceNo"" id=""tdInputInpInsuranceNo"">" & aDt & "</td>"
  response.write "<td width=""20""></td>"
  response.write "<td name=""tdLabelInpVisitDate"" id=""tdLabelInpVisitDate"" style=""font-weight: bold"">Discharge Date</td>"
  response.write "<td width=""20""></td>"
  response.write "<td name=""tdInputInpVisitDate"" id=""tdInputInpVisitDate"">" & dDt & "</td>"
  response.write "<td width=""20""></td>"
  response.write "</tr>"

  .MoveNext
  Loop

  End If
  .Close
  End With
  Set rst = Nothing
End Sub

'GetDiagnosis
Function GetDiagnosis(vst)
  Dim ot, rst, sql, cnt, nm, dt, lbTch, dia, crid, hrf, nm2
  Dim lst, arr, num, ul, did, fnd
  ot = ""
  sql = "select diagnosis.Diseaseid,Disease.Diseasename,diagnosis.ConsultReviewid "
  sql = sql & " from diagnosis,disease"
  sql = sql & " where diagnosis.diseaseid=disease.diseaseid and diagnosis.visitationid='" & vst & "'"
  sql = sql & " order by diagnosis.diseaseid"
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    .Open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      cnt = 0
      lst = ""
      Do While Not .EOF
        crid = .fields("ConsultReviewid")
        did = .fields("diseaseid")
        arr = Split(lst, "||")
        ul = UBound(arr)
        fnd = False
        For num = 0 To ul
          If UCase(Trim(arr(num))) = UCase(Trim(did)) Then
            fnd = True
            Exit For
          End If
        Next
        If Not fnd Then
          cnt = cnt + 1
          If cnt > 1 Then
            ot = ot & ", "
          End If
          lst = lst & "||" & did
          nm = .fields("diseasename")
          nm2 = Replace(nm, " ", " ")
          hrf = "wpgConsultReview.asp?PageMode=ProcessSelect&ConsultReviewID=" & crid
          ot = ot & "<a target=""_Blank"" href=""" & hrf & """ style=""border: 1px solid;border-radius: 8px;padding: 2px;text-decoration: none;line-height: 20px;""><nobr>" & nm2 & "</nobr></a>"
        End If
        .MoveNext
      Loop
    End If
    .Close

  End With


  If Trim(ot) = "" Then

    'pull diagnosis from labtests

    Dim rst3, sql3
    Set rst3 = server.CreateObject("ADODB.Recordset")

    sql3 = "select * from labrequest where visitationid = '" & vst & "' "

    With rst3
        .Open qryPro.FltQry(sql3), conn, 3, 4

        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
              '.fields("aFieldName")
                If Trim(.fields("clinicalDiagnosis")) <> "" Then
                  cnt = cnt + 1
                  If cnt > 1 Then
                    ot = ot & ", "
                  End If

                  hrf = "wpgLabRequest.asp?PageMode=ProcessSelect&LabRequestID=" & .fields("LabRequestID")
                  ot = ot & "<a target=""_Blank"" href=""" & hrf & """>" & Trim(.fields("clinicalDiagnosis")) & "</a>"

                End If

                .MoveNext
            Loop
        Else

        End If
    End With

  End If

  If Trim(ot) = "" Then
    'pull diagnosis from EMR

    Dim emrDiags
    emrDiags = getEMRDiagnosis(vst)
    ot = ot & emrDiags

  End If



  GetDiagnosis = ot
End Function


Function getEMRDiagnosis(vst)

  Dim ot
  Dim strDiags, ar1, ar2, ar3
  Dim i, x, y, dsName

  Dim rst, sql
  Set rst = CreateObject("ADODB.Recordset")

  ot = ""

  sql = ""
  sql = sql & "select er.column2 from emrresults as er "
  sql = sql & "left join emrrequest as eq on eq.emrrequestid = er.emrrequestid "
  sql = sql & "where eq.visitationid = '" & vst & "' and (er.emrcomponentid = 'TH06008' Or er.emrcomponentid = 'TH06008V') "

  strDiags = ""

  With rst
      .Open qryPro.FltQry(sql), conn, 3, 4

      If .RecordCount > 0 Then
          .MoveFirst
           If Not IsNull(.fields("column2")) Then
            strDiags = .fields("column2")
           End If
      End If
  End With

  If Trim(strDiags) <> "" Then
    ar1 = Split(strDiags, "~~")
    For i = 0 To UBound(ar1)
      If Trim(ar1(i)) <> "" Then
        ar2 = Split(Trim(ar1(i)), "03?%+*")
        If UBound(ar2) >= 2 Then

            If Trim(ar2(2)) <> "" Then
              ar3 = Split(Trim(ar2(2)), "||")
              For x = 0 To UBound(ar3)
                 If Trim(ar3(x)) <> "" Then
                    dsName = GetComboName("Disease", Trim(ar3(x)))
                    If dsName <> "" Then
                      ot = ot & " <u>" & dsName & "</u> &nbsp; "
                    End If
                 End If
              Next
            End If
        End If
      End If

    Next

  End If


  getEMRDiagnosis = ot

End Function

'DisplayProcedure
Sub DisplayProcedure(vst)
  Dim ot, rst, sql, cnt, nm, dt, lbTch, dia, crid, hrf, nm2
  Dim lst, arr, num, ul, did, fnd
  ot = ""
  sql = "select treatcharges.Treatmentid,Treatment.Treatmentname,treatcharges.ConsultReviewid "
  sql = sql & " from treatcharges,Treatment"
  sql = sql & " where treatcharges.Treatmentid=Treatment.Treatmentid and treatcharges.visitationid='" & vst & "'"
  sql = sql & " and Treatcharges.TreatGroupID<'014' and Treatcharges.Finalamt=0"
  sql = sql & " order by treatcharges.Treatmentid"
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    .Open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      cnt = 0
      lst = ""
      Do While Not .EOF
        crid = .fields("ConsultReviewid")
        did = .fields("Treatmentid")
        arr = Split(lst, "||")
        ul = UBound(arr)
        fnd = False
        For num = 0 To ul
          If UCase(Trim(arr(num))) = UCase(Trim(did)) Then
            fnd = True
            Exit For
          End If
        Next
        If Not fnd Then
          cnt = cnt + 1
          If cnt > 1 Then
            ot = ot & ", "
          End If
          lst = lst & "||" & did
          nm = .fields("Treatmentname")
          nm2 = Replace(nm, " ", " ")
          ot = ot & nm2
          'hrf = "wpgConsultReview.asp?PageMode=ProcessSelect&ConsultReviewID=" & crid
          'ot = ot & "<a target=""_Blank"" href=""" & hrf & """>" & nm2 & "</a>"
        End If
        .MoveNext
      Loop
      If cnt > 0 Then
        response.write "<tr>"
        response.write "<td colspan=""5"" style=""font-size:9pt""><b>PROCEDURE :  </b>" & ot & "</td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1"" height=""5""></td>"
        response.write "</tr>"
      End If
    End If
    .Close
  End With
End Sub

'AddPayments
Sub AddPayments(pat)
  Dim rst, sql, ot, cnt, hdr, dsc, pd, cn, bal, dt, sDt, eDt, vst, dDt, cn2
  Set rst = CreateObject("ADODB.Recordset")

  vst = GetRecordField("VisitationID")
  dDt = admDDt
  dt = GetComboNameFld("Visitation", vst, "VisitDate")
  sDt = FormatDate(dt) & " 0:00:00"
  eDt = Now()
  If Not IsNull(dDt) Then
    If IsDate(dDt) Then
      If CDate(dDt) > CDate(sDt) Then
        eDt = FormatDate(dDt) & " 23:59:59"
      End If
    End If
  End If

  sql = "select * from Receipt where Patientid='" & pat & "' and receiptdate between '" & sDt & "' and '" & eDt & "' order by receiptDate"
  cnt = 0
  With rst
    .Open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      'Receipt
      hdr = "Payments"
      response.write "<tr>"
      response.write "<td style=""font-weight: bold""><u>NO.</u></td>"
      response.write "<td style=""font-weight: bold""><u>PAYMENT DESCRIPTION</u></td>"
      response.write "<td align=""right"" style=""font-weight: bold""><u>PAID  </u></td>"
      response.write "<td align=""right"" style=""font-weight: bold""><u>REFUND  </u></td>"
      response.write "<td align=""right"" style=""font-weight: bold""><u>  BAL. AMT</u></td>"
      response.write "</tr>"

      Do While Not .EOF
        cnt = cnt + 1

        dsc = .fields("Remarks")
        pd = .fields("ReceiptAmount1")
        cn = .fields("paidamounT")
        If CDbl(cn) = 0 Then
          cn2 = "-"
        ElseIf CDbl(cn) < 0 Then
          cn = 0
          cn2 = "-"
        Else
          cn2 = FormatNumber(CStr(cn), 2, , , -1)
        End If
        bal = CDbl(pd) - CDbl(cn)
        gPaid = gPaid + bal
        response.write "<tr>"
        response.write "<td>" & CStr(cnt) & "</td>"
        response.write "<td>[REC# : " & UCase(.fields("ReceiptID")) & " ] " & dsc & "</td>"
        response.write "<td align=""right"">" & (FormatNumber(CStr(pd), 2, , , -1)) & "</td>"
        response.write "<td align=""right"">" & cn2 & "</td>"
        response.write "<td align=""right"">" & (FormatNumber(CStr(bal), 2, , , -1)) & "</td>"
        response.write "</tr>"
        .MoveNext
      Loop
      response.write "<tr>"
      response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
      response.write "</tr>"

      'Grand Total
      response.write "<tr>"
      response.write "<td></td>"
      response.write "<td colspan=""3"" align=""left"">TOTAL PAYMENT</td>"
      response.write "<td align=""right"">" & (FormatNumber(CStr(gPaid), 2, , , -1)) & "</td>"
      response.write "</tr>"

      response.write "<tr>"
      response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
      response.write "</tr>"
    End If
    .Close
  End With
  Set rst = Nothing
End Sub

'AddUsedPayments
Sub AddUsedPayments(pat)
  Dim rst, rst2, sql, ot, cnt, hdr, dsc, pd, cn, bal, dt, sDt, eDt, vst, dDt
  Dim cnt2, cn2, rec, usd, uCnt
  Set rst = CreateObject("ADODB.Recordset")
  Set rst2 = CreateObject("ADODB.Recordset")

  vst = GetRecordField("VisitationID")
  dDt = admDDt
  dt = GetComboNameFld("Visitation", vst, "VisitDate")
  sDt = FormatDate(dt) & " 0:00:00"
  eDt = Now()
  uCnt = 0
  If Not IsNull(dDt) Then
    If IsDate(dDt) Then
      If CDate(dDt) > CDate(sDt) Then
        eDt = FormatDate(dDt) & " 23:59:59"
      End If
    End If
  End If

  sql = "select * from Receipt where Patientid='" & pat & "' and receiptdate between '" & sDt & "' and '" & eDt & "' order by receiptDate"
  cnt = 0
  With rst
    .Open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst

      Do While Not .EOF
        cnt = cnt + 1
        rec = .fields("ReceiptID")
        dsc = ""
        pd = .fields("ReceiptAmount1")
        cn = .fields("paidamounT")
        If CDbl(cn) = 0 Then
          cn2 = "-"
        ElseIf CDbl(cn) < 0 Then
          cn = 0
          cn2 = "-"
        Else
          cn2 = FormatNumber(CStr(cn), 2, , , -1)
        End If
        bal = CDbl(pd) - CDbl(cn)
        usd = 0
        cnt2 = 0
        sql = "select * from PatientReceipt2 where Receiptid='" & rec & "' and VisitationID<>'" & vst & "' order by receiptDate2"
        rst2.Open qryPro.FltQry(sql), conn, 3, 4
        If rst2.RecordCount > 0 Then
          rst2.MoveFirst
          Do While Not rst2.EOF
            cnt2 = cnt2 + 1
            usd = usd + rst2.fields("PaidAmount")
            If cnt2 > 1 Then
              dsc = dsc & "; "
            End If
            dsc = dsc & "[V# : " & rst2.fields("VisitationID") & "] " & GetComboName("PaymentType", rst2.fields("PaymentTypeID"))
            rst2.MoveNext
          Loop
        End If
        rst2.Close
        gUsed = gUsed + usd
        If cnt2 > 0 Then
          uCnt = uCnt + 1
          If uCnt = 1 Then
            response.write "<tr>"
            response.write "<td style=""font-weight: bold""><u>NO.</u></td>"
            response.write "<td ><u><b>RECEIPT USED</b> [For Other Attendance/Visit]</u></td>"
            response.write "<td align=""right"" style=""font-weight: bold""><u>QTY  </u></td>"
            response.write "<td align=""right"" style=""font-weight: bold""><u>  </u></td>"
            response.write "<td align=""right"" style=""font-weight: bold""><u>  USED AMT</u></td>"
            response.write "</tr>"
          End If
          response.write "<tr>"
          response.write "<td>" & CStr(uCnt) & "</td>"
          response.write "<td>[REC# : " & UCase(.fields("ReceiptID")) & " ] " & dsc & "</td>"
          response.write "<td align=""right"">" & CStr(cnt2) & "</td>"
          response.write "<td align=""right"">-</td>"
          response.write "<td align=""right"">" & (FormatNumber(CStr(usd), 2, , , -1)) & "</td>"
          response.write "</tr>"
        End If
        .MoveNext
      Loop
      If uCnt > 0 Then
        response.write "<tr>"
        response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
        response.write "</tr>"

        'Grand Total
        response.write "<tr>"
        response.write "<td></td>"
        response.write "<td colspan=""3"" align=""left"">TOTAL RECEIPT USED</td>"
        response.write "<td align=""right"">" & (FormatNumber(CStr(gUsed), 2, , , -1)) & "</td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
        response.write "</tr>"
      End If
    End If
    .Close
  End With
  Set rst = Nothing
  Set rst2 = Nothing
End Sub

'AddAdmission
Sub AddAdmission_old(vst)
  Dim rst, sql, ot, cnt, hdr, adm, chg, dys
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select * from admission where visitationid='" & vst & "'"
  cnt = pos
  With rst
  .Open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  .MoveFirst
  'Admission
  hdr = "Admission"
  response.write "<tr>"
  response.write "<td colspan=""5"" style=""font-weight: bold"" height=""10"" valign=""bottom"">    " & hdr & "</td>"
  response.write "</tr>"

  Do While Not .EOF
  cnt = cnt + 1
  adm = .fields("admissionid")
  chg = 0 '.fields("bedcharge")
  dys = 0 '.fields("noofdays")
  gTot = gTot + (chg * dys)
  response.write "<tr>"
  response.write "<td>" & CStr(cnt) & "</td>"
  response.write "<td>" & GetComboName("Ward", .fields("wardid")) & " [" & GetComboName("AdmissionType", .fields("AdmissionTypeid")) & "]</td>"
  'response.write "<td align=""right"">" & CStr(dys) & "</td>"
  'response.write "<td align=""right"">" & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
  'response.write "<td align=""right"">" & (FormatNumber(CStr(chg * dys), 2, , , -1)) & "</td>"
  response.write "<td align=""right""> </td>"
  response.write "<td align=""right""> </td>"
  response.write "<td align=""right""> </td>"
  response.write "</tr>"
  .MoveNext
  Loop

  End If
  .Close
  End With
  Set rst = Nothing
  pos = cnt
End Sub

Sub AddAdmission(vst)
    Dim rst, sql, ot, cnt, hdr, adm, chg, dys, aDt, dDt, gADt, gDDt, dyCnt, gDyCnt, recCnt
    Dim bdTyp

    Set rst = CreateObject("ADODB.Recordset")
    sql = "select * from admission where visitationid='" & vst & "' and admissionstatusid<>'A003' order by admissiondate"
    cnt = pos
    Dim hourVIP
    strBillDesc = ""

    With rst
        .Open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .MoveFirst
            '23 Dec 2018
            recCnt = .RecordCount
            ''4 Dec 2018
            gADt = Now() 'Grand /Overall Admission Date
            gDDt = CDate("1 Jan 2017") 'Grand /Overall Dischage Date
            hourVIP = 0
            'Admission
            hdr = "Admission"
            hdr = "Hospitality [Bed and Ward Services]" '@bless - 23 Aug 2019'
            strBillDesc = strBillDesc & "<tr>"
            strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
            strBillDesc = strBillDesc & "</tr>"

            Do While Not .EOF
                cnt = cnt + 1
                aDt = Now() 'Current Admission Date
                dDt = Now() 'Current Dischage Date
                dtAdm = dDt
                dtDsch = dDt

                adm = .fields("admissionid")
                chg = .fields("bedcharge")
                dys = .fields("noofdays")
                BedTypeID = .fields("BedTypeID")

                If dys = 0 Then

                    ''4 Dec 2018 Individual Admissions
                    If IsDate(.fields("AdmissionDate")) Then
                        'If .fields("AdmissionDate") < aDt Then
                        aDt = .fields("AdmissionDate")
                        'End If
                    End If

                    If IsDate(.fields("DischargeDate")) Then
                        'If .fields("DischargeDate") > dDt Then
                        dDt = .fields("DischargeDate")
                        'End If
                    Else
                        dDt = Now()
                    End If

                    If (IsDate(aDt) And IsDate(dDt)) Then
                        dys = DateDiff("h", aDt, dDt)
                        dys = dys / 24
                        dyCnt = Int(dys)

                        If ((dys - dyCnt) > 0) Then
                            'If recCnt = 1 Then
                            dys = dyCnt + 1
                            'else
                            '  dys = dyCnt
                            'End If
                        Else
                            dys = dyCnt
                        End If
                    End If
                End If

                ''gTot = gTot + (chg * dys)

                ''4 Dec 2018 Over all Days
                If IsDate(.fields("AdmissionDate")) Then
                    If .fields("AdmissionDate") < gADt Then
                        gADt = .fields("AdmissionDate")
                    End If
                    dtAdm = .fields("AdmissionDate")
                End If

                If IsDate(.fields("DischargeDate")) Then
                    If .fields("DischargeDate") > gDDt Then
                        gDDt = .fields("DischargeDate")
                    End If
                    dtDsch = .fields("DischargeDate")
                Else
                    gDDt = Now()
                End If
                If UCase(.fields("BedTypeID")) = UCase("VIP") Then ''AKOSOMBO
                    hourVIP = hourVIP + DateDiff("h", dtAdm, dtDsch)
                    bdTyp = "VIP"
                ElseIf UCase(.fields("BedTypeID")) = UCase("ECO") Then ''ACCRA
                    hourVIP = hourVIP + DateDiff("h", dtAdm, dtDsch)
                    bdTyp = "ECO"
                End If

                'Compile Receipt No for Credit Patients
                'If UCase(rTyp) = "R002" Then 'Compile Receipt No for Credit Patients
                If Not IsNull(.fields("AdmissionInfo1")) Then
                    If Len(Trim(.fields("AdmissionInfo1"))) > 0 Then
                        recNo = recNo & "," & Trim(.fields("AdmissionInfo1"))
                    End If
                End If

                'End If

                strBillDesc = strBillDesc & "<tr>"
                strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
                strBillDesc = strBillDesc & "<td>" & GetComboName("Ward", .fields("wardid")) & " [" & GetComboName("AdmissionType", .fields("AdmissionTypeid")) & "]</td>"
                strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & CStr(dys) & " Days</td>"
                strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & (FormatNumber(CStr(chg * dys), 2, , , -1)) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
                strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
                strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
                strBillDesc = strBillDesc & "</tr>"
                .MoveNext
            Loop

            ''4 Dec 2018
            If (IsDate(gADt) And IsDate(gDDt)) Then
                gDys = DateDiff("h", gADt, gDDt)
                gDys = gDys / 24
                gDyCnt = Int(gDys)

                If ((gDys - gDyCnt) > 0) Then
                    gDys = gDyCnt + 1
                End If
            End If

            ' gTot = gTot + (chg * gDys)
            strBillDesc = strBillDesc & "<tr>"
            strBillDesc = strBillDesc & "<td>&nbsp;</td>"
            ' response.write "<td><b>Total Admissions</b></td>"
            strBillDesc = strBillDesc & "<td>Hospitality</td>" '@bless - 23 Aug 2019'
            strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & CStr(gDys) & " Days</td>"
            ' strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
            ' strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(chg * gDys), 2, , , -1)) & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
            strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
            strBillDesc = strBillDesc & "</tr>"

            ' ''VIP charges
            ' vipChg = 0
            ' vipDays = 0
            ' If hourVIP > 0 Then
            '     vipDays = hourVIP / 24
            '     vipDays = FormatNumber(vipDays + 0.49, 0) ''Ceiling
            '     vipChg = GetComboNameFld("BedType", bdTyp, "BedTypeCharge")
            '     If Not IsNumeric(vipChg) Or vipChg < 0 Then
            '         vipChg = 38
            '     End If
            '     gTot = gTot + (vipChg * vipDays)
            ' End If

            ' subTot = (chg * gDys) + (vipChg * vipDays)
            ' cnt = cnt + 1
            ' strBillDesc = strBillDesc & "<tr>"
            ' strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
            ' strBillDesc = strBillDesc & "<td>" & "VIP Hospitality" & "</td>"
            ' ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
            ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
            ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & CStr(dys) & " Days</td>"
            ' strBillDesc = strBillDesc & "<td align=""right"">" & FormatNumber(CStr(vipChg * vipDays), 2, , , -1) & "</td>"
            ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & (FormatNumber(CStr(chg * dys), 2, , , -1)) & "</td>"
            ' strBillDesc = strBillDesc & "</tr>"
        End If

        .Close
    End With
    response.write strBillDesc
    Set rst = Nothing
    pos = cnt
End Sub

Function GetAdmissionID(vst)
  Dim rst, sql, ot, cnt, adm
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select * from admission where visitationid='" & vst & "'"
  ot = ""
  With rst
  .Open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    .MoveFirst
    ot = .fields("admissionid")
  End If
  .Close
  End With
  GetAdmissionID = ot
  Set rst = Nothing
End Function

'AddDrugOLD
Sub AddDrugOLD(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, holdDrugName
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select drugid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from drugsaleitems where visitationid='" & vst & "' group by drugid"
  cnt = pos
  With rst
  .Open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  .MoveFirst
  'Pharmacy
  hdr = "Medical Items"
  response.write "<tr>"
  response.write "<td colspan=""5"" style=""font-weight: bold"" height=""10"" valign=""bottom"">    " & hdr & "</td>"
  response.write "</tr>"

  Do While Not .EOF
  cnt = cnt + 1
  unt = .fields("unt")
  qty = .fields("qty")
  tot = .fields("tot")
  gTot = gTot + tot
  response.write "<tr>"
  response.write "<td>" & CStr(cnt) & "</td>"

  holdDrugName = GetComboName("drug", .fields("drugid"))
  'response.write "<td>" & GetComboName("drug", .fields("drugid")) & "</td>"
  If InStr((UCase(holdDrugName)), "SYRUP") > 0 And Int(CInt(GetRecordField("PatientAge"))) > 15 Then
   response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
  ElseIf InStr((UCase(holdDrugName)), "TABLET") > 0 And Int(CInt(GetRecordField("PatientAge"))) <= 10 Then
   response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
  ElseIf InStr((UCase(holdDrugName)), "SUSPENSI0N") > 0 And Int(CInt(GetRecordField("PatientAge"))) > 15 Then
   response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
  ElseIf InStr((UCase(holdDrugName)), "CAPSULE") > 0 And Int(CInt(GetRecordField("PatientAge"))) < 10 Then
   response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
  ElseIf InStr((UCase(holdDrugName)), " SYR ") > 0 And Int(CInt(GetRecordField("PatientAge"))) > 15 Then
   response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
  ElseIf InStr((UCase(holdDrugName)), " TAB ") > 0 And Int(CInt(GetRecordField("PatientAge"))) <= 10 Then
   response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
  ElseIf InStr((UCase(holdDrugName)), " SUSP ") > 0 And Int(CInt(GetRecordField("PatientAge"))) > 15 Then
   response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
  ElseIf InStr((UCase(holdDrugName)), " CAPS ") > 0 And Int(CInt(GetRecordField("PatientAge"))) < 10 Then
   response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
  Else
   response.write "<td>" & holdDrugName & "</td>"
  End If

  'response.write "<td>" & holdDrugName & "</td>"
  response.write "<td align=""right"">" & CStr(qty) & "</td>"
  response.write "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
  response.write "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
  response.write "</tr>"
  .MoveNext
  Loop

  End If
  .Close
  End With
  Set rst = Nothing
  pos = cnt
End Sub

'GetReturnQty
Function GetReturnQty(vst, dg)
  Dim rstTblSql, sql, ot
  Set rstTblSql = CreateObject("ADODB.Recordset")
  ot = 0
  With rstTblSql

  'sql = "select sum(returnqty) as sm from drugreturnitems where visitationid='" & vst & "' and drugid='" & dg & "'"

  sql = "select sum(returnqty) as sm from ( "
      sql = sql & "select FinalAmt, returnqty from drugreturnitems where visitationid='" & vst & "' and drugid='" & dg & "' "
      sql = sql & "union all select FinalAmt, returnqty from drugreturnitems2 where visitationid='" & vst & "' and drugid='" & dg & "' "
  sql = sql & ") as t"

  .Open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  .MoveFirst
  ot = .fields("sm")
  If IsNull(ot) Then
  ot = 0
  End If
  End If
  .Close
  End With
  Set rstTblSql = Nothing
  GetReturnQty = ot
End Function

'AddDrug
Sub AddDrug(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, rQty, drg, fQty, tot2, rTot
  Set rst = CreateObject("ADODB.Recordset")

  'sql = "select drugid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from drugsaleitems where visitationid='" & vst & "' group by drugid"

  sql = "select drugid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from ("
  sql = sql & "select drugid, qty, unitcost, finalamt from drugsaleitems where visitationid='" & vst & "' "
  sql = sql & " union all "
  sql = sql & "select drugid, DispenseAmt1 as qty, unitcost, dispenseAmt2 as finalamt from drugsaleitems2 where visitationid='" & vst & "') as t "
  sql = sql & " group by drugid"

  cnt = pos
  With rst
  .Open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  .MoveFirst
  'Pharmacy
  hdr = "Medical Items"
  response.write "<tr>"
  response.write "<td colspan=""8"" style=""font-weight: bold"" height=""20"" valign=""bottom"">    " & hdr & "</td>"
  response.write "</tr>"

  Do While Not .EOF

  unt = .fields("unt")
  qty = .fields("qty")
  drg = .fields("drugid")

  tot2 = .fields("tot") 'Addedd 1 Oct 2015

  rQty = GetReturnQty(vst, drg)
  fQty = qty - rQty
  If fQty > 0 Then
    cnt = cnt + 1

    If rQty > 0 Then 'Addedd 1 Oct 2015
      rTot = GetReturnTot(vst, drg)
      unt = (tot2 - rTot) / fQty
    End If

    tot = fQty * unt '.Fields("tot")
    gTot = gTot + tot
    gTotSpon = gTotSpon + tot
    response.write "<tr>"
    response.write "<td>" & CStr(cnt) & "</td>"
    holdDrugName = GetComboName("drug", drg)
    If InStr((UCase(holdDrugName)), "SYRUP") > 0 And Int(CInt(GetRecordField("PatientAge"))) > 15 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    ElseIf InStr((UCase(holdDrugName)), "TABLET") > 0 And Int(CInt(GetRecordField("PatientAge"))) <= 10 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    ElseIf InStr((UCase(holdDrugName)), "SUSPENSI0N") > 0 And Int(CInt(GetRecordField("PatientAge"))) > 15 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
      ElseIf InStr((UCase(holdDrugName)), "CAPSULE") > 0 And Int(CInt(GetRecordField("PatientAge"))) < 10 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    ElseIf InStr((UCase(holdDrugName)), " SUSP ") > 0 And Int(CInt(GetRecordField("PatientAge"))) > 15 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    ElseIf InStr((UCase(holdDrugName)), " TAB ") > 0 And Int(CInt(GetRecordField("PatientAge"))) <= 10 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    ElseIf InStr((UCase(holdDrugName)), " SYR ") > 0 And Int(CInt(GetRecordField("PatientAge"))) > 15 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    ElseIf InStr((UCase(holdDrugName)), " CAPS ") > 0 And Int(CInt(GetRecordField("PatientAge"))) < 10 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    Else
      'response.write "<td>" & GetComboName("drug", drg) & "</td>"
      response.write "<td>" & holdDrugName & "</td>"
    End If
    'response.write "<td>" & GetComboName("drug", drg) & "</td>"
    response.write "<td align=""right"">" & CStr(fQty) & "</td>"
    response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
    response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
    response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
    response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
    response.write "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
    response.write "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
    response.write "</tr>"
  End If
  .MoveNext
  Loop

  End If
  .Close
  End With
  Set rst = Nothing
  pos = cnt
End Sub
Sub AddDrugCopay(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, untCP, qty, tot, rQty, drg, fQty, tot2, rTot
  Set rst = CreateObject("ADODB.Recordset")

  'sql = "select drugid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from drugsaleitems where visitationid='" & vst & "' group by drugid"

  sql = "select drugid,sum(qty) as qty, avg(unitcost) as unt, sum(finalamt) as tot, sum(CoPayCash) as CoPayCash, sum(CoPaySponsor) as CoPaySponsor from ("
  ' sql = sql & "select drugid, qty, unitcost, finalamt from drugsaleitems where visitationid='" & vst & "' "
  sql = sql & "select drugid, qty, unitcost, finalamt, (CAST(MainInfo1 AS FLOAT)) as CoPayCash, (CAST(MainItemInfo1 AS FLOAT)) as CoPaySponsor from drugsaleitems where visitationid='" & vst & "'"
  If (Len(dispTyp) > 0) Then
  sql = sql & " And CAST(MainItemInfo1 As Float) > 0 "
  End If
  sql = sql & " union all "
  ' sql = sql & "select drugid, DispenseAmt1 as qty, unitcost, dispenseAmt2 as finalamt from drugsaleitems2 where visitationid='" & vst & "') as t "
  sql = sql & "select drugid, dispenseamt1 as qty, unitcost, dispenseamt2 as finalamt, (CAST(MainInfo1 AS FLOAT)) as CoPayCash, (CAST(DispenseInfo1 AS FLOAT)) as CoPaySponsor from drugsaleitems2 where visitationid='" & vst & "' "
   If (Len(dispTyp) > 0) Then
  sql = sql & " And CAST(DispenseInfo1 As Float) > 0"
  End If
  sql = sql & " ) as t  group by drugid"
 
  cnt = pos
  With rst
  .Open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  .MoveFirst
  'Pharmacy
  hdr = "Medical Items"
  response.write "<tr>"
  response.write "<td colspan=""8"" style=""font-weight: bold"" height=""20"" valign=""bottom"">" & hdr & "</td>"
  response.write "</tr>"

  Do While Not .EOF

  unt = .fields("unt")
  qty = .fields("qty")
  drg = .fields("drugid")

  tot2 = .fields("tot") 'Addedd 1 Oct 2015
  copayCash = .fields("CoPayCash")
  copaySponsor = .fields("CoPaySponsor")

  rQty = GetReturnQty(vst, drg)
  fQty = qty - rQty
  untCP = unt
  If fQty > 0 Then
    cnt = cnt + 1

    If rQty > 0 Then 'Addedd 1 Oct 2015
      rTot = GetReturnTot(vst, drg)
      unt = (tot2 - rTot) / fQty
    End If

    If CDbl(copaySponsor) = 0 Then
      untCP = 0
    End If

    tot = fQty * unt '.Fields("tot")
    gTot = gTot + tot
    'gTotSpon = gTotSpon + tot '@bless/frank - 28 Nov 2024
    gTotSpon = gTotSpon + copaySponsor
    response.write "<tr>"
    response.write "<td>" & CStr(cnt) & "</td>"
    holdDrugName = GetComboName("drug", drg)
    If InStr((UCase(holdDrugName)), "SYRUP") > 0 And Int(CInt(GetRecordField("PatientAge"))) > 15 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    ElseIf InStr((UCase(holdDrugName)), "TABLET") > 0 And Int(CInt(GetRecordField("PatientAge"))) <= 10 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    ElseIf InStr((UCase(holdDrugName)), "SUSPENSI0N") > 0 And Int(CInt(GetRecordField("PatientAge"))) > 15 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
      ElseIf InStr((UCase(holdDrugName)), "CAPSULE") > 0 And Int(CInt(GetRecordField("PatientAge"))) < 10 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    ElseIf InStr((UCase(holdDrugName)), " SUSP ") > 0 And Int(CInt(GetRecordField("PatientAge"))) > 15 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    ElseIf InStr((UCase(holdDrugName)), " TAB ") > 0 And Int(CInt(GetRecordField("PatientAge"))) <= 10 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    ElseIf InStr((UCase(holdDrugName)), " SYR ") > 0 And Int(CInt(GetRecordField("PatientAge"))) > 15 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    ElseIf InStr((UCase(holdDrugName)), " CAPS ") > 0 And Int(CInt(GetRecordField("PatientAge"))) < 10 Then
    response.write "<td><font color='red'>" & holdDrugName & "</font></td>"
    Else
      'response.write "<td>" & GetComboName("drug", drg) & "</td>"
      response.write "<td>" & holdDrugName & "</td>"
    End If
    'response.write "<td>" & GetComboName("drug", drg) & "</td>"
    response.write "<td align=""right"">" & CStr(fQty) & "</td>"
    response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
    response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
    response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(copayCash), 2, , , -1)) & "</td>"
    response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
    response.write "<td align=""right"">" & (FormatNumber(CStr(untCP), 2, , , -1)) & "</td>"
    response.write "<td align=""right"">" & (FormatNumber(CStr(copaySponsor), 2, , , -1)) & "</td>"
                        ''@bless  //link for copay
                        If allowCopay Then
                            If (CDbl(amtAllowCopay) >= CDbl(copaySponsor)) Then
                                If (copaySponsor > 0) Then
                                    response.write "<td class=""no-print"" align=""left""><button class=""btn-info no-print"" style=""margin:3px;"" title=""Co pay " & drgNm & """ onclick=""PushToCopay('" & drg & "', 'DRUGSALE', '" & drgNm & " [" & FormatNumber(copaySponsor, 2) & "]" & "')"">[Copay]</button></td>"
                                Else
                                    ' ''If UCase(jSchd)=UCase(uName) Then
                                    ' If UCase(jSchd) = UCase("SystemAdmin") Or UCase(jSchd) = UCase(uName) Or UCase(GetComboNameFld("SystemUser", uName, "StaffID")) = UCase("STF001") Then
                                    response.write "<td class=""no-print"" align=""left""><button class=""btn-warning no-print"" style=""margin:3px;"" title=""Undo Co pay " & drgNm & """ onclick=""PushToCopay('" & drg & "', 'DRUGSALE-UNDO', '" & drgNm & " [" & FormatNumber(copaySponsor, 2) & "]" & "')"">[Undo Copay]</button></td>"
                                    ' End If
                                End If
                            End If
                        End If

    response.write "</tr>"
  End If
  .MoveNext
  Loop

  End If
  .Close
  End With
  Set rst = Nothing
  pos = cnt
End Sub

'GetReturnTot   'Addedd 1 Oct 2015
Function GetReturnTot(vst, dg)
  Dim rstTblSql, sql, ot
  Set rstTblSql = CreateObject("ADODB.Recordset")
  ot = 0
  With rstTblSql

  'sql = "select sum(FinalAmt) as sm from drugreturnitems where visitationid='" & vst & "' and drugid='" & dg & "'"

  sql = "select sum(finalamt) as sm from ( "
      sql = sql & "select FinalAmt, returnqty from drugreturnitems where visitationid='" & vst & "' and drugid='" & dg & "' "
      ' sql = sql & "union all select FinalAmt, returnqty from drugreturnitems2 where visitationid='" & vst & "' and drugid='" & dg & "' "
      sql = sql & "union all select MainItemValue1 as FinalAmt, returnqty from drugreturnitems2 where visitationid='" & vst & "' and drugid='" & dg & "' "
  sql = sql & ") as t"

  .Open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  .MoveFirst
  ot = .fields("sm")
  If IsNull(ot) Then
  ot = 0
  End If
  End If
  .Close
  End With
  Set rstTblSql = Nothing
  GetReturnTot = ot
End Function

Sub AddNonDrug(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select itemid,sum(qty) as qty,avg(retailunitcost) as unt,sum(finalamt) as tot from stockissueitems where visitationid='" & vst & "' group by itemid"
  cnt = pos
  With rst
  .Open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  .MoveFirst
  'Non Drug
  hdr = "Non Drug Consummables"
  response.write "<tr>"
  response.write "<td colspan=""8"" style=""font-weight: bold"" height=""10"" valign=""bottom"">    " & hdr & "</td>"
  response.write "</tr>"

  Do While Not .EOF
           cnt = cnt + 1
           itm = .fields("itemid")
           unt = .fields("unt")
           qty = .fields("qty")
           tot = .fields("tot")
           rQty = GetItemReturnQty(vst, itm)
           fQty = qty - rQty

   If fQty > 0 Then
       tot = fQty * unt '.Fields("tot")
       gTot = gTot + tot
       subTot = subTot + tot
       gTotSpon = gTotSpon + tot

          response.write "<tr>"
          response.write "<td>" & CStr(cnt) & "</td>"
          response.write "<td>" & GetComboName("items", .fields("itemid")) & "</td>"
          response.write "<td align=""right"">" & CStr(fQty) & "</td>"
          response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
          response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
          response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
          response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
          response.write "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
          response.write "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
          response.write "</tr>"
    End If
      .MoveNext
      Loop

    End If
  .Close
  End With
  Set rst = Nothing
  pos = cnt
End Sub


'GetItemReturnQty
Function GetItemReturnQty(vst, itm)
    Dim rst, sql, ot
    Set rst = CreateObject("ADODB.Recordset")

    sql = "select sum(FinalAmt) as amt, sum(returnqty) as qty from StockReturnItems where visitationid='" & vst & "' And ItemID='" & itm & "' "
    sql = sql & "  "
    ot = 0
    rst.Open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        If Not IsNull(rst.fields("qty")) Then
            ot = rst.fields("qty")
        End If
    End If
    rst.Close
    Set rst = Nothing
    GetItemReturnQty = ot
End Function

'AddLab
Sub AddLab(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot
  Set rst = CreateObject("ADODB.Recordset")

  ' 'sql = "select labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from investigation where visitationid='" & vst & "'"
  ' ''sql = sql & " and testcategoryid<>'T006' and testcategoryid<>'T007' and testcategoryid<>'T008' group by labtestid"
  ' 'sql = sql & " group by labtestid"

  ' sql = "select labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from ("
  ' sql = sql & "select labtestid, qty, unitcost, finalamt from investigation where visitationid='" & vst & "'"
  ' sql = sql & "union all "
  ' sql = sql & "select labtestid, qty, unitcost, finalamt from investigation2 where visitationid='" & vst & "') as t"
  sql = "select labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot, sum(ReceiptAmt1) as CoPayCash, sum(ReceiptAmt2) as CoPaySponsor from ("
  sql = sql & "select labtestid, qty, unitcost, finalamt, ReceiptAmt1, ReceiptAmt2 from investigation where visitationid='" & vst & "' "
  If (Len(dispTyp) > 0) Then
  sql = sql & " AND receiptamt2 > 0"
  End If
  sql = sql & "union all "
  sql = sql & "select labtestid, qty, unitcost, finalamt, ReceiptAmt1, ReceiptAmt2 from investigation2 where visitationid='" & vst & "'"
   If (Len(dispTyp) > 0) Then
  sql = sql & " AND receiptamt2 > 0"
  End If
  sql = sql & " ) as t"
  'sql = sql & " and testcategoryid<>'T006' and testcategoryid<>'T007' and testcategoryid<>'T008' group by labtestid"
  sql = sql & " group by labtestid"
   
  cnt = pos
  With rst
    .Open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      'Investigations
      hdr = "Investigations"
      response.write "<tr>"
      response.write "<td colspan=""5"" style=""font-weight: bold"" height=""10"" valign=""bottom"">    " & hdr & "</td>"
      response.write "</tr>"

      Do While Not .EOF
        cnt = cnt + 1
        unt = .fields("unt")
        qty = .fields("qty")
        tot = .fields("tot")
        gTot = gTot + tot
        ''@bless - 15 Jan 2024 //implement copay
        drg = .fields("LabTestID")
        copayCash = .fields("CoPayCash")
        copaySponsor = .fields("CoPaySponsor")
        untCP = unt
        drgNm = GetComboName("labtest", drg)
        gTotSpon = gTotSpon + copaySponsor
        response.write "<tr>"
        response.write "<td>" & CStr(cnt) & "</td>"
        response.write "<td>" & drgNm & "</td>"
        response.write "<td align=""right"">" & CStr(qty) & "</td>"
        response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
        response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
        ' response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
        response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(copayCash), 2, , , -1)) & "</td>"
        response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
        ' response.write "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
        ' response.write "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
        response.write "<td align=""right"" class="""">" & (FormatNumber(CStr(untCP), 2, , , -1)) & "</td>"
        response.write "<td align=""right"" class="""">" & (FormatNumber(CStr(copaySponsor), 2, , , -1)) & "</td>"
        ''@bless  //link for copay
        If allowCopay Then
            If (CDbl(amtAllowCopay) >= CDbl(copaySponsor)) Then
                If (copaySponsor > 0) Then
                    response.write "<td class=""no-print"" align=""left""><button class=""btn-info no-print"" style=""margin:3px;"" title=""Co pay " & drgNm & """ onclick=""PushToCopay('" & drg & "', 'LABREQUEST', '" & drgNm & " [" & FormatNumber(copaySponsor, 2) & "]" & "')"">[Copay]</button></td>"
                Else
                    ' ''If UCase(jSchd)=UCase(uName) Then
                   ' If UCase(jSchd) = UCase("SystemAdmin") Or UCase(jSchd) = UCase(uName) Or UCase(GetComboNameFld("SystemUser", uName, "StaffID")) = UCase("STF001") Then
                    response.write "<td class=""no-print"" align=""left""><button class=""btn-warning no-print"" style=""margin:3px;"" title=""Undo Co pay " & drgNm & """ onclick=""PushToCopay('" & drg & "', 'LABREQUEST-UNDO', '" & drgNm & " [" & FormatNumber(copaySponsor, 2) & "]" & "')"">[Undo Copay]</button></td>"
                   ' End If
                End If
            End If
        End If
        response.write "</tr>"
        .MoveNext
      Loop

    End If
    .Close
  End With
  Set rst = Nothing
  pos = cnt
End Sub

'AddXRay
Sub AddXray(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from investigation where visitationid='" & vst & "'"
  sql = sql & " and (testcategoryid='T006' or testcategoryid='T007' or testcategoryid='T008') group by labtestid"
  cnt = pos
  With rst
    .Open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      'X-Ray
      hdr = "X-Ray/Scan/ECG"
      response.write "<tr>"
      response.write "<td colspan=""5"" style=""font-weight: bold"" height=""10"" valign=""bottom"">    " & hdr & "</td>"
      response.write "</tr>"

      Do While Not .EOF
        cnt = cnt + 1
        unt = .fields("unt")
        qty = .fields("qty")
        tot = .fields("tot")
        gTot = gTot + tot
        response.write "<tr>"
        response.write "<td>" & CStr(cnt) & "</td>"
        response.write "<td>" & GetComboName("labtest", .fields("labtestid")) & "</td>"
        response.write "<td align=""right"">" & CStr(qty) & "</td>"
        response.write "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
        response.write "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
        response.write "</tr>"
        .MoveNext
    Loop

  End If
  .Close
  End With
  Set rst = Nothing
  pos = cnt
End Sub

'AddTreat
Sub AddTreat(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot
  Set rst = CreateObject("ADODB.Recordset")
  ' sql = "select treatmentid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from treatcharges where visitationid='" & vst & "' group by treatmentid"
  ''Inpatient/Hospitality charges
  sql = "select  TreatTypeID, sum(qty) as qty, avg(unitcost) as unt, Sum(FinalAmt) as tot, Sum(DiscAmt) as DiscAmt, "
  sql = sql & " sum(InitAmt) as InitAmt, Sum(MainValue2) as CopayCash, Sum(MainValue1) as CopaySponsor  "
  sql = sql & " from TreatCharges "
  sql = sql & " where VisitationID='" & vst & "'  And TreatTypeID='T008' " '' Inpatient "
  sql = sql & " And (qty > 0 or FinalAmt > 0) " ''@bless - 19 Dec 2023
  If (Len(dispTyp) > 0) Then
  sql = sql & " And Mainvalue1 > 0"
  End If
  sql = sql & " group by TreatTypeID "
  
  cnt = pos

  With rst
    .Open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      'treatment
      hdr = "Inpatient Accommodation"
      response.write "<tr>"
      response.write "<td colspan=""5"" style=""font-weight: bold"" height=""10"" valign=""bottom"">" & hdr & "</td>"
      response.write "</tr>"

      Do While Not .EOF
        cnt = cnt + 1
        unt = .fields("unt")
        qty = .fields("qty")
        tot = .fields("tot")
        copaySpon = .fields("CopaySponsor")
        disc = .fields("DiscAmt")
        copayCash = .fields("CopayCash")
        untSpn = copaySpon / qty
        gTot = gTot + tot
        gTotSpon = gTotSpon + copaySpon
        response.write "<tr>"
        response.write "<td>" & CStr(cnt) & "</td>"
        ' response.write "<td>" & GetComboName("treatment", .fields("treatmentid")) & "</td>"
        response.write "<td>" & GetComboName("TreatType", .fields("TreatTypeid")) & "</td>"
        response.write "<td align=""right"">" & CStr(qty) & "</td>"
        response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
        response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
        If copayCash > 0 Then
          response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(copayCash), 2, , , -1)) & "</td>"
        Else
          response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
        End If
        If disc > 0 Then
          response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(disc), 2, , , -1)) & "</td>"
        Else
          response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
        End If
        response.write "<td align=""right"">" & (FormatNumber(CStr(untSpn), 2, , , -1)) & "</td>"
        response.write "<td align=""right"">" & (FormatNumber(CStr(copaySpon), 2, , , -1)) & "</td>"
        response.write "</tr>"
        .MoveNext
      Loop

    End If
    .Close
  End With

  ''Non-Inpatient charges
  sql = "select TreatmentID, TreatTypeID, sum(qty) as qty, avg(unitcost) as unt, Sum(FinalAmt) as tot, Sum(DiscAmt) as DiscAmt, "
  sql = sql & " sum(InitAmt) as InitAmt, Sum(MainValue2) as CopayCash, Sum(MainValue1) as CopaySponsor  "
  sql = sql & " from TreatCharges "
  sql = sql & " where VisitationID='" & vst & "'  And TreatTypeID<>'T008' " ''Not Inpatient "
  sql = sql & " And (qty > 0 or FinalAmt > 0) " ''@bless - 19 Dec 2023
  If (Len(dispTyp) > 0) Then
  sql = sql & " And Mainvalue1 > 0"
  End If
  sql = sql & " group by TreatTypeID, TreatmentID "
  cnt = pos

  With rst
    .Open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      'treatment
      hdr = "Services and Consumables"
      response.write "<tr>"
      response.write "<td colspan=""5"" style=""font-weight: bold"" height=""10"" valign=""bottom"">" & hdr & "</td>"
      response.write "</tr>"

      Do While Not .EOF
        cnt = cnt + 1
        unt = .fields("unt")
        qty = .fields("qty")
        tot = .fields("tot")
        copaySpon = .fields("CopaySponsor")
        disc = .fields("DiscAmt")
        copayCash = .fields("CopayCash")
        untSpn = copaySpon / qty
        gTot = gTot + tot
        gTotSpon = gTotSpon + copaySpon
        response.write "<tr>"
        response.write "<td>" & CStr(cnt) & "</td>"
        response.write "<td>" & GetComboName("treatment", .fields("treatmentid")) & "</td>"
        response.write "<td align=""right"">" & CStr(qty) & "</td>"
        response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
        response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
        If copayCash > 0 Then
          response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(copayCash), 2, , , -1)) & "</td>"
        Else
          response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
        End If
        If disc > 0 Then
          response.write "<td align=""right"" class=""hidden-on-print"">" & (FormatNumber(CStr(disc), 2, , , -1)) & "</td>"
        Else
          response.write "<td align=""right"" class=""hidden-on-print"">" & "" & "</td>"
        End If
        response.write "<td align=""right"">" & (FormatNumber(CStr(untSpn), 2, , , -1)) & "</td>"
        response.write "<td align=""right"">" & (FormatNumber(CStr(copaySpon), 2, , , -1)) & "</td>"
        response.write "</tr>"
        .MoveNext
      Loop

    End If
    .Close
  End With

  Set rst = Nothing
  pos = cnt
End Sub


Sub AddClaimVettingLinks()
    Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, vCat
    vCat = GetRecordField("VisitCategoryID")

    response.write "<table border=""1"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse:collapse"" width=""100%""><tr>"
    ''response.write "<td class=""cpHdrTd"">Claim Vetting Links</td>"

    response.write "<td align=""center"">"
    'Clickable Url Link
    lnkCnt = lnkCnt + 1
    lnkID = "lnk" & CStr(lnkCnt)
    lnkText = "<b>&nbsp;&nbsp;Open Folder </b>"
    lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation&VisitationID=" & vst
    navPop = "POP"
    inout = "IN"
    fntSize = ""
    fntColor = "#ff4444"
    bgColor = ""
    wdth = ""
    lnkCnt = lnkCnt + 1
    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
    response.write "</td>"

    response.write "<td align=""center"">"
    'Clickable Url Link
    lnkCnt = lnkCnt + 1
    lnkID = "lnk" & CStr(lnkCnt)
    lnkText = "<b>&nbsp;&nbsp;Edit Visit </b>"
    lnkUrl = "wpgVisitation.asp?PageMode=ProcessSelect&VisitationID=" & vst
    navPop = "POP"
    inout = "IN"
    fntSize = ""
    fntColor = "#ff4444"
    bgColor = ""
    wdth = ""
    lnkCnt = lnkCnt + 1
    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
    response.write "</td>"

    response.write "<td align=""center"">"
    'Clickable Url Link
    lnkCnt = lnkCnt + 1
    lnkID = "lnk" & CStr(lnkCnt)
    lnkText = "<b>&nbsp;&nbsp;Add Diagnosis </b>"
    lnkUrl = "wpgConsultReview.asp?PageMode=AddNew&PullupData=VisitationID||" & vst
    navPop = "POP"
    inout = "IN"
    fntSize = ""
    fntColor = "#ff4444"
    bgColor = ""
    wdth = ""
    lnkCnt = lnkCnt + 1
    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
    response.write "</td>"

    response.write "<td align=""center"">"
    'Clickable Url Link
    lnkCnt = lnkCnt + 1
    lnkID = "lnk" & CStr(lnkCnt)
    lnkText = "<b>&nbsp;&nbsp;Process Bill </b>"
    lnkUrl = "wpgNavigateFrame.asp?FrameType=WorkFlow&PositionForTableName=Visitation&VisitationID=" & vst
    navPop = "POP"
    inout = "IN"
    fntSize = ""
    fntColor = "#ff4444"
    bgColor = ""
    wdth = ""
    lnkCnt = lnkCnt + 1
    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
    response.write "</td>"

    '  response.write "<td align=""center"">"
    '  'Clickable Url Link
    '  lnkCnt = lnkCnt + 1
    '  lnkID = "lnk" & CStr(lnkCnt)
    '  lnkText = "<b>&nbsp;&nbsp;Open Next Dept. Claim </b>"
    '  lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=ProcessNextClaim&VisitCat=" & vCat & "&ProNextClaimType=Dept&PositionForTableName=WorkingDay&WorkingDayID=DAY20190112&VisitationID=" & vst
    '  navPop = "NAV"
    '  inout = "IN"
    '  fntSize = ""
    '  fntColor = "#ff4444"
    '  bgColor = ""
    '  wdth = ""
    '  lnkCnt = lnkCnt + 1
    '  AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
    '  response.write "</td>"

    response.write "<td align=""center"">"
    'Clickable Url Link
    lnkCnt = lnkCnt + 1
    lnkID = "lnk" & CStr(lnkCnt)
    lnkText = "<b>&nbsp;&nbsp;Open Next Daily Claim </b>"
    lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=ProcessNextClaim&VisitCat=" & vCat & "&ProNextClaimType=Daily&PositionForTableName=WorkingDay&WorkingDayID=DAY20190112&VisitationID=" & vst
    navPop = "NAV"
    inout = "IN"
    fntSize = ""
    fntColor = "#ff4444"
    bgColor = ""
    wdth = ""
    lnkCnt = lnkCnt + 1
    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
    response.write "</td>"

    '  response.write "<td align=""center"">"
    '  'Clickable Url Link
    '  lnkCnt = lnkCnt + 1
    '  lnkID = "lnk" & CStr(lnkCnt)
    '  lnkText = "<b>&nbsp;&nbsp;Open Next Consult Type Claim </b>"
    '  lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=ProcessNextClaim&VisitCat=" & vCat & "&ProNextClaimType=ConsultType&PositionForTableName=WorkingDay&WorkingDayID=DAY20190112&VisitationID=" & vst
    '  navPop = "NAV"
    '  inout = "IN"
    '  fntSize = ""
    '  fntColor = "#ff4444"
    '  bgColor = ""
    '  wdth = ""
    '  lnkCnt = lnkCnt + 1
    '  AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
    '  response.write "</td>"

    ' If (UCase(GetRecordField("MedicalServiceID")) = "M003") Or (UCase(GetRecordField("MedicalServiceID")) = "M004") Then ' dmit/Detail
    '     response.write "<td align=""center"">"
    '     'Clickable Url Link
    '     lnkCnt = lnkCnt + 1
    '     lnkID = "lnk" & CStr(lnkCnt)
    '     lnkText = "<b>&nbsp;&nbsp;Open Admission Bill </b>"
    '     ' lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=Admission1&PositionForTableName=Admission&AdmissionID=" & GetAdmissionID3(vst)
    '     lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=Admission2&PositionForTableName=Admission&AdmissionID=" & GetAdmissionID3(vst)
    '     navPop = "POP"
    '     inout = "IN"
    '     fntSize = ""
    '     fntColor = "#ff4444"
    '     bgColor = ""
    '     wdth = ""
    '     lnkCnt = lnkCnt + 1
    '     AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
    '     response.write "</td>"
    ' End If

   If UCase(vCat) = "V011" Then
    response.write "<td align=""center"">"
    'Clickable Url Link
    lnkCnt = lnkCnt + 1
    lnkID = "lnk" & CStr(lnkCnt)
    lnkText = "<b>&nbsp;&nbsp;Vet Claim</b>"
    lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=ProcessNextClaim&VisitCat=" & vCat & "&ProNextClaimType=Daily&PositionForTableName=WorkingDay&WorkingDayID=DAY20190112&VisitationID=" & vst & "&Vetted=Yes"
    navPop = "NAV"
    inout = "IN"
    fntSize = ""
    fntColor = "#ff4444"
    bgColor = ""
    wdth = ""
    lnkCnt = lnkCnt + 1
    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
    response.write "</td>"
  End If

    response.write "</tr></table>"
End Sub



Function GetStockReceiptNo(vst)

    Dim rs, ot, sql, rNo, cnt

    Set rs = CreateObject("ADODB.Recordset")
    ot = ""
    cnt = 0
    sql = "select MainInfo1 from StockIssue where visitationid='" & vst & "'"

    With rs
        .Open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .MoveFirst

            Do While Not .EOF

                If Not IsNull(.fields("MainInfo1")) Then
                    rNo = Trim(.fields("MainInfo1"))

                    If Len(rNo) > 0 Then
                        ot = ot & "," & rNo
                    End If
                End If

                .MoveNext
            Loop

        End If

        .Close
    End With

    Set rs = Nothing
    GetStockReceiptNo = ot
End Function


Sub addJS()
    Dim js
    response.write "<link rel=""stylesheet"" type=""text/css"" href=""CSS/bootstrap.min.css"">"
    response.write "<script src=""Scripts/jquery-3.3.1.js"" language=""javascript""></script>"
    response.write "<script src=""Scripts/bootstrap.min.js"" language=""javascript""></script>"
    js = "<script language=""javascript"">"
    js = js & " this.document.title = """ & GetComboName("Visitation", vst) & """;" & vbNewLine
    ''Disable _clearEle();
    ' ' ' js = js & " window.print();" & vbNewline
    ' ' js = js & " document.getElementsByTagName('BODY')[0].onbeforeprint = function() {showBillSummary()}; " & vbNewline
    ' ' js = js & " function showBillSummary() { " & vbNewline
    ' ' js = js & "   var url; " & vbNewline
    ' ' js = js & "   url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=Admission1B&PositionForTableName=Admission&AdmissionID=" & adm & "';" & vbNewline
    ' ' js = js & "   window.location.href = url;" & vbNewline
    ' ' js = js & " };" & vbNewline

    ' js = js & "  function _clearEle() { " & vbNewLine
    ' js = js & "   var ele; " & vbNewLine
    ' js = js & "   ele = document.getElementById('trPrintControl'); " & vbNewLine
    ' js = js & "   if (ele) { var ele1 = ele.parentNode; if (ele1) {ele1.innerHTML='<font size=6 color=red>Policy Alert! You cannot print a PROVISIONAL bill</font>'; alert('PROVISIONAL Bill Cannot be printed');} } " & vbNewLine
    ' js = js & "  } " & vbNewLine
    ' ' js = js & "  window.onbeforeprint = _clearEle(); " & vbNewLine
    ' js = js & "  window.addEventListener(""beforeprint"", _clearEle); " & vbNewLine
    js = js & " function PushToCopay(drg, tbl, str) { " & vbNewLine
    js = js & "  if (confirm('UNDER CONSTRUCTION!\n\nAre you sure client should copay for ' + str + '?\n\nACTION CANNOT BE UNDONE!') == true) {  " & vbNewLine
    js = js & "   let url = ""wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationClaimB&PositionForTableName=Visitation""; " & vbNewLine
    ' js = js & "   url = url + ""&ProcessCoPay=YES&ProcessCoPayAmount=" & amtAllowCopay & """; " & vbNewLine
    js = js & "   url = url + ""&TableID="" + tbl + ""&VisitationID=" & vst & "&ItemID="" + drg; " & vbNewLine
    ' js = js & "   window.location.href = url;  " & vbNewLine
    js = js & "   var frm = document.getElementById('form1');frm.setAttribute('action',url);"
    js = js & "   frm.submit(); " & vbNewLine
    js = js & "   var i = document.getElementsByTagName(""button""); for(var n in i){ if(i[n]) {i[n].style.visibility = ""hidden""; i[n].style.display = ""none"";} };"
    js = js & "  }   " & vbNewLine
    js = js & " }   " & vbNewLine
    js = js & " function cmdSaveOnClick() {"
    js = js & "  alert('submit');"
    js = js & "  var ele = document.getElementById('cmdSave');ele.setAttribute('class','hidden');ele.setAttribute('disabled','disabled');"
    js = js & "  ;"
    js = js & "  document.getElementById('form1').submit();"
    js = js & " };"
    
  js = js & " function openPopup(anc){"
  js = js & "    let win=window.open(anc.dataset.href, '_blank', 'resizeable=yes,scrollbars=yes,width=820,height=560,status=yes  ');"
   js = js & "    "
   js = js & "   let intvl = setInterval(function(){"
   js = js & "        if(win.closed !== false){"
  js = js & "           clearInterval(intvl);"
  js = js & "            window.location.reload();"
  js = js & "         }"
  js = js & "     }, 200);"
  js = js & " }"
    
    
    js = js & " ;"
    js = js & " ;"

    js = js & "</script>" & vbNewLine
    response.write js
    ' js = "<script>"
    ' ' js = js & " alert('Bless');"
    ' js = js & "</script>"
    ' response.Write js
End Sub

Sub addCss()
    Dim str
    str = ""
    str = str & "<style>"
    str = str & "@media print{ "
    str = str & " #trPrintControl, .no-print {display:none !important;} "
    str = str & "} "
    str = str & " .sub-total { border-top: 1px solid #a2a2a2; border-bottom: 1px solid #a2a2a2; font-weight: 600; } "
    str = str & " .grand-total { border-top: 1px solid #0e0e0e; border-bottom: 1px solid #0e0e0e; font-weight: 800; } "
    str = str & "</style>"
    response.write str

    response.write "  <style>"
    response.write "    .shown-on-print{"
    response.write "      display: none;"
    response.write "    }"
    response.write "    @media print {"
    response.write "      .hidden-on-print {"
    response.write "        display: none;"
    response.write "      }"
    response.write "      a::after {"
    response.write "        content: none !important;"
    response.write "      }"
    response.write "      .shown-on-print{"
    response.write "        display: block;"
    response.write "      }"
    response.write ".imgcenter {"
    response.write "  display: flex;"
    response.write "  justify-content: center;"
    response.write "}"
    response.write "    }"
    response.write "  </style>"
End Sub

'response.write "<tr>"
'response.write "<tr>"
'response.write "<td align=""center"">"
'response.write "<table id=""tblHiddenFields"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
'response.write "<tr>"
'response.write "<td align=""center"">"
'response.write "<table id=""tblFooter"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
'response.write "<tr>"
'response.write "<td colspan=""7"" bgcolor=""#FFFFFF"" height=""10"" style=""font-size: 8pt"" align=""right"">"
'response.write "Hospital @2013</td>"
'response.write "</tr>"
'response.write "</table>"
'response.write "</td>"
'response.write "</tr>"
'response.write "</table>"
'response.write "</td>"
'response.write "</tr>"
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
