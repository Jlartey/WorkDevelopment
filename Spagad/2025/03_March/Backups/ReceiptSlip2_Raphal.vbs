'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
'response.write Glob_GetBootstrap5()
'Dim rDt, nMin, sDt, nMax
'sDt = GetRecordField("ServiceNo")
'rDt = GetRecordField("ReceiptDate")
'If IsDate(sDt) Then
'  nMin = DateDiff("n", CDate(sDt), Now())
'  nMax = 15
'Else
'  nMin = DateDiff("n", CDate(rDt), Now())
'  nMax = 120
'End If
'If nMin < nMax Or True Then
'  ''Test Export
'  'SetPageVariable "AutoHidePrintControl", "1"
'  'InitPageScript
'  'InitExportParam "jpg", "0.1", "0.1", "0.1", "0.1"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center"">"
'  ResponseExport2 "<table border=""0"" width=""600"" cellspacing=""0"" cellpadding=""0"">"
'  ResponseExport2 "<tr height=""60"">"
'  ResponseExport2 "<td align=""Center"" valign=""top""></td>"
'  ResponseExport2 "<td align=""center"" height=""20"" bgcolor=""white"" style=""font-size: 14pt; color: #CCCCFF;  font-family:Arial"">"
'  'DisplayHeader
'  Glob_AddReportHeader
'  ResponseExport2 "</td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "</table>"
'  ResponseExport2 "</td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-size:12pt"">"
'  'ResponseExport2 "<div style=""border-bottom: 1px solid #c5b8b8; padding-bottom:4px 0"">"
'  ResponseExport2 "OFFICIAL RECEIPT [ORIGINAL]</td>"
'  'ResponseExport2 "<div>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center"" valign=""top"" height=""420"" >"
'  ResponseExport2 "<table border=""0"" width=""600"" cellspacing=""0"" cellpadding=""2"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial"">"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInplabrequestID"" id=""tdLabelInplabrequestID"">RECEIPT NO.</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInplabrequestID"" id=""tdInputInplabrequestID"" style=""font-family: Arial; color: #111111; font-size:11pt"" >" & (GetRecordField("ReceiptID")) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInplabrequestName"" id=""tdLabelInplabrequestName"">RECEIPT NAME</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td colspan=""4"" name=""tdInputInplabrequestName"" id=""tdInputInplabrequestName"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"">" & (GetRecordField("ReceiptName")) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"">ACCOUNT NO.</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">" & (UCase(GetRecordField("PatientID"))) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"">ACCOUNT TYPE</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  If (UCase((GetRecordField("PatientID"))) = "P1") Or (UCase((GetRecordField("PatientID"))) = "P2") Then
'    ResponseExport2 "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">Walk-In</td>"
'  ElseIf Len(GetComboName("Sponsor", (GetRecordField("PatientID")))) > 0 Then
'    ResponseExport2 "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">Company</td>"
'  Else
'    ResponseExport2 "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">Patient</td>"
'  End If
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"">PAYMENT TYPE</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & (GetRecordField("PaymentModeName")) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"">DATE /TIME </td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (FormatDateDetail(GetRecordField("ReceiptDate"))) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"">AMOUNT</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & FormatNumber(GetRecordField("ReceiptAmount1"), 2, , -1) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"">CURRENCY</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (GetRecordField("CurrencyTypeName")) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr height=""5"">"
'  ResponseExport2 "<td colspan=""8""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"">AMOUNT IN WORDS</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td colspan=""6"">" & (UCase(GetPaymentWord(GetRecordField("ReceiptAmount1")))) & "</td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr height=""5"">"
'  ResponseExport2 "<td colspan=""8""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"">DESCRIPTION</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td colspan=""6"">" & (GetRecordField("Remarks")) & "</td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr height=""15"">"
'  ResponseExport2 "<td colspan=""8""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpSystemUserID"" id=""tdLabelInpSystemUserID""></td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpSystemUserID"" id=""tdInputInpSystemUserID""></td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" align=""right"">CASHIER&nbsp;:&nbsp;</td>"
'  ResponseExport2 "<td width=""10"" ></td>"
'  ResponseExport2 "<td align=""center""><u>" & (Replace(GetComboName("Staff", GetComboNameFld("SystemUser", GetRecordField("SystemUserID"), "StaffID")), " ", "&nbsp;")) & "</u></td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpSystemUserID"" id=""tdLabelInpSystemUserID""></td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpSystemUserID"" id=""tdInputInpSystemUserID""></td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td></td>"
'  ResponseExport2 "<td width=""10"" ></td>"
'  'ResponseExport2 "<td align=""center"">FOR GENERAL MANAGER</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "</tr>"
'
'  'Closing Table
'  ResponseExport2 "</td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "</table>"
'
'  'Second Print
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center"">"
'  ResponseExport2 "<table border=""0"" width=""600"" cellspacing=""0"" cellpadding=""0"">"
'  ResponseExport2 "<tr height=""60"">"
'  ResponseExport2 "<td align=""Center"" valign=""top""></td>"
'  ResponseExport2 "<td align=""center"" height=""20"" bgcolor=""white"" style=""font-size: 14pt; color: #CCCCFF;  font-family:Arial"">"
'  'DisplayHeader
'  Glob_AddReportHeader
'  ResponseExport2 "</td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "</table>"
'  ResponseExport2 "</td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-size:12pt"">"
'  'ResponseExport2 "<div style=""border-bottom: 1px solid #c5b8b8; padding-bottom:4px"">"
'  ResponseExport2 "OFFICIAL RECEIPT [DUPLICATE]</td>"
'  'ResponseExport2 "<div>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center"" valign=""top"">"
'  ResponseExport2 "<table border=""0"" width=""600"" cellspacing=""0"" cellpadding=""2"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial"">"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInplabrequestID"" id=""tdLabelInplabrequestID"">RECEIPT NO.</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInplabrequestID"" id=""tdInputInplabrequestID"" style=""font-family: Arial; color: #111111"" >" & (GetRecordField("ReceiptID")) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInplabrequestName"" id=""tdLabelInplabrequestName"">RECEIPT NAME</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInplabrequestName"" id=""tdInputInplabrequestName"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"">" & (GetRecordField("ReceiptName")) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" >ACCOUNT NO.</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">" & (UCase(GetRecordField("PatientID"))) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"">ACCOUNT TYPE</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  If (UCase((GetRecordField("PatientID"))) = "P1") Or (UCase((GetRecordField("PatientID"))) = "P2") Then
'    ResponseExport2 "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">Walk-In</td>"
'  ElseIf Len(GetComboName("Sponsor", (GetRecordField("PatientID")))) > 0 Then
'    ResponseExport2 "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">Company</td>"
'  Else
'    ResponseExport2 "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">Patient</td>"
'  End If
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"">PAYMENT TYPE</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & (GetRecordField("PaymentModeName")) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"">DATE /TIME </td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (FormatDateDetail(GetRecordField("ReceiptDate"))) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"">AMOUNT</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & FormatNumber(GetRecordField("ReceiptAmount1"), 2, , -1) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"">CURRENCY</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (GetRecordField("CurrencyTypeName")) & "</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr height=""5"">"
'  ResponseExport2 "<td colspan=""8""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"">AMOUNT IN WORDS</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td colspan=""6"">" & (UCase(GetPaymentWord(GetRecordField("ReceiptAmount1")))) & "</td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr height=""5"">"
'  ResponseExport2 "<td colspan=""8""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"">DESCRIPTION</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td colspan=""6"">" & (GetRecordField("Remarks")) & "</td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr height=""15"">"
'  ResponseExport2 "<td colspan=""8""></td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpSystemUserID"" id=""tdLabelInpSystemUserID""></td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpSystemUserID"" id=""tdInputInpSystemUserID""></td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" align=""right"">CASHIER&nbsp;:&nbsp;</td>"
'  ResponseExport2 "<td width=""10"" ></td>"
'  ResponseExport2 "<td align=""center""><u>" & (Replace(GetComboName("Staff", GetComboNameFld("SystemUser", GetRecordField("SystemUserID"), "StaffID")), " ", "&nbsp;")) & "</u></td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""right"" name=""tdLabelInpSystemUserID"" id=""tdLabelInpSystemUserID""></td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td name=""tdInputInpSystemUserID"" id=""tdInputInpSystemUserID""></td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "<td></td>"
'  ResponseExport2 "<td width=""10"" ></td>"
'  'ResponseExport2 "<td align=""center"">FOR GENERAL MANAGER</td>"
'  ResponseExport2 "<td width=""10""></td>"
'  ResponseExport2 "</tr>"
'
'  'Closing Table
'  ResponseExport2 "</td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "</table>"
'End If 'nMin
'
'Sub DisplayHeader()
'  ResponseExport2 "<table border=""0"" width=""600"" cellspacing=""0"" cellpadding=""0"">"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""Center"" valign=""top"" width=""100""></td>"
'  ResponseExport2 "<td align=""center"" width=""500"">"
'
'  ResponseExport2 "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""1"">"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center"" style=""font-size: 12pt"" colspan=""6"">Hospital</td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center"" style=""font-size: 10pt"" colspan=""6"">PMB 16, MINISTRY, ACCRA</td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td style=""font-size: 8pt"">OSU&nbsp;</td>"
'  ResponseExport2 "<td style=""font-size: 8pt"">TEL:0302761974&#45;8</td>"
'
'  ResponseExport2 "<td style=""font-size: 8pt"">&nbsp;SPECIALIST&nbsp;</td>"
'  ResponseExport2 "<td style=""font-size: 8pt"">TEL:0302797147</td>"
'
'  ResponseExport2 "<td style=""font-size: 8pt"">&nbsp;MOTHER&nbsp;&&nbsp;CHILD&nbsp;</td>"
'  ResponseExport2 "<td style=""font-size: 8pt"">TEL:0302798290, 0231797953</td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td style=""font-size: 8pt"">TEMA&nbsp;</td>"
'  ResponseExport2 "<td style=""font-size: 8pt"">TEL:0303212992</td>"
'
'  ResponseExport2 "<td style=""font-size: 8pt"">&nbsp;LEGON&nbsp;</td>"
'  ResponseExport2 "<td style=""font-size: 8pt"">TEL:0236840627</td>"
'
'  ResponseExport2 "<td style=""font-size: 8pt"">&nbsp;LEGON&nbsp;</td>"
'  ResponseExport2 "<td style=""font-size: 8pt"">TEL:0303403861&#45;2</td>"
'
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td style=""font-size: 8pt"">DOME&nbsp;</td>"
'  ResponseExport2 "<td style=""font-size: 8pt"">TEL:0236840846</td>"
'
'
'  ResponseExport2 "<td style=""font-size: 8pt"">&nbsp;LEGON&nbsp;</td>"
'  ResponseExport2 "<td style=""font-size: 8pt"">TEL:0236839369</td>"
'
'  ResponseExport2 "<td style=""font-size: 8pt"">&nbsp;PENSION&nbsp;HOUSE&nbsp;</td>"
'  ResponseExport2 "<td style=""font-size: 8pt"">TEL:0236840942</td>"
'  ResponseExport2 "</tr>"
'
'  ResponseExport2 "<tr>"
'  ResponseExport2 "<td align=""center"" style=""font-size: 10pt"" colspan=""6"">WEB:&nbsp;Hospital.com, EMAIL:&nbsp;info@Hospital.com</td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "</table>"
'
'  ResponseExport2 "</td>"
'  ResponseExport2 "</tr>"
'  ResponseExport2 "</table>"
'End Sub
'Function GetPaymentWord(inAmt)
'Dim amt, fAmt, wAmt, ot
'ot = ""
'amt = Abs(CDbl(inAmt))
'wAmt = Int(amt)
'fAmt = Round(amt - wAmt, 2)
'ot = ot & GetAmountWord(wAmt) & " GHANA CEDI(S)"
'
'If fAmt > 0 Then
'ot = ot & " " & GetAmountWord(100 * fAmt) & " PESEWA(S)"
'End If
'GetPaymentWord = ot
'End Function
'Function GetAmountWord(inAmt)
' Dim amt, ot, amtRem, amtUnit
' amt = inAmt
' ot = ""
' If amt >= 1000000000 Then
'    amtUnit = "Billion"
'    ot = ot & " " & GetLess1000(Int(amt / 1000000000))
'    ot = ot & " " & amtUnit
'    amtRem = amt - (Int(amt / 1000000000) * 1000000000)
'  ElseIf amt >= 1000000 Then
'    amtUnit = "Million"
'    ot = ot & " " & GetLess1000(Int(amt / 1000000))
'    ot = ot & " " & amtUnit
'    amtRem = amt Mod 1000000
'  ElseIf amt >= 1000 Then
'    amtUnit = "Thousand"
'    ot = ot & " " & GetLess1000(Int(amt / 1000))
'    ot = ot & " " & amtUnit
'    amtRem = amt Mod 1000
'  Else
'    ot = ot & " " & GetLess1000(Int(amt / 1))
'    amtRem = 0
'  End If
'  If amtRem > 0 Then
'    ot = ot & " " & GetAmountWord(amtRem)
'  End If
'  GetAmountWord = ot
'End Function
'
'Function GetLess1000(Less1000)
'  Dim ot, Less1000Rem
'  ot = ""
'  If Less1000 >= 100 Then
'    ot = ot & " " & GetDigit(CStr(Int(Less1000 / 100)))
'    ot = ot & " Hundred"
'    Less1000Rem = Less1000 Mod 100
'    If Less1000Rem > 0 Then
'      ot = ot & " and"
'    End If
'  ElseIf Less1000 >= 10 Then
'    If Less1000 >= 10 And Less1000 <= 19 Then
'      Select Case Less1000
'        Case 10
'         ot = ot & "Ten"
'        Case 11
'         ot = ot & "Eleven"
'        Case 12
'         ot = ot & "Twelve"
'        Case 13
'         ot = ot & "Thirteen"
'        Case 14
'         ot = ot & "Fourteen"
'        Case 15
'          ot = ot & "Fifeteen"
'        Case 16
'          ot = ot & "Sixteen"
'        Case 17
'          ot = ot & "Seventeen"
'        Case 18
'          ot = ot & "Eighteen"
'        Case 19
'          ot = ot & "Nineteen"
'        Case Else
'
'      End Select
'      Less1000Rem = 0
'    Else
'      ot = ot & " " & GetTens(Int(Less1000 / 10))
'      Less1000Rem = Less1000 Mod 10
'    End If
'  ElseIf Less1000 < 10 Then
'    ot = ot & " " & GetDigit(CStr(Less1000))
'    Less1000Rem = 0
'  End If
'
'  If Less1000Rem > 0 Then
'    ot = ot & " " & GetLess1000(Less1000Rem)
'  End If
'  GetLess1000 = ot
'End Function
'
'
'Function GetTens(tens)
' Dim ot
'ot = ""
'  Select Case tens
'    Case 1
'
'    Case 2
'      ot = ot & "Twenty"
'    Case 3
'      ot = ot & "Thirty"
'    Case 4
'      ot = ot & "Forty"
'    Case 5
'      ot = ot & "Fifty"
'    Case 6
'      ot = ot & "Sixty"
'    Case 7
'      ot = ot & "Seventy"
'    Case 8
'      ot = ot & "Eighty"
'    Case 9
'      ot = ot & "Ninety"
'    Case Else
'  End Select
'GetTens = ot
'End Function
'
'Function GetDigit(digit)
'  Dim ot
'  ot = ""
'  Select Case digit
'    Case "0"
'     ot = "Zero"
'    Case "1"
'     ot = "One"
'    Case "2"
'     ot = "Two"
'    Case "3"
'      ot = "Three"
'    Case "4"
'      ot = "Four"
'    Case "5"
'      ot = "Five"
'    Case "6"
'      ot = "Six"
'    Case "7"
'      ot = "Seven"
'    Case "8"
'      ot = "Eight"
'    Case "9"
'      ot = "Nine"
'    Case "10"
'      ot = "Ten"
'    Case "11"
'      ot = "Eleven"
'    Case "12"
'      ot = "Twelve"
'    Case Else
'  End Select
'GetDigit = ot
'End Function
'
'Sub InitPageScript()
'  Dim htStr
'  htStr = ""
'  htStr = htStr & "<script id=""scptPrintLayoutExtraScript"" LANGUAGE=""javascript"">"
'  htStr = htStr & vbCrLf
'  htStr = htStr & "function PLExtraScriptOnLoad(){" & vbCrLf
'  htStr = htStr & "iFrmOnLoadAdjustHeight('iFrmExportStream');" & vbCrLf
'  htStr = htStr & "}" & vbCrLf
'  htStr = htStr & "</script>"
'  response.write htStr
'End Sub

Dim rDt, nMin, sDt, nMax
sDt = GetRecordField("ServiceNo")
rDt = GetRecordField("ReceiptDate")
rcpID = GetRecordField("ReceiptID")

If IsDate(sDt) Then
    nMin = DateDiff("n", CDate(sDt), Now())
    nMax = 15
Else
    nMin = DateDiff("n", CDate(rDt), Now())
    nMax = 120
End If
If nMin < nMax Or True Then
    StylesAdded

'    If (jSchd = "M06") Then
'        pharmacySlip rcpID
'    Else
        Printout
'    End If

End If

Sub Printout()

    response.write " <div style='width: 74mm; margin: auto' >"
'    response.write " <header class='header'>"
'    response.write "     <div class='first-head'>"
'    response.write "         <img src='images/banner1.bmp'/>"
'    'response.write "         <div>"
'    'response.write "             <h2>MEDIGRACE MEDICAL CENTER And PHARMARCY</h2>"
'    'response.write "             <p style='text-align: left;'>Tel: 0303 942 390 / 02029497773 | Email: medigrace.medical@gmail.com</p>"
'    'response.write "             <p style='text-align: left;'>Location: Tseddo near LADMA | GPS: GL-036-4495</p>"
'    'response.write "         </div>"
'    response.write "     </div>"
'    response.write "     <div class='first-head'>"
'    'response.write "         <img src='images/banner1.bmp'/>"
'    response.write "         <div>"
'    response.write "             <h2>MEDIGRACE MEDICAL CENTER And PHARMARCY</h2>"
'    response.write "             <p style='text-align: left;'>Tel: 0303 942 390 / 02029497773 | Email: medigrace.medical@gmail.com</p>"
'    response.write "             <p style='text-align: left;'>Location: Tseddo near LADMA | GPS: GL-036-4495</p>"
'    response.write "         </div>"
'    response.write "     </div>"
'    response.write " </header>"
    
    response.write " <main class='main' style='width: 70mm; margin: auto'>"
    response.write " <div class='header'>"
    response.write "     <div style='padding: 0.5rem 0.5rem;' class='first-head'>"
    response.write "         <img src='images/banner1.bmp'/>"
    'response.write "         <div>"
    'response.write "             <h2>MEDIGRACE MEDICAL CENTER And PHARMARCY</h2>"
    'response.write "             <p style='text-align: left;'>Tel: 0303 942 390 / 02029497773 | Email: medigrace.medical@gmail.com</p>"
    'response.write "             <p style='text-align: left;'>Location: Tseddo near LADMA | GPS: GL-036-4495</p>"
    'response.write "         </div>"
    response.write "     </div>"
    response.write "     <div class='first-head'>"
    'response.write "         <img src='images/banner1.bmp'/>"
    response.write "         <div style='flex-direction: column; padding: 0.5rem 0.5rem;'>"
    response.write "             <h2>FOCOS Orthopaedic Hospital</h2>"
    response.write "             <p style='text-align: left;'>Tel: +233 59 692 0909/1 | Email: info@focosgh.com</p>"
    response.write "             <p style='text-align: left;'>Location: No.8 Teshie Street, Pantang, Accra | GPS: GM-109-8032</p>"
    response.write "         </div>"
    response.write "     </div>"
    response.write " </div>"
    response.write "     <div style='padding: 0.5rem 0.5rem;'>"
    response.write "         <p style='text-align: left;'>Receipt No: <span>" & (GetRecordField("ReceiptID")) & "</span></p>"
    response.write "         <p>Date / Time: <span>" & (FormatDateDetail(GetRecordField("ReceiptDate"))) & "</span></p>"
    response.write "     </div>"
    response.write "     <table class='table'>"
    response.write "         <tr>"
    response.write "             <td>Received from: </td>"
    response.write "             <td><span>" & GetComboName("Patient", GetRecordField("PatientID")) & "</span></td>"
    response.write "         </tr>"
    response.write "         <tr>"
    response.write "             <td>The sum of:</td>"
    response.write "             <td>" & (UCase(GetPaymentWord(GetRecordField("ReceiptAmount1")))) & "</td>"
    response.write "         </tr>"
    response.write "         <tr>"
    response.write "             <td>Being: </td>"
    response.write "             <td><span>Payment For " & (GetRecordField("Remarks")) & "</span></td>"
    response.write "         </tr>"
    ' response.write "         <tr>"
    ' response.write "             <td>Cash/Cheque No:</td>"
    ' response.write "             <td></td>"
    ' response.write "         </tr>"
    ' response.write "         <tr>"
    ' response.write "             <td>Balance: </td>"
    ' response.write "             <td><span></span></td>"
    ' response.write "         </tr>"
    response.write "     </table>"
    response.write "     <div style='padding: 0.5rem 0.5rem;'>"
    response.write "         <p class='amount'>GH&#8373: " & FormatNumber(GetRecordField("ReceiptAmount1"), 2, , -1) & "</p>"
    'response.write "         <p class='signature'>Signature: ..............................................................................</p>"
    response.write "     </div>"
    response.write " </main>"
    response.write " <h6 class='signature'><Label>cashier: &nbsp;</Label>" & (Replace(GetComboName("Staff", GetComboNameFld("SystemUser", GetRecordField("SystemUserID"), "StaffID")), " ", "&nbsp;")) & "</h6>"
    response.write " </div>"
End Sub

Sub pharmacySlip(rcpID)

    response.write " <header class='header'>"
    response.write "     <div class='first-head'>"
    response.write "         <img src='images/banner1.bmp'/>"
    response.write "         <div>"
    response.write "             <h2>FOCOS Orthopaedic Hospital</h2>"
    response.write "             <p style='text-align: left;>Tel: +233 59 692 0909/1 | Email: info@focosgh.com</p>"
    response.write "             <p style='text-align: left;>Location: No.8 Teshie Street, Pantang, Accra | GPS: GM-109-8032</p>"
    response.write "         </div>"
    response.write "     </div>"
    response.write " </header>"
    response.write " <main class='main'>"
    response.write "     <div>"
    response.write "         <p style='text-align: left;'>Receipt No: <span>" & (GetRecordField("ReceiptID")) & "</span></p>"
    response.write "         <p>Date / Time: <span>" & (FormatDateDetail(GetRecordField("ReceiptDate"))) & "</span></p>"
    response.write "     </div>"
    response.write "     <table class='table'>"
    response.write "         <tr>"
    response.write "             <td><b>DESCRIPTION</b></td>"
    response.write "             <td><b>PRICE</b></td>"
    response.write "             <td><b>QTY</b></td>"
    response.write "             <td><b>AMOUNT</b></td>"
    response.write "         </tr>"
    response.write GetDrugs(rcpID)
    response.write "     </table>"
    response.write "     <div>"
    response.write "         <p class='amount'>Amount Paid: GH&#8373 " & FormatNumber(GetRecordField("ReceiptAmount1"), 2, , -1) & "</p>"
    response.write "         <p class='signature'>Signature: ...........................</p>"
    response.write "     </div>"
    response.write " </main>"
    response.write " <h6 class='signature'><Label>Attendant:</Label>" & (Replace(GetComboName("Staff", GetComboNameFld("SystemUser", GetRecordField("SystemUserID"), "StaffID")), " ", "&nbsp;")) & "</h6>"

End Sub

Sub StylesAdded()

    response.write " <style>"
    response.write "     body{"
    response.write "         display: grid;"
    response.write "         justify-content: center;"
    response.write "         font-family:sans-serif;"
    response.write "     }"
    response.write "     .header{"
    response.write "         display: flex;"
    response.write "         flex-direction: column;"
    response.write "         justify-content: center;"
    response.write "     }"
    response.write "     .header h2{"
    response.write "         margin: 2px 0px;"
    response.write "         font-size: larger;"
    response.write "         color: black;"
    response.write "     }"
    response.write "     .first-head{"
    response.write "         display: flex;"
    response.write "         justify-content: center;"
    response.write "         align-items: center;"
    response.write "         font-size: small;"
    response.write "         padding-bottom: 10px;"
    response.write "     }"
    response.write "     .first-head p{"
    response.write "         margin: 2px 0px;"
    response.write "     }"
    response.write "     .first-head img{"
    response.write "         height: 8vh;"
    response.write "         width: auto;"
    response.write "     }"
    response.write "     .main{"
    response.write "         position: relative;;"
    response.write "     }"
    response.write "     .main::after{"
    response.write "         content: '';"
    response.write "         display: block;"
    response.write "         position: absolute;"
    response.write "         top: 0;"
    response.write "         right: 0;"
    response.write "         bottom: 0;"
    response.write "         left: 0;"
    response.write "         background-image: url(images/banner1.bmp);"
    response.write "         opacity: 0.06; "
    response.write "         z-index: -1;"
    response.write "         background-position: center;"
    response.write "         background-size: 80%;"
    response.write "         background-repeat: no-repeat;"
    response.write "     }"
    response.write "     .main p{"
    response.write "         text-align: end;"
    response.write "         font-size: small;"
    response.write "     }"
    response.write "     .main span{"
    response.write "         font-weight: bold;"
    response.write "         font-size: medium;"
    response.write "         font-family: monospace;"
    response.write "     }"
    response.write "     h4{"
    response.write "         font-size: medium;"
    response.write "         width: 100%;"
    response.write "         padding: 6px 0px 6px 0px;"
    response.write "         background-color: gainsboro;"
    response.write "         text-align: center;"
    response.write "     }"
    response.write "    .table{"
    response.write "         font-family: monospace;"
    'response.write "         padding-left: 20px;"
    response.write "         width: 100%;"
    response.write "         font-size: large;"
    response.write "     }"
    response.write "     .table span{"
    response.write "         font-size: large;"
    response.write "         font-weight: normal;"
    response.write "     }"
    response.write "     .amount{"
    response.write "         font-size: x-large;"
    response.write "         font-weight: bold;"
    response.write "     }"
    response.write "     .signature{"
    response.write "         font-size: small;"
    response.write "         padding-top: 5vh;"
    response.write "     }"
    response.write "     .main div{"
    response.write "         display: flex;"
    'response.write "         padding-left: 20px;"
    response.write "     }"
    response.write " </style>"

End Sub

Function GetPaymentWord(inAmt)
    Dim amt, fAmt, wAmt, ot
    ot = ""
    amt = Abs(CDbl(inAmt))
    wAmt = Int(amt)
    fAmt = Round(amt - wAmt, 2)
    ot = ot & GetAmountWord(wAmt) & " GHANA CEDI(S)"

    If fAmt > 0 Then
        ot = ot & " " & GetAmountWord(100 * fAmt) & " PESEWA(S)"
    End If
    GetPaymentWord = ot
End Function

Function GetAmountWord(inAmt)
    Dim amt, ot, amtRem, amtUnit
    amt = inAmt
    ot = ""
    If amt >= 1000000000 Then
        amtUnit = "Billion"
        ot = ot & " " & GetLess1000(Int(amt / 1000000000))
        ot = ot & " " & amtUnit
        amtRem = amt - (Int(amt / 1000000000) * 1000000000)
    ElseIf amt >= 1000000 Then
        amtUnit = "Million"
        ot = ot & " " & GetLess1000(Int(amt / 1000000))
        ot = ot & " " & amtUnit
        amtRem = amt Mod 1000000
    ElseIf amt >= 1000 Then
        amtUnit = "Thousand"
        ot = ot & " " & GetLess1000(Int(amt / 1000))
        ot = ot & " " & amtUnit
        amtRem = amt Mod 1000
    Else
        ot = ot & " " & GetLess1000(Int(amt / 1))
        amtRem = 0
    End If
    If amtRem > 0 Then
        ot = ot & " " & GetAmountWord(amtRem)
    End If
    GetAmountWord = ot
End Function

Function GetLess1000(Less1000)
    Dim ot, Less1000Rem
    ot = ""
    If Less1000 >= 100 Then
        ot = ot & " " & GetDigit(CStr(Int(Less1000 / 100)))
        ot = ot & " Hundred"
        Less1000Rem = Less1000 Mod 100
        If Less1000Rem > 0 Then
            ot = ot & " And"
        End If
    ElseIf Less1000 >= 10 Then
        If Less1000 >= 10 And Less1000 <= 19 Then
            Select Case Less1000
             Case 10
                ot = ot & "Ten"
             Case 11
                ot = ot & "Eleven"
             Case 12
                ot = ot & "Twelve"
             Case 13
                ot = ot & "Thirteen"
             Case 14
                ot = ot & "Fourteen"
             Case 15
                ot = ot & "Fifeteen"
             Case 16
                ot = ot & "Sixteen"
             Case 17
                ot = ot & "Seventeen"
             Case 18
                ot = ot & "Eighteen"
             Case 19
                ot = ot & "Nineteen"
             Case Else

            End Select
            Less1000Rem = 0
        Else
            ot = ot & " " & GetTens(Int(Less1000 / 10))
            Less1000Rem = Less1000 Mod 10
        End If
    ElseIf Less1000 < 10 Then
        ot = ot & " " & GetDigit(CStr(Less1000))
        Less1000Rem = 0
    End If

    If Less1000Rem > 0 Then
        ot = ot & " " & GetLess1000(Less1000Rem)
    End If
    GetLess1000 = ot
End Function

Function GetTens(tens)
    Dim ot
    ot = ""
    Select Case tens
     Case 1

     Case 2
        ot = ot & "Twenty"
     Case 3
        ot = ot & "Thirty"
     Case 4
        ot = ot & "Forty"
     Case 5
        ot = ot & "Fifty"
     Case 6
        ot = ot & "Sixty"
     Case 7
        ot = ot & "Seventy"
     Case 8
        ot = ot & "Eighty"
     Case 9
        ot = ot & "Ninety"
     Case Else
    End Select
    GetTens = ot
End Function

Function GetDigit(digit)
    Dim ot
    ot = ""
    Select Case digit
     Case "0"
        ot = "Zero"
     Case "1"
        ot = "One"
     Case "2"
        ot = "Two"
     Case "3"
        ot = "Three"
     Case "4"
        ot = "Four"
     Case "5"
        ot = "Five"
     Case "6"
        ot = "Six"
     Case "7"
        ot = "Seven"
     Case "8"
        ot = "Eight"
     Case "9"
        ot = "Nine"
     Case "10"
        ot = "Ten"
     Case "11"
        ot = "Eleven"
     Case "12"
        ot = "Twelve"
     Case Else
    End Select
    GetDigit = ot
End Function

Function GetDrugs(rcpID)
    Dim rst, sql, fnlamt, html, testID

    Set rst = CreateObject("ADODB.Recordset")

    sql = " SELECT PatientFlag2.FlagInfo2 FROM Receipt"
    sql = sql & " LEFT JOIN PatientFlag2 ON Receipt.ReceiptInfo1 = PatientFlag2.PatientFlag2ID"
    sql = sql & " WHERE Receipt.ReceiptID = '" & rcpID & "'"

    html = " "
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        ItemList = Split(rst.fields("FlagInfo2"), "**")
        
        For Each item In ItemList
            details = Split(item, "||")
            
            If UBound(details) >= 2 Then
                drugTable = details(0)
                drugID = details(1)
                drugQty = details(2)
                
                html = html & GetDrugItems(drugTable, drugID, drugQty)
            End If
        Next
    End If

    rst.Close
    Set rst = Nothing
    GetDrugs = html
End Function

Function GetDrugItems(table, ky, qty)
    Dim rst, sql, fnlamt, html, testID

    Set rst = CreateObject("ADODB.Recordset")

    sql = " Select DrugID, qty, UnitCost, FinalAmt FROM " & table & " WHERE DrugID = '" & ky & "'"

    html = " "
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            testID = rst.fields("DrugID")
            testAmt = qty
            testCst = rst.fields("UnitCost")
            testFnl = rst.fields("FinalAmt")
            html = html & "<tr>"
            html = html & "<td>" & GetComboName("Drug", testID) & "</td> "
            html = html & "<td>" & testCst & "</td> "
            html = html & "<td>" & testAmt & "</td> "
            html = html & "<td>" & testFnl & "</td> "
            html = html & "</tr> "
            rst.MoveNext
        Loop
    End If

    rst.Close
    Set rst = Nothing
    GetDrugItems = html
End Function

          
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
'None
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
