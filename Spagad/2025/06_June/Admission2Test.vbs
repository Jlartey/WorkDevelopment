'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim sql, rstPrn1, rstPrn2, cat, catNm, tot, vst, hdr, pos, gTot, gPaid, gUsed, recNo, pat, gWv, gDys, admUrl

Dim withReg, wd, insTyp, gIns, insWv, otTot, rTyp, wvPer, adm

Dim cmptdBy, cmptdDt, cmptdSig, vldPrint, hasDiscount, gTotDisc, gTotSpon, gTotCash

Dim allowCopay, amtAllowCopay, amtInvoice, IsCancelledAdmission, totCoPay

' response.write "<font size=""8"" color=""red"">REPORT IS UNDER MAINTENANCE</font><br>"
vldPrint = True
IsCancelledAdmission = False
addStat = UCase(GetRecordField("AdmissionStatusID"))
If UCase(addStat) <> "A002" And UCase(addStat) <> "A004" And UCase(addStat) <> "A008" Then
    response.write "<font size=""8"" color=""red"">PROVISIONAL BILL</font><br>"
    SetPageVariable "AutoHidePrintControl", "1"
    vldPrint = False
End If
allowCopay = False
amtAllowCopay = 0
totCoPay = 0


''Redirect for sub encounters
' If InStr(1, GetRecordField("VisitationID"), "-C") Then
If InStr(1, GetRecordField("VisitationID"), "-") > 1 Then
    adm = GetRecordField("VisitationID")
    Set rst = CreateObject("ADODB.Recordset")
    arr = Split(GetRecordField("VisitationID"), "-")
    If UBound(arr) >= 0 Then
        vst = arr(0)
        ' sql = "select top 1 * from Admission Where VisitationID='" & vst & "' "
        sql = "select top 1 * from Admission Where VisitationID='" & vst & "' And AdmissionStatusID<>'A003' order by AdmissionDate Desc " ''05 Jul 2024 @bless
        sql = sql & " "
        With rst
            rst.open qryPro.FltQry(sql), conn, 3, 4
            If rst.RecordCount > 0 Then
                adm = rst.fields("AdmissionID")
            End If
            rst.Close
        End With
    End If
    Set rst = Nothing
    response.redirect "wpgPrtPrintLayoutAll.asp?PrintLayoutName=Admission2&PositionForTableName=Admission&AdmissionID=" & adm
End If

''@bless - 22 Dec 2023 //Process copay for DrugSale and LabRequest items
If UCase(Trim(Request.queryString("ProcessCoPay"))) = UCase("Yes") Then
    If IsNumeric(Trim(Request.queryString("ProcessCoPayAmount"))) Then
        amtAllowCopay = Trim(Request.queryString("ProcessCoPayAmount"))
        If CDbl(amtAllowCopay) > 0 Then
            ' If UCase(jSchd)=UCase("SystemAdmin") Or UCase(jSchd)=UCase(uName) Or UCase(GetComboNameFld("SystemUser", uName, "StaffID"))=UCase("STF001") Then
                allowCopay = True
                tbl = Trim(Request.queryString("TableID"))
                recKy = Trim(Request.queryString("ItemID"))
                vst = Trim(Request.queryString("VisitationID"))
                If Len(tbl) > 1 And Len(recKy) > 0 And UCase(GetRecordField("VisitationID")) = UCase(vst) Then
                    UpdateCopayTables tbl, recKy, vst
                End If
            ' End If
        End If
    End If
End If

Dim revBy, revDt, revSig

Set rstPrn1 = CreateObject("ADODB.Recordset")
Set rstPrn2 = CreateObject("ADODB.Recordset")
cmptdBy = "<font style=""color:red"">Bill Not Completed</font>"
cmptdDt = "<font style=""color:red"">Bill Not Completed</font>"
cmptdSig = "<font style=""color:red"">Bill Un-Signed</font>"

revBy = "<font style=""color:red"">Bill Not Reviewed</font>"
revDt = "<font style=""color:red"">Bill Not Reviewed</font>"
revSig = "<font style=""color:red"">Bill Un-Signed</font>"

'@bless - 07 July 2020 >> Get the data for insert into a table for loggin of patient bill history
'@bless - Bill history is setup in logPatientBill >> TestVar96
Dim strBillDesc, strBillAmt, strBillOust, strBillCompltBy, strBillRevwBy
strBillDesc = ""
strBillAmt = ""
strBillOust = ""
strBillCompltBy = ""
strBillRevwBy = ""

pos = 0
tot = 0
gTot = 0
gPaid = 0
gUsed = 0
gWv = 0
gIns = 0
insWv = 0
wvPer = 0

gTotDisc = 0
gTotSpon = 0
gTotCash = 0

withReg = False
vst = Trim(GetRecordField("VisitationID"))
pat = Trim(GetRecordField("PatientID"))
recNo = GetAdmitReceiptNo(vst)
wd = Trim(GetRecordField("WardID"))
adm = Trim(GetRecordField("AdmissionID"))
sql = GetTableSql("Visitation")
sql = sql & " and  Visitation.Visitationid='" & Trim(vst) & "'"
addCss
insGrp = ""
hasDiscount = VisitHasDicount(vst)


With rstPrn1
    .open qryPro.FltQry(sql), conn, 3, 4

    If .RecordCount > 0 Then
        .movefirst
        insTyp = Trim(.fields("InsuranceTypeID"))
        rTyp = Trim(.fields("ReceiptTypeID"))
        insGrp = Trim(.fields("InsuranceGroupID"))

        'Compile Receipt No for Credit Patients
        If UCase(rTyp) = "R002" Then 'Compile Receipt No for Credit Patients
            If Not IsNull(.fields("VisitInfo1")) Then
                If Len(Trim(.fields("VisitInfo1"))) > 0 Then
                    recNo = recNo & "," & Trim(.fields("VisitInfo1"))
                End If
            End If
        End If

        'response.write "<tr>"
        'response.write "<td align=""center"">"
        'response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""0"" cellpadding=""0"">"
        'response.write "<tr>"
        'response.write "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"">"
        'response.write "<u>RESTRICTED</u></td>"
        'response.write "</tr>"

        ' AddBillReportHeader

        ''response.write "</td>"
        ''response.write "</tr>"
        strBillDesc = strBillDesc & "<tr>"
        strBillDesc = strBillDesc & "<td align=""center""><hr color=""#999999"" size=""1""></td>"
        strBillDesc = strBillDesc & "</tr>"

        strBillDesc = strBillDesc & "<tr height=""60"">"
        strBillDesc = strBillDesc & "<td align=""center"" colspan=""6"" height=""20"" bgcolor=""white"" style=""font-size: 14pt; color: #CCCCFF;  font-family:Arial"">"
        strBillDesc = strBillDesc & Glob_DisplayHeader1(adm, "OFFICIAL IN-PATIENT INVOICE", FormatDateDetail(GetRecordField("DischargeDate")))
        strBillDesc = strBillDesc & "</td>"
        strBillDesc = strBillDesc & "</tr>"

        strBillDesc = strBillDesc & "<tr>"
        strBillDesc = strBillDesc & "<td align=""center""><hr color=""#999999"" size=""1""></td>"
        strBillDesc = strBillDesc & "</tr>"
        strBillDesc = strBillDesc & "<tr>"
        strBillDesc = strBillDesc & "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:14pt"">"
        strBillDesc = strBillDesc & GetRecordField("BranchName") & "</td>"
        strBillDesc = strBillDesc & "</tr>"
        strBillDesc = strBillDesc & "<tr>"
        strBillDesc = strBillDesc & "<td align=""center"">"
        strBillDesc = strBillDesc & "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""2"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 9pt; font-family: Arial"">"
        strBillDesc = strBillDesc & "       <tr>"
        strBillDesc = strBillDesc & "<td name=""tdLabelInpVisitationID"" id=""tdLabelInpVisitationID"" style=""font-weight: bold"">Visit No.</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdInputInpVisitationID"" id=""tdInputInpVisitationID"">" & (.fields("VisitationID")) & "</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdLabelInpVisitTypeID"" id=""tdLabelInpVisitTypeID"" style=""font-weight: bold"">Admission No.</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdInputInpVisitTypeID"" id=""tdInputInpVisitTypeID"">" & (GetRecordField("AdmissionID")) & "</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "</tr>"

        strBillDesc = strBillDesc & "<tr>"
        strBillDesc = strBillDesc & "<td name=""tdLabelInpPatientID"" id=""tdLabelInpPatientID"" style=""font-weight: bold"">Patient No.</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdInputInpPatientID"" id=""tdInputInpPatientID"">" & (.fields("PatientID")) & "</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdLabelInpVisitTypeID"" id=""tdLabelInpVisitTypeID"" style=""font-weight: bold"">Patient Name</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdInputInpVisitTypeID"" id=""tdInputInpID"">" & (.fields("PatientName")) & "</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "</tr>"

        strBillDesc = strBillDesc & "<tr>"
        strBillDesc = strBillDesc & "<td name=""tdLabelInpGenderID"" id=""tdLabelInpGenderID"" style=""font-weight: bold"">Gender</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & (.fields("GenderName")) & "</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdLabelInpPatientAge"" id=""tdLabelInpPatientAge"" style=""font-weight: bold"">Age</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        ' strBillDesc = strBillDesc & "<td name=""tdInputInpPatientAge"" id=""tdInputInpPatientAge"">" & CStr(Int(CInt(.fields("PatientAge")))) & "</td>"
        strBillDesc = strBillDesc & "<td name=""tdInputInpPatientAge"" id=""tdInputInpPatientAge"">" & CStr(Int(CInt(.fields("PatientAge")))) & " [" & GetComboNameFld("Visitation", vst, "VisitInfo6") & "]</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "</tr>"

        strBillDesc = strBillDesc & "<tr>"
        strBillDesc = strBillDesc & "<td name=""tdLabelInpInsuredPatientID"" id=""tdLabelInpInsuredPatientID"" style=""font-weight: bold"">Billing Account No.</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdInputInpInsuredPatientID"" id=""tdInputInpInsuredPatientID"">" & (.fields("InsuredPatientID")) & "</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdLabelInpInsuranceSchemeID"" id=""tdLabelInpInsuranceSchemeID"" style=""font-weight: bold"">Billing Info.</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdInputInpInsuranceSchemeID"" id=""tdInputInpInsuranceSchemeID"">" & (.fields("InsuranceSchemeName")) & "</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "</tr>"

        strBillDesc = strBillDesc & "<tr>"
        strBillDesc = strBillDesc & "<td name=""tdLabelInpInsuranceNo"" id=""tdLabelInpInsuranceNo"" style=""font-weight: bold"">InsuranceNo</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdInputInpInsuranceNo"" id=""tdInputInpInsuranceNo"">" & (.fields("InsuranceNo")) & "</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdLabelInpVisitDate"" id=""tdLabelInpVisitDate"" style=""font-weight: bold"">Visit Date</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "<td name=""tdInputInpVisitDate"" id=""tdInputInpVisitDate"">" & (FormatDate(.fields("VisitDate"))) & "</td>"
        strBillDesc = strBillDesc & "<td width=""20""></td>"
        strBillDesc = strBillDesc & "</tr>"

        DisplayAdmitDate vst

        strBillDesc = strBillDesc & "</table>"
        strBillDesc = strBillDesc & "</td>"
        strBillDesc = strBillDesc & "</tr>"

        'Bill Details
        strBillDesc = strBillDesc & "<tr>"
        strBillDesc = strBillDesc & "<td align=""center"">"
        strBillDesc = strBillDesc & "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""2"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 9pt; font-family: Arial"">"

        strBillDesc = strBillDesc & "<tr>"
        ' strBillDesc = strBillDesc & "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1"" height=""20""></td>"
        strBillDesc = strBillDesc & "<td colspan=""8"" align=""center""><hr color=""#999999"" size=""1"" height=""20""></td>"
        strBillDesc = strBillDesc & "</tr>"
        strBillDesc = strBillDesc & "<tr>"
        strBillDesc = strBillDesc & "<td style=""font-weight: bold""><u>NO.</u></td>"
        strBillDesc = strBillDesc & "<td style=""font-weight: bold""><u>SERVICE DESCRIPTION</u></td>"
        strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>QTY&nbsp;&nbsp</u></td>"
        strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>UNIT&nbsp;</u>&emsp;</td>"
        strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>&nbsp;TOTAL</u>&nbsp;</td>"
        If hasDiscount Then
        strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>&nbsp;&nbsp;DISCOUNT</u></td>"
        Else
        strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""></td>"
        End If
        ' strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>&nbsp;&nbsp;CO-PAY</u><br>&emsp;<u>(SPONSOR)</u></td>"
        ' strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>&nbsp;&nbsp;CO-PAY</u><br>&emsp;<u>(PAYING)</u></td>"
        If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
        strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
        strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
        Else
        strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>&nbsp;&nbsp;CO-PAY</u>&emsp;<u>(SPONSOR)</u></td>"
        strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>&nbsp;&nbsp;CO-PAY</u>&emsp;<u>(PAYING)</u></td>"
        End If
        strBillDesc = strBillDesc & "</tr>"

        'Registration
        ''AddRegister pat, GetComboNameFld("Visitation", vst, "WorkingDayID")
        'Consultation
        hdr = (.fields("SpecialistTypeName")) & " [Consultation]"
        strBillDesc = strBillDesc & "<tr>"
        ' strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;Consultation</td>"
        strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;<h6>Consultation</h6></td>"
        strBillDesc = strBillDesc & "</tr>"
        pos = pos + 1
        otCash = .fields("Visitcost3")
        otSponsor = .fields("Visitcost") - otCash
        gTot = gTot + .fields("Visitcost")
        gTotDisc = gTotDisc + .fields("Visitcost2") ''Discount
        gTotCash = gTotCash + otCash ''Discount cash total
        gTotSpon = gTotSpon + otSponsor
        strBillDesc = strBillDesc & "<tr>"
        strBillDesc = strBillDesc & "<td>" & CStr(pos) & "</td>"
        strBillDesc = strBillDesc & "<td>" & hdr & "</td>"
        strBillDesc = strBillDesc & "<td align=""right"">1</td>"
        strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(.fields("Visitcost")), 2, , , -1)) & "</td>"
        strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(.fields("Visitcost")), 2, , , -1)) & "</td>"
        If hasDiscount Then
        strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
        Else
        strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
        End If

        If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
        strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
        strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
        Else
        End If
        strBillDesc = strBillDesc & "</tr>"

        AddAdmission vst

        ''============== DRUGS
        ''==========================================
        AddDrug vst
        AddDrug2 vst

        '@bless - 5 Jun 2022 //Entries on [sub encounter] Credit => VisitationID-C
        If UCase(rTyp) = "R002" Then 'Credit
          AddDrug vst & "-C"
          AddDrug vst & "-D"
          AddDrug2 vst & "-C"
          AddDrug2 vst & "-D"
        End If

        ''============== CONSUMABLES
        ''==========================================
        AddNonDrug vst
        '@bless - 5 Jun 2022 //Entries on [sub encounter] Credit => VisitationID-C
        If UCase(rTyp) = "R002" Then 'Credit
            AddNonDrug vst & "-C"
            AddNonDrug vst & "-D"
        End If

        ''============== INVESTIGATIONS
        ''==========================================
        AddLab vst
        AddLab2 vst

        If UCase(rTyp) = "R002" Then 'Credit
            AddLab vst & "-C"
            AddLab vst & "-D"
            AddLab2 vst & "-C"
            AddLab2 vst & "-D"
        End If

        ''============== SERVICE/PROCEDURE
        ''==========================================
        AddTreat vst
        If UCase(rTyp) = "R002" Then 'Credit
            AddTreat vst & "-C"
            AddTreat vst & "-D"
        End If

        ''============== BLOOD BANK
        ''==========================================
        ' 'Blood Bank notice '20200428 @bless - Include specific item for blood bank/Requested by LabHead'
        ' ' If UCase(jSchd) = UCase(uName) Then
        '     AddBloodBankNotice vst
        ' ' End If

        If UCase(addStat) <> "A002" And UCase(addStat) <> "A004" And UCase(addStat) <> "A008" Then
            strBillDesc = strBillDesc & "<tr>"
            strBillDesc = strBillDesc & "<td colspan=""8""><font size=""8"" color=""red"">PROVISIONAL BILL</font></td>"
            strBillDesc = strBillDesc & "<tr>"
        End If
        strBillDesc = strBillDesc & "<tr>"
        ' strBillDesc = strBillDesc & "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
        strBillDesc = strBillDesc & "<td colspan=""8"" align=""center""><hr color=""#999999"" size=""1""></td>"
        strBillDesc = strBillDesc & "</tr>"

        'Grand Total
        ' gTot = FormatNumber((gTot + 0.04), 1)
        strBillDesc = strBillDesc & "<tr class=""grand-total"">"
        strBillDesc = strBillDesc & "<td></td>"
        strBillDesc = strBillDesc & "<td colspan=""3"" align=""left"">TOTAL BILL</td>"
        ' strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTot), 2, , , -1)) & "</td>"
        strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTot), 2, , , -1)) & "</td>"
        If hasDiscount Then
        strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTotDisc), 2, , , -1)) & "</td>"
        Else
        strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
        End If

        If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
        strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
        strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
        Else
        strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTotSpon), 2, , , -1)) & "</td>"
        strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTotCash), 2, , , -1)) & "</td>"
        totCoPay = gTotCash ''05 Jul 2024
        End If
        strBillDesc = strBillDesc & "</tr>"
        ' strBillAmt = (FormatNumber(CStr(gTot), 2, , , -1))
        strBillAmt = (FormatNumber(CStr(gTot), 2, , , -1))

        'Add Insurance
        ''If (UCase(Trim(wd)) <> "WD01") And (UCase(Trim(wd)) <> "WD11") And (UCase(Trim(wd)) <> "WD05") And (UCase(Trim(wd)) <> "WD19") Then
        'If UCase(insTyp) = "I101" Then 'NHIS
        '  AddNHISInsurance vst
        'End If
        ''End If
        'Add Waivers
        'AddWaiver vst
        AddWaiver2 vst '25 Jun 2019

        If (gWv > 0) Or (gIns > 0) Then
            strBillAmt = FormatNumber((gTot - gWv - gIns) + 0.04, 1)
            strBillAmt = (FormatNumber(strBillAmt, 2, , , -1))
            strBillDesc = strBillDesc & "<tr>"
            strBillDesc = strBillDesc & "<td></td>"
            strBillDesc = strBillDesc & "<td colspan=""3"" align=""left"">FINAL TOTAL BILL</td>"
            ' strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTot - gWv - gIns), 2, , , -1)) & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & strBillAmt & "</td>"
            If hasDiscount Then
            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTotDisc), 2, , , -1)) & "</td>"
            Else
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            End If

            If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            Else
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTotSpon), 2, , , -1)) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTotCash), 2, , , -1)) & "</td>"
            End If
            strBillDesc = strBillDesc & "</tr>"
        End If

        strBillDesc = strBillDesc & "<tr>"
        ' strBillDesc = strBillDesc & "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
        strBillDesc = strBillDesc & "<td colspan=""8"" align=""center""><hr color=""#999999"" size=""1""></td>"
        strBillDesc = strBillDesc & "</tr>"
        amtInvoice = 0

        'Payments
        recNo = CleanReceiptNos(recNo)
        AddPayments pat, recNo

        'UsedPayments
        AddUsedPayments pat, recNo

        'Grand Total
        otTot = gTot - gPaid + gUsed - gWv - gIns - insWv
        conn.execute qryPro.FltQry("update Admission Set MainInfo3='', MainValue1=0 where VisitationID='" & vst & "'")
        conn.execute qryPro.FltQry("update Admission Set MainInfo3='', MainValue1=0, MainValue2=0 where VisitationID LIKE '" & vst & "-%'") ''@bless - 05 Jul 2024

        'Refund Required
        If (gTot - gWv - gIns - insWv) < 0 Then 'Insurance more than Total cost
            ' strBillOust = (FormatNumber(CStr((gUsed - gPaid) + 0.04), 1, , , -1))
            strBillOust = (FormatNumber(CStr((gUsed - gPaid)), 2, , , -1))
            strBillDesc = strBillDesc & "<tr class=""grand-total"">"
            strBillDesc = strBillDesc & "<td></td>"
            strBillDesc = strBillDesc & "<td colspan=""3"" align=""left""><b>OUTSTANDING BILL</b></td>"
            ' strBillDesc = strBillDesc & "<td align=""right""><b>" & (FormatNumber(CStr(gUsed - gPaid), 2, , , -1)) & "</b></td>"
            strBillDesc = strBillDesc & "<td align=""right""><b>" & (FormatNumber(strBillOust, 2, , , -1)) & "</b></td>"
            If hasDiscount Then
            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTotDisc), 2, , , -1)) & "</td>"
            Else
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            End If

            If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            Else
            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTotSpon), 2, , , -1)) & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTotCash), 2, , , -1)) & "</td>"
            End If
            strBillDesc = strBillDesc & "</tr>"
            strBillOust = (FormatNumber(strBillOust, 2, , , -1))
        Else
            ' otTot = (FormatNumber(CStr(otTot + 0.04), 1, , , -1))
            ''deduct disc
            If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
                strBillOust = otTot - gTotDisc
            Else
                strBillOust = otTot
            End If
            strBillDesc = strBillDesc & "<tr class=""grand-total"">"
            strBillDesc = strBillDesc & "<td></td>"
            strBillDesc = strBillDesc & "<td colspan=""3"" align=""left""><b>OUTSTANDING BILL</b></td>"
            strBillDesc = strBillDesc & "<td align=""right""><b>" & (FormatNumber(CStr(otTot), 2, , , -1)) & "</b></td>"
            ' strBillDesc = strBillDesc & "<td align=""right""><b>" & (FormatNumber(CStr(strBillOust), 2, , , -1)) & "</b></td>"
            If hasDiscount Then
            ' strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTotDisc), 2, , , -1)) & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            Else
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            End If

            If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            Else
            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTotSpon), 2, , , -1)) & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gTotCash), 2, , , -1)) & "</td>"
            End If
            strBillDesc = strBillDesc & "</tr>"
            ' ''Patients paying/cash component of bill
            ' If (UCase(rTyp) = "R001" Or UCase(insGrp) = UCase("NHIS") Or UCase(insGrp) = UCase("CASH")) Or (UCase(insTyp) = UCase("I100")) Then 'Compile Oustanding bill for paying Patients
            ' Else
            '     strBillOust = gTotCash
            ' End If

            ''bill summary
            strBillDesc = strBillDesc & "<tr class=""grand-total"">"
            If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
                strBillDesc = strBillDesc & "<td></td>"
                strBillDesc = strBillDesc & "<td colspan=""2"" align=""right"">"
                strBillDesc = strBillDesc & "<b>Total Patient Bill</b>&nbsp:&nbsp<br>"
                If hasDiscount Then
                    strBillDesc = strBillDesc & "<b>Total Gratis/Discount</b>&nbsp:&nbsp<br>"
                End If
                strBillDesc = strBillDesc & "<b>Total Patient Balance/Payments</b>&nbsp:&nbsp<br>"
                strBillDesc = strBillDesc & "<b style=""font-size: medium;"">Outstanding Amount To Pay</b>&nbsp:&nbsp<br>"
                strBillDesc = strBillDesc & "</td>"
                strBillDesc = strBillDesc & "<td colspan=""1"" align=""right"">"
                strBillDesc = strBillDesc & "<b>" & strBillAmt & "</b><br>"
                If hasDiscount Then
                    strBillDesc = strBillDesc & "<b>(" & (FormatNumber(CStr(gTotDisc), 2, , , -1)) & ")</b><br>"
                End If
                strBillDesc = strBillDesc & "<b>(" & (FormatNumber(CStr(gPaid - gUsed), 2, , , -1)) & ")</b><br>" '@Frank 27/03/2024 subtracted used receipt to show the actual amount the client payed on this bill
                ' strBillDesc = strBillDesc & "<b style=""font-size: medium;"">" & (FormatNumber(CStr(strBillOust), 2, , , -1)) & "</b><br>"

                If strBillOust < 0 Then
                    strBillDesc = strBillDesc & "<b style=""font-size: medium;"">" & "(NIL)" & "</b><br>"
                    conn.execute qryPro.FltQry("update Admission Set MainInfo3='" & "0" & "', MainValue1=0 where AdmissionID='" & adm & "'")
                Else
                    strBillDesc = strBillDesc & "<b style=""font-size: medium;"">" & (FormatNumber(CStr(strBillOust), 2, , , -1)) & "</b><br>"
                    'conn.execute qryPro.FltQry("update Admission Set MainInfo3='" & Replace(FormatNumber(strBillOust, 2), ",", "") & ", MainValue1=" & Replace(FormatNumber(strBillOust, 2), ",", "") & "' where AdmissionID='" & adm & "'")
                    conn.execute qryPro.FltQry("update Admission Set MainInfo3='" & Replace(FormatNumber(strBillOust, 2), ",", "") & "', MainValue1=" & Replace(FormatNumber(strBillOust, 2), ",", "") & " where AdmissionID='" & adm & "'") ''@frank 13 jun 2024
                    amtInvoice = strBillOust
                End If
                strBillDesc = strBillDesc & "</td>"

            Else
                strBillDesc = strBillDesc & "<td></td>"
                strBillDesc = strBillDesc & "<td colspan=""4"" align=""right"">"
                strBillDesc = strBillDesc & "<b>Total Patient Bill</b>&nbsp:&nbsp<br>"
                If hasDiscount Then
                    strBillDesc = strBillDesc & "<b>Total Gratis/Discount</b>&nbsp:&nbsp<br>"
                End If
                strBillDesc = strBillDesc & "<b>Total Patient Balance/Payments</b>&nbsp:&nbsp<br>"
                strBillDesc = strBillDesc & "<b>Total Sponsor Coverage</b>&nbsp:&nbsp<br>"
                strBillDesc = strBillDesc & "<b style=""font-size: medium;"">Outstanding Amount To Pay</b>&nbsp:&nbsp<br>"
                strBillDesc = strBillDesc & "</td>"
                strBillDesc = strBillDesc & "<td colspan=""1"" align=""right"">"
                strBillDesc = strBillDesc & "<b>" & strBillAmt & "</b><br>"
                If hasDiscount Then
                    strBillDesc = strBillDesc & "<b>(" & (FormatNumber(CStr(gTotDisc), 2, , , -1)) & ")</b><br>"
                End If
                strBillDesc = strBillDesc & "<b>(" & (FormatNumber(CStr(gPaid - gUsed), 2, , , -1)) & ")</b><br>" '@Frank 27/03/2024 subtracted used receipt to show the actual amount the client payed on this bill
                strBillDesc = strBillDesc & "<b>(" & (FormatNumber(CStr(gTotSpon), 2, , , -1)) & ")</b><br>"
                ' strBillDesc = strBillDesc & "<b style=""font-size: medium;"">" & (FormatNumber(CStr(gTotCash), 2, , , -1)) & "</b><br>"
                If gTotCash < 0 Then
                    strBillDesc = strBillDesc & "<b style=""font-size: medium;"">" & "(NIL)" & "</b><br>"
                    conn.execute qryPro.FltQry("update Admission Set MainInfo3='" & "0" & "', MainValue1=0 where AdmissionID='" & adm & "'")
                Else
                    gTotCash2 = (FormatNumber(CStr(gTotCash - gPaid), 2, , , -1)) ''@bless - 09 Jan 2024 ''deduct receipts from Outstanding
                    If (gTotCash - gPaid) < 0 Then
                        strBillDesc = strBillDesc & "<b style=""font-size: medium;"">" & "(NIL)" & "</b><br>"
                    Else
                        strBillDesc = strBillDesc & "<b style=""font-size: medium;"">" & (FormatNumber(CStr(gTotCash - gPaid), 2, , , -1)) & "</b><br>"
                    End If
                    ' conn.execute qryPro.FltQry("update Admission Set MainInfo3='" & Replace(FormatNumber(gTotCash2, 2), ",", "") & "', MainValue1=" & Replace(FormatNumber(gTotCash, 2), ",", "") & " where AdmissionID='" & adm & "'")
                    conn.execute qryPro.FltQry("update Admission Set MainInfo3='" & Replace(FormatNumber(gTotCash2, 2), ",", "") & "', MainValue1=" & Replace(FormatNumber(gTotCash, 2), ",", "") & " , MainValue2=" & Replace(FormatNumber(totCoPay, 2), ",", "") & " where AdmissionID='" & adm & "'")
                    amtInvoice = gTotCash
                End If
                amtAllowCopay = gTotSpon
                strBillDesc = strBillDesc & "</td>"
            End If
            strBillDesc = strBillDesc & "</tr>"

            strBillOust = (FormatNumber(CStr(strBillOust), 2, , , -1))
        End If

        If IsCancelledAdmission Then ''@bless - 26 Apr 2024 // Admissions are cancelled, no bill
            ''conn.execute qryPro.FltQry("update Admission Set MainInfo3='' where VisitationID='" & vst & "'")
        End If

        strBillDesc = strBillDesc & "<tr>"
        ' strBillDesc = strBillDesc & "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
        strBillDesc = strBillDesc & "<td colspan=""8"" align=""center""><hr color=""#999999"" size=""1""></td>"
        strBillDesc = strBillDesc & "</tr>"

        SetBillCompleted vst
        SetBillReviewed vst
        strBillCompltBy = Trim(cmptdSig) & "||" & Trim(cmptdBy) & "||" & Trim(cmptdDt)
        strBillRevwBy = Trim(revSig) & "||" & Trim(revBy) & "||" & Trim(revDt)


        strBillDesc = strBillDesc & "<tr>"
        ' strBillDesc = strBillDesc & "<td colspan=""5"" align=""center"">"
        strBillDesc = strBillDesc & "<td colspan=""8"" align=""center"">"
        strBillDesc = strBillDesc & "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""2"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 8pt; font-family: Arial"">"

        strBillDesc = strBillDesc & "<tr>"
        strBillDesc = strBillDesc & "<td width=""25%"" align=""left""><u><b>REVIEWED&nbsp;BY</b></u></td>"
        strBillDesc = strBillDesc & "<td width=""25%"" align=""left""><u><b>COMPLETED&nbsp;BY</b></u></td>"
        strBillDesc = strBillDesc & "<td width=""25%"" align=""left""><u><b>PRINTED&nbsp;BY</b></u></td>"
        strBillDesc = strBillDesc & "<td width=""25%"" align=""left""><u><b>NURSING&nbsp;OFFICER&nbsp;IN&nbsp;CHARGE</b></u></td>"
        strBillDesc = strBillDesc & "</tr>"

        strBillDesc = strBillDesc & "<tr height=""40"">"
        strBillDesc = strBillDesc & "<td align=""left""><u>" & cmptdBy & "</u></td>"
        strBillDesc = strBillDesc & "<td align=""left""><u>" & revBy & "</u></td>"
        strBillDesc = strBillDesc & "<td align=""left""><u>" & UCase(GetComboName("Staff", GetComboNameFld("SystemUser", uName, "StaffID"))) & "</u></td>"
        strBillDesc = strBillDesc & "<td align=""left"" valign=""bottom""><hr color=""#999999"" size=""1""></td>"
        strBillDesc = strBillDesc & "</tr>"

        strBillDesc = strBillDesc & "<tr height=""40"">"
        strBillDesc = strBillDesc & "<td valign=""bottom""><b>SIGN&nbsp;:&nbsp;</b><u>" & cmptdSig & "</u></td>"
        strBillDesc = strBillDesc & "<td valign=""bottom""><b>SIGN&nbsp;:&nbsp;</b><u>" & revSig & "</u></td>"
        strBillDesc = strBillDesc & "<td valign=""bottom""><b>SIGN&nbsp;:&nbsp;</b><hr color=""#999999"" size=""1""></td>"
        strBillDesc = strBillDesc & "<td valign=""bottom""><b>SIGN&nbsp;:&nbsp;</b><hr color=""#999999"" size=""1""></td>"
        strBillDesc = strBillDesc & "</tr>"

        strBillDesc = strBillDesc & "<tr height=""40"">"
        strBillDesc = strBillDesc & "<td valign=""bottom""><b>DATE&nbsp;:&nbsp;</b><u>" & cmptdDt & "</u></td>"
        strBillDesc = strBillDesc & "<td valign=""bottom""><b>DATE&nbsp;:&nbsp;</b><u>" & revDt & "</u></td>"
        strBillDesc = strBillDesc & "<td valign=""bottom""><b>DATE&nbsp;:&nbsp;</b><u>" & FormatDateDetail(Now()) & "</u></td>"
        strBillDesc = strBillDesc & "<td valign=""bottom""><b>DATE&nbsp;:&nbsp;</b><hr color=""#999999"" size=""1""></td>"
        strBillDesc = strBillDesc & "</tr>"

        strBillDesc = strBillDesc & "</table>"
        ''@Bless - 2 Dec 2023 //Issue Receipt by cashier
        ' If otTot > 0 Then
        If amtInvoice > 0 Then
            ' url = "wpgReceipt.asp?PageMode=AddNew&PullUpData=PatientID||" & GetRecordField("PatientID") & "&BillAmount=" & Replace(otTot, ",", "")
            url = "wpgReceipt.asp?PageMode=AddNew&PullUpData=PatientID||" & GetRecordField("PatientID") & "&BillAmount=" & Replace(amtInvoice, ",", "")
            url = url & "&TableName=Admission&VisitationID=" & vst & "&AdmissionID=" & adm ''@bless - 5 Jul 2024
            strBillDesc = strBillDesc & "<a class='btn btn-info no-print' href='" & url & "' target='_blank' title='Issue Receipt'>Issue Receipt</a>"

            If UCase(Trim(Request("DisplayType"))) = UCase("IssueReceipt") Then
                response.redirect url
            End If
        End If

        If UCase("S17") = UCase(jSchd) Then
            ' url = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=Admission2&PositionForTableName=Admission&AdmissionID=" & adm & "&VisitationID=" & Replace(vst, "?", "") & "&ProcessCoPay=YES&ProcessCoPayAmount=" & amtInvoice
            ' strBillDesc = strBillDesc & "&emsp;<a class='btn btn-info no-print' href='" & url & "' target='_self' title='Test Drug Co-Payment'>Test Drug Co-Payment</a>"

            url = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationClaimB&PositionForTableName=Visitation&VisitationID=" & Replace(vst, "?", "") & ""
            strBillDesc = strBillDesc & "&emsp;<a class='btn btn-info no-print' href='" & url & "' target='_self' title='Test Drug Co-Payment'>Vet Claim</a>"
        End If
        strBillDesc = strBillDesc & "</td>"
        strBillDesc = strBillDesc & "</tr>"

        strBillDesc = strBillDesc & Glob_Footer()

        strBillDesc = strBillDesc & "</table>"
        strBillDesc = strBillDesc & "</td>"
        strBillDesc = strBillDesc & "</tr>"
    End If

    .Close
    ' ' ' If UCase(rTyp) = "R001" Then 'Compile Oustanding bill for paying Patients
    ' ' ' If UCase(rTyp) = "R001"  Or UCase(insGrp)=UCase("NHIS") Then 'Compile Oustanding bill for paying Patients
    ' ' If (UCase(rTyp) = "R001" Or UCase(insGrp) = UCase("NHIS") Or UCase(insGrp) = UCase("CASH")) Or (UCase(insTyp) = UCase("I100")) Then 'Compile Oustanding bill for paying Patients
    ' ' If otTot > 0 Then
    ' If strBillOust > 0 Then
    '     ' conn.Execute qryPro.FltQry("update Admission Set MainInfo3='" & otTot & "' where AdmissionID='" & adm & "'")
    '     conn.Execute qryPro.FltQry("update Admission Set MainInfo3='" & strBillOust & "' where AdmissionID='" & adm & "'")
    '     response.write "Here..."
    ' Else
    '     conn.Execute qryPro.FltQry("update Admission Set MainInfo3='" & "0" & "' where AdmissionID='" & adm & "'")
    '     response.write "Here 2..."
    ' End If
    ' End If
    response.write strBillDesc
    response.flush

    logPatientBill
End With

    addJS ''@bless - 22 Dec 2023
If Not vldPrint Then
    ''addJS ''@bless - Disabled //Ometse Wilhemina
End If

Function VisitHasDicount(vst)
    Dim rst, sql, ot, amt
    Set rst = CreateObject("ADODB.Recordset")
    ot = False

    sql = "select sum(DiscAmt) as DiscAmt from treatcharges where VisitationID='" & vst & "' "
    sql = sql & " Or VisitationID IN ('" & vst & "-C','" & vst & "-D') "
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        If Not IsNull(rst.fields("DiscAmt")) Then
            amt = rst.fields("DiscAmt")
            If IsNumeric(amt) And amt > 0 Then
                ot = True
            End If
        End If
    End If
    rst.Close

    Set rst = Nothing
    VisitHasDicount = ot
End Function

Sub UpdateCopayTables(tbl, id, vst)
    Select Case UCase(tbl) ''Copay
        Case UCase("DrugSale") ''Drug
            ' conn.execute qryPro.FltQry("UPDATE DrugSaleItems SET MainInfo1=Round(FinalAmt, 4), MainItemInfo1='0' Where VisitationID='" & vst & "' And DrugID='" & id & "' And Len(VisitationID) >= 5;")
            conn.execute qryPro.FltQry("UPDATE DrugSaleItems SET MainInfo1=Round(FinalAmt, 4) - MainItemValue2, MainItemInfo1='0' Where VisitationID='" & vst & "' And DrugID='" & id & "' And Len(VisitationID) >= 5;")
            ' conn.execute qryPro.FltQry("UPDATE DrugSaleItems2 SET MainInfo1=Round(DispenseAmt2, 4), DispenseInfo1='0' Where VisitationID='" & vst & "' And DrugID='" & id & "' And Len(VisitationID) >= 5;")
            conn.execute qryPro.FltQry("UPDATE DrugSaleItems2 SET MainInfo1=Round(DispenseAmt2, 4) - MainValue2, DispenseInfo1='0' Where VisitationID='" & vst & "' And DrugID='" & id & "' And Len(VisitationID) >= 5;")
            Glob_LogAuditTrail "ConsultReview", vst, "URT003", uName
            SetPageMessages GetComboName("Drug", id) & " is copaid by client "

                ' ''@bless - 25 Dec 2023 //Account for Returns
                ' Dim rst, sql
                ' Set rst = CreateObject("ADODB.Recordset")
                ' sql = "select DrugSaleID from DrugSaleItems Where VisitationID='" & vst & "' And DrugID='" & id & "' "
                ' sql = sql & " "
                ' With rst
                '     rst.Open qryPro.FltQry(sql), conn, 3, 4
                '     If rst.RecordCount > 0 Then
                '         rst.MoveFirst
                '         Do While Not rst.EOF
                '             drgSl = rst.fields("DrugSaleID")
                '             amt = rst.Fields("FinalAmt")
                '             rAmt = GetReturnAmt3(vst, id, drgSl)
                '             If IsNumeric(amt) And IsNumeric(rAmt) And (IsNumeric(amt) >= IsNumeric(rAmt))  Then
                '                 amt = amt - rAmt
                '             End If
                '             rst.Fields("MainInfo1") = amt
                '             ' rst.Fields("MainItemInfo1") = "0"
                '             rst.UpdateBatch
                '             rst.MoveNext
                '         Loop
                '     End If
                '     rst.Close
                ' End With

                ' sql = "select DrugSaleID from DrugSaleItems2 Where VisitationID='" & vst & "' And DrugID='" & id & "' "
                ' sql = sql & " "
                ' With rst
                '     rst.Open qryPro.FltQry(sql), conn, 3, 4
                '     If rst.RecordCount > 0 Then
                '         rst.MoveFirst
                '         Do While Not rst.EOF
                '             drgSl = rst.fields("DrugSaleID")
                '             amt = rst.Fields("DispenseAmt2")
                '             rAmt = GetReturnAmt4(vst, id, drgSl)
                '             If IsNumeric(amt) And IsNumeric(rAmt) And (IsNumeric(amt) >= IsNumeric(rAmt))  Then
                '                 amt = amt - rAmt
                '             End If
                '             rst.Fields("MainInfo1") = amt
                '             ' rst.Fields("DispenseInfo1") = "0"
                '             rst.UpdateBatch
                '             rst.MoveNext
                '         Loop
                '     End If
                '     rst.Close
                ' End With
                ' set rst = Nothing


        Case UCase("DRUGSALE-UNDO") ''Drug uNDONE
            conn.execute qryPro.FltQry("UPDATE DrugSaleItems SET MainItemInfo1=Round(FinalAmt, 4) - MainItemValue2, MainInfo1='0' Where VisitationID='" & vst & "' And DrugID='" & id & "' And Len(VisitationID) >= 5;")
            conn.execute qryPro.FltQry("UPDATE DrugSaleItems2 SET DispenseInfo1=Round(DispenseAmt2, 4) - MainValue2, MainInfo1='0' Where VisitationID='" & vst & "' And DrugID='" & id & "' And Len(VisitationID) >= 5;")
            Glob_LogAuditTrail "ConsultReview", vst, "URT003", uName
            SetPageMessages GetComboName("Drug", id) & " is reversed on copay by client "

            ' ''@bless - 25 Dec 2023 //Account for Returns
            ' Dim rst, sql
            ' Set rst = CreateObject("ADODB.Recordset")
            ' sql = "select DrugSaleID from DrugSaleItems Where VisitationID='" & vst & "' And DrugID='" & id & "' "
            ' sql = sql & " "
            ' With rst
            '     rst.Open qryPro.FltQry(sql), conn, 3, 4
            '     If rst.RecordCount > 0 Then
            '         rst.MoveFirst
            '         Do While Not rst.EOF
            '             drgSl = rst.fields("DrugSaleID")
            '             amt = rst.Fields("FinalAmt")
            '             rAmt = GetReturnAmt3(vst, id, drgSl)
            '             If IsNumeric(amt) And IsNumeric(rAmt) And (IsNumeric(amt) >= IsNumeric(rAmt))  Then
            '                 amt = amt - rAmt
            '             End If
            '             rst.Fields("MainInfo1") = "0"
            '             rst.Fields("MainItemInfo1") = amt
            '             rst.UpdateBatch ''Update
            '             rst.MoveNext
            '         Loop
            '     End If
            '     rst.Close
            ' End With

            ' sql = "select DrugSaleID from DrugSaleItems2 Where VisitationID='" & vst & "' And DrugID='" & id & "' "
            ' sql = sql & " "
            ' With rst
            '     rst.Open qryPro.FltQry(sql), conn, 3, 4
            '     If rst.RecordCount > 0 Then
            '         rst.MoveFirst
            '         Do While Not rst.EOF
            '             drgSl = rst.fields("DrugSaleID")
            '             amt = rst.Fields("DispenseAmt2")
            '             rAmt = GetReturnAmt4(vst, id, drgSl)
            '             If IsNumeric(amt) And IsNumeric(rAmt) And (IsNumeric(amt) >= IsNumeric(rAmt))  Then
            '                 amt = amt - rAmt
            '             End If
            '             rst.Fields("MainInfo1") = "0"
            '             rst.Fields("DispenseInfo1") = amt
            '             rst.UpdateBatch ''Update
            '             rst.MoveNext
            '         Loop
            '     End If
            '     rst.Close
            ' End With
            ' set rst = Nothing


        Case UCase("LabRequest") ''Lab

        Case UCase("LabRequest") ''Lab

        Case UCase("StockIssue") ''General Items

    End Select

End Sub

'logPatientBill() >> Log patientbill after bill is ready for printing
Sub logPatientBill()
    Dim rstV, sqlV
    Set rstV = CreateObject("ADODB.Recordset")
    Dim tblId, tblNm, dtNow
    dtNow = Now()

    Set rst1 = CreateObject("ADODB.Recordset")
    ' sql = "select * from AdmissionPro where VisitationID='" & vst & "' "
    ' sql = sql & " and TransProcessStat2ID='T013'" 'Bill Reviewed
    ' sql = sql & " and (JobScheduleID NOT LIKE 'W%' and TransProcessDate1 > '21 May 2019 00:00:00') "
    ' sql = sql & " order by AdmissionProID desc "

    ' rst1.Open qryPro.FltQry(sql), conn, 3, 4

    ' If rst1.RecordCount > 0 Then
    '     rst1.MoveFirst

        tblId = vst & "||" & strBillAmt & "||" & strBillOust
        tblNm = uName & "**" & jSchd & "||" & GetComboName("Visitation", vst)

        sqlV = "select * from TestVar96 where TestVar96ID='" & tblId & "' "
        sqlV = sqlV & "  "
        With rstV
            rstV.open qryPro.FltQry(sqlV), conn, 3, 4
            If rstV.RecordCount > 0 Then
            Else
                rstV.AddNew
                rstV.fields("TestVar96ID") = Left(Trim(tblId), 50)
                rstV.fields("TestVar96Name") = Left(Trim(tblNm), 255)
                rstV.fields("VarPos") = FormatDateDetail(dtNow)
                rstV.fields("Description") = "<table><tr><td>" & strBillDesc & "</td></tr></table>" '' strBillCompltBy
                rstV.fields("KeyPrefix") = Left(strBillRevwBy & "**" & strBillCompltBy, 255)
                rstV.UpdateBatch
            End If
            rstV.Close
        End With
    ' End If
    ' rst1.Close
    Set rst1 = Nothing
    Set rstV = Nothing
End Sub

Function GetUndertakingDate(vst)

    Dim rst, sql, ot
    Set rst = CreateObject("ADODB.Recordset")
    ot = Now()

    sql = "select * from patientwound where VisitationID='" & vst & "' "
    sql = sql & " and WoundTypeID='" & "T006" & "' "

    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        rst.movefirst

        If Not IsNull(rst.fields("EndDate")) Then
            ot = rst.fields("EndDate")
        End If
    End If
    rst.Close

    Set rst = Nothing
    GetUndertakingDate = ot
End Function

'AddPayments
Sub AddPayments(pat, recNo)

    Dim rst, sql, ot, cnt, hdr, dsc, pd, cn, bal, dt, sDt, eDt, vst, dDt, cn2

    Dim arr, ul, num, whcls, r, rCnt, sqlOk, sql2, regRec

    Set rst = CreateObject("ADODB.Recordset")

    vst = GetRecordField("VisitationID")
    'dDt = GetRecordField("DischargeDate")
    'dt = GetComboNameFld("Visitation", vst, "VisitDate")
    'sDt = FormatDate(dt) & " 0:00:00"
    'eDt = Now()
    'If Not IsNull(dDt) Then
    '  If IsDate(dDt) Then
    '    If CDate(dDt) > CDate(sDt) Then
    '      eDt = FormatDate(dDt) & " 23:59:59"
    '    End If
    '  End If
    'End If
    'Receipt No
    sqlOk = False

    If UCase(pat) = "P1" Then
        whcls = ""
        arr = Split(recNo, ",")
        ul = UBound(arr)
        rCnt = 0

        For num = 0 To ul
            r = Trim(arr(num))

            If Len(r) > 0 Then
                sqlOk = True
                rCnt = rCnt + 1

                If rCnt = 1 Then
                    whcls = whcls & " where "
                Else
                    whcls = whcls & " or "
                End If

                whcls = whcls & " (PatientID='" & pat & "' and ReceiptID='" & r & "') "
            End If

        Next

        If Len(Trim(whcls)) > 0 Then
            sqlOk = True
            sql = "select * from Receipt "
            sql = sql & " " & whcls
            sql = sql & " order by receiptDate"
        End If

    Else
        '  sqlOk = True
        '  whcls = ""
        '  arr = Split(recNo, ",")
        '  ul = UBound(arr)
        '  For num = 0 To ul
        '    r = Trim(arr(num))
        '    If Len(r) > 0 Then
        '      whcls = whcls & " or (PatientID='" & pat & "' and ReceiptID='" & r & "') "
        '    End If
        '  Next
        '  sql = "select * from Receipt where (Patientid='" & pat & "' and receiptdate between '" & sDt & "' and '" & eDt & "')"
        '  sql = sql & " " & whcls
        '  sql = sql & " order by receiptDate"

        sqlOk = True
        sql2 = "select distinct receiptid from patientreceipt2 where visitationid='" & vst & "' and PatientID='" & pat & "'"
        whcls = ""
        arr = Split(recNo, ",")
        ul = UBound(arr)

        For num = 0 To ul
            r = Trim(arr(num))

            If Len(r) > 0 Then
                whcls = whcls & " or (PatientID='" & pat & "' and ReceiptID='" & r & "') "
            End If

        Next

        'Add Registration Receipt
        If withReg Then
            regRec = GetComboNameFld("Patient", pat, "PatientInfo1")

            If Not IsNull(regRec) Then
                regRec = Trim(regRec)

                If Len(regRec) > 0 Then
                    whcls = whcls & " or (PatientID='" & pat & "' and ReceiptID='" & regRec & "') "
                End If
            End If
        End If

        sql = "select * from Receipt where (receiptID in (" & sql2 & "))"
        sql = sql & " " & whcls
        sql = sql & " order by receiptDate"
    End If

    cnt = 0

    If sqlOk Then

        With rst
            .open qryPro.FltQry(sql), conn, 3, 4

            If .RecordCount > 0 Then
                .movefirst
                'Receipt
                hdr = "Payments"
                strBillDesc = strBillDesc & "<tr>"
                strBillDesc = strBillDesc & "<td style=""font-weight: bold""><u>NO.</u></td>"
                strBillDesc = strBillDesc & "<td style=""font-weight: bold""><u>PAYMENT DESCRIPTION</u></td>"
                strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>PAID&nbsp;&nbsp</u></td>"
                strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>REFUND&nbsp;&nbsp</u></td>"
                strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>&nbsp;&nbspBAL.&nbsp;AMT</u></td>"
                strBillDesc = strBillDesc & "</tr>"

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
                    strBillDesc = strBillDesc & "<tr>"
                    strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
                    strBillDesc = strBillDesc & "<td>[REC#&nbsp;:&nbsp;" & UCase(.fields("ReceiptID")) & "&nbsp;]&nbsp;" & dsc & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(pd), 2, , , -1)) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & cn2 & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(bal), 2, , , -1)) & "</td>"
                    strBillDesc = strBillDesc & "</tr>"
                    .movenext
                Loop

                strBillDesc = strBillDesc & "<tr>"
                strBillDesc = strBillDesc & "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
                strBillDesc = strBillDesc & "</tr>"

                'Grand Total
                strBillDesc = strBillDesc & "<tr class=""sub-total"">"
                strBillDesc = strBillDesc & "<td></td>"
                strBillDesc = strBillDesc & "<td colspan=""3"" align=""left"">TOTAL RECEIPT PAYMENT</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gPaid), 2, , , -1)) & "</td>"
                strBillDesc = strBillDesc & "</tr>"

                strBillDesc = strBillDesc & "<tr>"
                strBillDesc = strBillDesc & "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
                strBillDesc = strBillDesc & "</tr>"
            End If

            .Close
        End With

    End If 'Sql

    Set rst = Nothing
End Sub

'AddUsedPayments
Sub AddUsedPayments(pat, recNo)

    Dim rst, rst2, sql, ot, cnt, hdr, dsc, pd, cn, bal, dt, sDt, eDt, vst, dDt

    Dim cnt2, cn2, rec, usd, uCnt, sql2, regRec

    Dim arr, ul, num, whcls, r, rCnt, sqlOk

    Set rst = CreateObject("ADODB.Recordset")
    Set rst2 = CreateObject("ADODB.Recordset")

    vst = GetRecordField("VisitationID")
    'dDt = GetRecordField("DischargeDate")
    'dt = GetComboNameFld("Visitation", vst, "VisitDate")
    'sDt = FormatDate(dt) & " 0:00:00"
    'eDt = Now()
    uCnt = 0
    'If Not IsNull(dDt) Then
    '  If IsDate(dDt) Then
    '    If CDate(dDt) > CDate(sDt) Then
    '      eDt = FormatDate(dDt) & " 23:59:59"
    '    End If
    '  End If
    'End If

    'Receipt No
    sqlOk = False

    If UCase(pat) = "P1" Then
        whcls = ""
        arr = Split(recNo, ",")
        ul = UBound(arr)
        rCnt = 0

        For num = 0 To ul
            r = Trim(arr(num))

            If Len(r) > 0 Then
                sqlOk = True
                rCnt = rCnt + 1

                If rCnt = 1 Then
                    whcls = whcls & " where "
                Else
                    whcls = whcls & " or "
                End If

                whcls = whcls & " (PatientID='" & pat & "' and ReceiptID='" & r & "') "
            End If

        Next

        If Len(Trim(whcls)) > 0 Then
            sqlOk = True
            sql = "select * from Receipt "
            sql = sql & " " & whcls
            sql = sql & " order by receiptDate"
        End If

    Else
        '  sqlOk = True
        '  whcls = ""
        '  arr = Split(recNo, ",")
        '  ul = UBound(arr)
        '  For num = 0 To ul
        '    r = Trim(arr(num))
        '    If Len(r) > 0 Then
        '      whcls = whcls & " or (PatientID='" & pat & "' and ReceiptID='" & r & "') "
        '    End If
        '  Next
        '  sql = "select * from Receipt where (Patientid='" & pat & "' and receiptdate between '" & sDt & "' and '" & eDt & "')"
        '  sql = sql & " " & whcls
        '  sql = sql & " order by receiptDate"

        sqlOk = True
        whcls = ""
        arr = Split(recNo, ",")
        ul = UBound(arr)

        For num = 0 To ul
            r = Trim(arr(num))

            If Len(r) > 0 Then
                whcls = whcls & " or (PatientID='" & pat & "' and ReceiptID='" & r & "') "
            End If

        Next

        '  'Add Registration Receipt
        '  If withReg Then
        '    regRec = GetComboNameFld("Patient", pat, "PatientInfo1")
        '    If Not IsNull(regRec) Then
        '      regRec = Trim(regRec)
        '      If Len(regRec) > 0 Then
        '        whcls = whcls & " or (PatientID='" & pat & "' and ReceiptID='" & regRec & "') "
        '      End If
        '    End If
        '  End If
        ' ''Exclude Sub Encounters Payments as UsedPayments @bless //19 June 2022
        ' sql2 = "select distinct receiptid from patientreceipt2 where visitationid='" & vst & "' and PatientID='" & pat & "'"
        sql2 = "select distinct receiptid from patientreceipt2 where visitationid='" & vst & "' and PatientID='" & pat & "'"
        ' sql2 = sql2 & " And visitationid<>'" & vst & "-C'  "

        sql = "select * from Receipt where (receiptID in (" & sql2 & "))"
        sql = sql & " " & whcls
        sql = sql & " order by receiptDate"
    End If 'P1

    If sqlOk Then
        cnt = 0

        With rst
            .open qryPro.FltQry(sql), conn, 3, 4

            If .RecordCount > 0 Then
                .movefirst

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
                    ' sql = "select * from PatientReceipt2 where Receiptid='" & rec & "' and VisitationID<>'" & vst & "' and VisitationID<>'NONE' order by receiptDate2"
                    ''Exclude Sub Encounters Payments as UsedPayments @bless //19 June 2022 //Transaction Included in bills
                    sql = "select * from PatientReceipt2 where Receiptid='" & rec & "' and (VisitationID<>'" & vst & "' And VisitationID<>'" & vst & "-C') "
                    sql = sql & " and VisitationID<>'NONE' order by receiptDate2"
                    rst2.open qryPro.FltQry(sql), conn, 3, 4

                    If rst2.RecordCount > 0 Then
                        rst2.movefirst

                        Do While Not rst2.EOF
                            cnt2 = cnt2 + 1
                            usd = usd + rst2.fields("PaidAmount")

                            If cnt2 > 1 Then
                                dsc = dsc & "; "
                            End If

                            dsc = dsc & "[V#&nbsp;:&nbsp;" & rst2.fields("VisitationID") & "]&nbsp;" & GetComboName("PaymentType", rst2.fields("PaymentTypeID"))
                            rst2.movenext
                        Loop

                    End If ''R1003388

                    rst2.Close
                    gUsed = gUsed + usd

                    If cnt2 > 0 Then
                        uCnt = uCnt + 1

                        If uCnt = 1 Then
                            strBillDesc = strBillDesc & "<tr>"
                            strBillDesc = strBillDesc & "<td style=""font-weight: bold""><u>NO.</u></td>"
                            strBillDesc = strBillDesc & "<td ><u><b>RECEIPT USED</b>&nbsp;[For Other Attendance/Visit]</u></td>"
                            strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>QTY&nbsp;&nbsp</u></td>"
                            strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>&nbsp;&nbsp</u></td>"
                            strBillDesc = strBillDesc & "<td align=""right"" style=""font-weight: bold""><u>&nbsp;&nbspUSED&nbsp;AMT</u></td>"
                            strBillDesc = strBillDesc & "</tr>"
                        End If

                        strBillDesc = strBillDesc & "<tr>"
                        strBillDesc = strBillDesc & "<td>" & CStr(uCnt) & "</td>"
                        strBillDesc = strBillDesc & "<td>[REC#&nbsp;:&nbsp;" & UCase(.fields("ReceiptID")) & "&nbsp;]&nbsp;" & dsc & "</td>"
                        strBillDesc = strBillDesc & "<td align=""right"">" & CStr(cnt2) & "</td>"
                        strBillDesc = strBillDesc & "<td align=""right"">-</td>"
                        strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(usd), 2, , , -1)) & "</td>"
                        strBillDesc = strBillDesc & "</tr>"
                    End If

                    .movenext
                Loop

                If uCnt > 0 Then
                    strBillDesc = strBillDesc & "<tr>"
                    strBillDesc = strBillDesc & "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
                    strBillDesc = strBillDesc & "</tr>"

                    'Grand Total
                    strBillDesc = strBillDesc & "<tr>"
                    strBillDesc = strBillDesc & "<td></td>"
                    strBillDesc = strBillDesc & "<td colspan=""3"" align=""left"">TOTAL RECEIPT USED</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(gUsed), 2, , , -1)) & "</td>"
                    strBillDesc = strBillDesc & "</tr>"

                    strBillDesc = strBillDesc & "<tr>"
                    strBillDesc = strBillDesc & "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
                    strBillDesc = strBillDesc & "</tr>"
                End If
            End If

            .Close
        End With

    End If 'SqlOk

    Set rst = Nothing
    Set rst2 = Nothing
End Sub

'DisplayAdmitDate
Sub DisplayAdmitDate(vst)

    Dim rst, sql, ot, cnt, hdr, adm, chg, dys, aDt, dDt, recCnt

    Set rst = CreateObject("ADODB.Recordset")
    sql = "select * from admission where visitationid='" & vst & "' and admissionstatusid<>'A003' order by admissiondate"

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            recCnt = .RecordCount
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

                strBillDesc = strBillDesc & "<tr>"
                strBillDesc = strBillDesc & "<td name=""tdLabelInpInsuranceNo"" id=""tdLabelInpInsuranceNo"" style=""font-weight: bold"">" & CStr(cnt) & ".&nbsp;&nbsp;Admission Date</td>"
                strBillDesc = strBillDesc & "<td width=""20""></td>"
                strBillDesc = strBillDesc & "<td name=""tdInputInpInsuranceNo"" id=""tdInputInpInsuranceNo"">" & aDt & "</td>"
                strBillDesc = strBillDesc & "<td width=""20""></td>"

                If cnt < recCnt Then
                    strBillDesc = strBillDesc & "<td name=""tdLabelInpVisitDate"" id=""tdLabelInpVisitDate"" style=""font-weight: bold"">Trans-Out Date</td>"
                Else
                    strBillDesc = strBillDesc & "<td name=""tdLabelInpVisitDate"" id=""tdLabelInpVisitDate"" style=""font-weight: bold"">Discharge Date</td>"
                End If

                strBillDesc = strBillDesc & "<td width=""20""></td>"
                strBillDesc = strBillDesc & "<td name=""tdInputInpVisitDate"" id=""tdInputInpVisitDate"">" & dDt & "</td>"
                strBillDesc = strBillDesc & "<td width=""20""></td>"
                strBillDesc = strBillDesc & "</tr>"
                .movenext
            Loop

        End If

        .Close
    End With

    Set rst = Nothing
End Sub

' 'AddAdmission
' Sub AddAdmission_20220509(vst)

'     Dim rst, sql, ot, cnt, hdr, adm, chg, dys, aDt, dDt, gADt, gDDt, dyCnt, gDyCnt, recCnt

'     Set rst = CreateObject("ADODB.Recordset")
'     sql = "select * from admission where visitationid='" & vst & "' and admissionstatusid<>'A003' order by admissiondate"
'     cnt = pos

'     With rst
'         .open qryPro.FltQry(sql), conn, 3, 4

'         If .RecordCount > 0 Then
'             .MoveFirst
'             '23 Dec 2018
'             recCnt = .RecordCount
'             ''4 Dec 2018
'             gADt = Now() 'Grand /Overall Admission Date
'             gDDt = CDate("1 Jan 2017") 'Grand /Overall Dischage Date
'             'Admission
'             hdr = "Admission"
'             hdr = "Hospitality" '@bless - 23 Aug 2019'
'             strBillDesc = strBillDesc & "<tr>"
'             strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
'             strBillDesc = strBillDesc & "</tr>"

'             Do While Not .EOF
'                 cnt = cnt + 1
'                 aDt = Now() 'Current Admission Date
'                 dDt = Now() 'Current Dischage Date

'                 adm = .fields("admissionid")
'                 chg = .fields("bedcharge")
'                 dys = .fields("noofdays")

'                 If dys = 0 Then

'                     ''4 Dec 2018 Individual Admissions
'                     If IsDate(.fields("AdmissionDate")) Then
'                         'If .fields("AdmissionDate") < aDt Then
'                         aDt = .fields("AdmissionDate")
'                         'End If
'                     End If

'                     If IsDate(.fields("DischargeDate")) Then
'                         'If .fields("DischargeDate") > dDt Then
'                         dDt = .fields("DischargeDate")
'                         'End If
'                     Else
'                         dDt = Now()
'                     End If

'                     If (IsDate(aDt) And IsDate(dDt)) Then
'                         dys = DateDiff("h", aDt, dDt)
'                         dys = dys / 24
'                         dyCnt = Int(dys)

'                         If ((dys - dyCnt) > 0) Then
'                             'If recCnt = 1 Then
'                             dys = dyCnt + 1
'                             'else
'                             '  dys = dyCnt
'                             'End If
'                         Else
'                             dys = dyCnt
'                         End If
'                     End If
'                 End If

'                 ''gTot = gTot + (chg * dys)

'                 ''4 Dec 2018 Over all Days
'                 If IsDate(.fields("AdmissionDate")) Then
'                     If .fields("AdmissionDate") < gADt Then
'                         gADt = .fields("AdmissionDate")
'                     End If
'                 End If

'                 If IsDate(.fields("DischargeDate")) Then
'                     If .fields("DischargeDate") > gDDt Then
'                         gDDt = .fields("DischargeDate")
'                     End If

'                 Else
'                     gDDt = Now()
'                 End If

'                 'Compile Receipt No for Credit Patients
'                 'If UCase(rTyp) = "R002" Then 'Compile Receipt No for Credit Patients
'                 If Not IsNull(.fields("AdmissionInfo1")) Then
'                     If Len(Trim(.fields("AdmissionInfo1"))) > 0 Then
'                         recNo = recNo & "," & Trim(.fields("AdmissionInfo1"))
'                     End If
'                 End If

'                 'End If

'                 strBillDesc = strBillDesc & "<tr>"
'                 strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
'                 strBillDesc = strBillDesc & "<td>" & GetComboName("Ward", .fields("wardid")) & " [" & GetComboName("AdmissionType", .fields("AdmissionTypeid")) & "]</td>"
'                 strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & CStr(dys) & " Days</td>"
'                 strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
'                 strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & (FormatNumber(CStr(chg * dys), 2, , , -1)) & "</td>"
'                 strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
'                 strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
'                 strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
'                 strBillDesc = strBillDesc & "</tr>"
'                 .MoveNext
'             Loop

'             ''4 Dec 2018
'             If (IsDate(gADt) And IsDate(gDDt)) Then
'                 gDys = DateDiff("h", gADt, gDDt)
'                 gDys = gDys / 24
'                 gDyCnt = Int(gDys)

'                 If ((gDys - gDyCnt) > 0) Then
'                     gDys = gDyCnt + 1
'                 End If
'             End If

'             gTot = gTot + (chg * gDys)
'             strBillDesc = strBillDesc & "<tr>"
'             strBillDesc = strBillDesc & "<td>&nbsp;</td>"
'             ' response.write "<td><b>Total Admissions</b></td>"
'             strBillDesc = strBillDesc & "<td><b>Total Hospitality</b></td>" '@bless - 23 Aug 2019'
'             strBillDesc = strBillDesc & "<td align=""right"">" & CStr(gDys) & " Days</td>"
'             strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
'             strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(chg * gDys), 2, , , -1)) & "</td>"
'             strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
'             strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
'             strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
'             strBillDesc = strBillDesc & "</tr>"
'         End If

'         .Close
'     End With

'     Set rst = Nothing
'     pos = cnt
' End Sub

Sub AddAdmission(vst)
    Dim rst, sql, ot, cnt, hdr, adm, chg, dys, aDt, dDt, gADt, gDDt, dyCnt, gDyCnt, recCnt
    Dim bdTyp

    Set rst = CreateObject("ADODB.Recordset")
    sql = "select * from admission where visitationid='" & vst & "' and admissionstatusid<>'A003' order by admissiondate"
    cnt = pos
    Dim hourVIP

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            '23 Dec 2018
            recCnt = .RecordCount
            ''4 Dec 2018
            gADt = Now() 'Grand /Overall Admission Date
            gDDt = CDate("1 Jan 2017") 'Grand /Overall Dischage Date
            hourVIP = 0
            'Admission
            hdr = "Admission"
            hdr = "Hospitality [Bed and Ward Services]" '@bless - 23 Aug 2019'
            hdr = "<h6>" & hdr & "</h6>"
            strBillDesc = strBillDesc & "<tr>"
            ' strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
            strBillDesc = strBillDesc & "<td colspan=""8"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
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
                        ''@bless 5 Jun 2022 //If the duration is less than an hour, so that DateDiff(hour) is less than 1
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
                ' If UCase(.fields("BedTypeID")) = UCase("VIP") Then ''AKOSOMBO
                '     hourVIP = hourVIP + DateDiff("h", dtAdm, dtDsch)
                '     bdTyp = "VIP"
                ' ElseIf UCase(.fields("BedTypeID")) = UCase("ECO") Then ''ACCRA
                '     hourVIP = hourVIP + DateDiff("h", dtAdm, dtDsch)
                '     bdTyp = "ECO"
                ' End If

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
                strBillDesc = strBillDesc & "<td colspan=""7"">" & GetComboName("Ward", .fields("wardid")) & " [" & GetComboName("AdmissionType", .fields("AdmissionTypeid")) & "]</td>"
                ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & CStr(dys) & " Days</td>"
                ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
                ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & (FormatNumber(CStr(chg * dys), 2, , , -1)) & "</td>"
                ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
                ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
                ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
                strBillDesc = strBillDesc & "</tr>"
                .movenext
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
            styl = ""
            If gDys = 0 Then
                styl = " background-color:#ff8b8b;font-weight:bold; "
            End If

            ' gTot = gTot + (chg * gDys)
            strBillDesc = strBillDesc & "<tr>"
            strBillDesc = strBillDesc & "<td>&nbsp;</td>"
            ' response.write "<td><b>Total Admissions</b></td>"
            strBillDesc = strBillDesc & "<td style=""" & styl & """ colspan=""2"">Hospitality</td>" '@bless - 23 Aug 2019'
            strBillDesc = strBillDesc & "<td style=""" & styl & """ colspan=""2"" align=""right"">" & CStr(gDys) & " Days</td>"
            ' strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
            ' strBillDesc = strBillDesc & "<td style=""" & styl & """ align=""right"">" & (FormatNumber(CStr(chg * gDys), 2, , , -1)) & "</td>"
            ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
            ' strBillDesc = strBillDesc & "<td style=""" & styl & """ align=""right""></td>"
            strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
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
            ' ' strBillDesc = strBillDesc & "<tr>"
            ' ' strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
            ' ' strBillDesc = strBillDesc & "<td>" & "VIP Hospitality" & "</td>"
            ' ' ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
            ' ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
            ' ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & CStr(dys) & " Days</td>"
            ' ' strBillDesc = strBillDesc & "<td align=""right"">" & FormatNumber(CStr(vipChg * vipDays), 2, , , -1) & "</td>"
            ' ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>" ' & (FormatNumber(CStr(chg * dys), 2, , , -1)) & "</td>"
            ' ' strBillDesc = strBillDesc & "</tr>"

            ' ' strBillDesc = strBillDesc & "<tr>"
            ' ' strBillDesc = strBillDesc & "<td>&nbsp;</td>"
            ' ' ' response.write "<td><b>Total Admissions</b></td>"
            ' ' strBillDesc = strBillDesc & "<td><b>Total Hospitality</b></td>" '@bless - 23 Aug 2019'
            ' ' strBillDesc = strBillDesc & "<td align=""right"">" & "&nbsp;" & "</td>"
            ' ' strBillDesc = strBillDesc & "<td align=""right"">" & "&nbsp;" & "</td>"
            ' ' strBillDesc = strBillDesc & "<td align=""right""><u>" & (FormatNumber(CStr((chg * gDys) + (vipChg * vipDays)), 2, , , -1)) & "</u></td>"
            ' ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
            ' ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
            ' ' strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
            ' ' strBillDesc = strBillDesc & "</tr>"

            ' strBillDesc = strBillDesc & "<tr height=""20"">"
            ' strBillDesc = strBillDesc & "<td align=""right"" colspan=""3""><b>Total " & hdr & "<b></td>"
            ' strBillDesc = strBillDesc & "<td align=""right"" colspan=""2""><u>" & (FormatNumber(CStr(subTot), 2, , , -1)) & "</u></td>"
            ' strBillDesc = strBillDesc & "</tr>"
        Else
            IsCancelledAdmission = True
        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub

Function GetAdmitReceiptNo(vst)

    Dim rs, ot, sql, rNo, cnt

    Set rs = CreateObject("ADODB.Recordset")
    ot = ""
    cnt = 0
    sql = "select AdmissionInfo1 from Admission where visitationid='" & vst & "'"

    With rs
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst

            Do While Not .EOF

                If Not IsNull(.fields("AdmissionInfo1")) Then
                    rNo = Trim(.fields("AdmissionInfo1"))

                    If Len(rNo) > 0 Then
                        ot = ot & "," & rNo
                    End If
                End If

                .movenext
            Loop

        End If

        .Close
    End With

    Set rs = Nothing
    GetAdmitReceiptNo = ot
End Function

Function GetDrugReceiptNo(vst)

    Dim rs, ot, sql, rNo, cnt

    Set rs = CreateObject("ADODB.Recordset")
    ot = ""
    cnt = 0
    sql = "select MainInfo2 from DrugSale where visitationid='" & vst & "'"

    With rs
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst

            Do While Not .EOF

                If Not IsNull(.fields("MainInfo2")) Then
                    rNo = Trim(.fields("MainInfo2"))

                    If Len(rNo) > 0 Then
                        ot = ot & "," & rNo
                    End If
                End If

                .movenext
            Loop

        End If

        .Close
    End With

    Set rs = Nothing
    GetDrugReceiptNo = ot
End Function

Function GetLabReceiptNo(vst)

    Dim rs, ot, sql, rNo, cnt

    Set rs = CreateObject("ADODB.Recordset")
    ot = ""
    cnt = 0
    sql = "select ReceiptInfo1 from LabRequest where visitationid='" & vst & "'"

    With rs
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst

            Do While Not .EOF

                If Not IsNull(.fields("ReceiptInfo1")) Then
                    rNo = Trim(.fields("ReceiptInfo1"))

                    If Len(rNo) > 0 Then
                        ot = ot & "," & rNo
                    End If
                End If

                .movenext
            Loop

        End If

        .Close
    End With

    Set rs = Nothing
    GetLabReceiptNo = ot
End Function

Function GetLabReceiptNoDearLab(vst)
    Dim rs, ot, sql, rNo, cnt

    Set rs = CreateObject("ADODB.Recordset")
    ot = ""
    cnt = 0
    sql = "select ReceiptInfo1 from LabRequest where visitationid='" & vst & "'"

    With rs
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst

            Do While Not .EOF

                If Not IsNull(.fields("ReceiptInfo1")) Then
                    rNo = Trim(.fields("ReceiptInfo1"))

                    If Len(rNo) > 0 Then
                        ot = ot & "," & rNo
                    End If
                End If

                .movenext
            Loop

        End If

        .Close
    End With

    Set rs = Nothing
    GetLabReceiptNoDearLab = ot
End Function

Function GetTreatReceiptNo(vst)
    Dim rs, ot, sql, rNo, cnt

    Set rs = CreateObject("ADODB.Recordset")
    ot = ""
    cnt = 0
    sql = "select * from ConsultReview where visitationid='" & vst & "'"

    With rs
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst

            Do While Not .EOF

                If Not IsNull(.fields("KeyPrefix")) Then
                    rNo = Trim(.fields("KeyPrefix"))

                    If Len(rNo) > 0 Then
                        ot = ot & "," & rNo
                    End If
                End If

                .movenext
            Loop

        End If

        .Close
    End With

    Set rs = Nothing
    GetTreatReceiptNo = ot
End Function

Function GetStockReceiptNo(vst)

    Dim rs, ot, sql, rNo, cnt

    Set rs = CreateObject("ADODB.Recordset")
    ot = ""
    cnt = 0
    sql = "select MainInfo1 from StockIssue where visitationid='" & vst & "'"

    With rs
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst

            Do While Not .EOF

                If Not IsNull(.fields("MainInfo1")) Then
                    rNo = Trim(.fields("MainInfo1"))

                    If Len(rNo) > 0 Then
                        ot = ot & "," & rNo
                    End If
                End If

                .movenext
            Loop

        End If

        .Close
    End With

    Set rs = Nothing
    GetStockReceiptNo = ot
End Function

'AddDrug
Sub AddDrug(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, rQty, drg, fQty, rNo
    Dim subTot, subTotSpon, subTotCash

    Set rst = CreateObject("ADODB.Recordset")
    ' sql = "select drugid, sum(qty) as qty, avg(unitcost) as unt, sum(finalamt) as tot from drugsaleitems "
    sql = "select drugid, sum(qty) as qty, avg(unitcost) as unt, sum(finalamt) as tot, sum(CAST(MainItemInfo1 AS FLOAT)) as MainItemInfo1, sum(CAST(MainInfo1 AS FLOAT)) as MainInfo1 from drugsaleitems "
    sql = sql & " where visitationid='" & vst & "' "
    sql = sql & " group by drugid "
    cnt = pos

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            subTot = 0
            subTotSpon = 0
            subTotCash = 0

            'Compile Receipt No for Credit Patients
            If UCase(rTyp) = "R002" Then 'Compile Receipt No for Credit Patients
                rNo = Trim(GetDrugReceiptNo(vst))

                If Len(rNo) > 0 Then
                    recNo = recNo & "," & rNo
                End If
            End If

            'Pharmacy
            hdr = "Medical Items"
            If InStr(1, vst, "-") > 1 Then
                hdr = hdr & " [Sponsor Not Covered (Paying)]"
            End If
            strBillDesc = strBillDesc & "<tr>"
            ' strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
            strBillDesc = strBillDesc & "<td colspan=""8"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&emsp;<h6>" & hdr & "</h6></td>"
            strBillDesc = strBillDesc & "</tr>"

            ''@bless - 03 Nov 2023 ''Calculate unit cost from total and quantity. Average of total if unit costs are different and return is done
            ' Do While Not .EOF
            '     cnt = cnt + 1
            '     unt = .fields("unt")
            '     qty = .fields("qty")
            '     drg = .fields("drugid")
            '     rQty = GetReturnQty(vst, drg)
            '     fQty = qty - rQty

            '     If fQty > 0 Then
            '         tot = fQty * unt '.Fields("tot")
            '         gTot = gTot + tot
            '         subTot = subTot + tot
            '         strBillDesc = strBillDesc & "<tr>"
            '         strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
            '         strBillDesc = strBillDesc & "<td>" & GetComboName("drug", drg) & "</td>"
            '         strBillDesc = strBillDesc & "<td align=""right"">" & CStr(fQty) & "</td>"
            '         strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
            '         strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
            '         strBillDesc = strBillDesc & "</tr>"
            '     End If

            '     .MoveNext
            ' Loop
            Do While Not .EOF
                cnt = cnt + 1
                amt = .fields("tot")
                qty = .fields("qty")
                drg = .fields("drugid")
                rAmt = GetReturnAmt(vst, drg)
                tot = amt - rAmt
                rQty = GetReturnQty(vst, drg)
                fQty = qty - rQty

                If fQty > 0 Then
                    unt = tot / fQty
                    If UCase(GetComboNameFld("Visitation", vst, "ReceiptTypeID")) = UCase("R001") Then ''Cash
                        otCash = tot
                        otSponsor = 0
                    Else
                        otCash = 0
                        otSponsor = tot
                    End If
                    ''@bless - 24 Dec 2023 //Override defaults //Get/Set Copayment amounts
                    otCash = .fields("MainInfo1")
                    otSponsor = .fields("MainItemInfo1")
                    ' If otCash > 0 Then otCash = tot ''25 Dec 2023

                    gTot = gTot + tot
                    gTotSpon = gTotSpon + otSponsor
                    gTotCash = gTotCash + otCash

                    subTot = subTot + tot
                    subTotSpon = subTotSpon + otSponsor
                    subTotCash = subTotCash + otCash
                    drgNm = GetComboName("drug", drg)

                    strBillDesc = strBillDesc & "<tr>"
                    strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
                    strBillDesc = strBillDesc & "<td>" & drgNm & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & CStr(fQty) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                    If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
                    strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                    Else
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(otSponsor), 2, , , -1)) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(otCash), 2, , , -1)) & "&nbsp;</td>"
                        ''@bless  //link for copay
                        If allowCopay Then
                            If (CDbl(amtAllowCopay) >= CDbl(otSponsor)) Then
                                If (otSponsor > 0) Then
                                    strBillDesc = strBillDesc & "<td class=""no-print"" align=""left""><button class=""btn-info no-print"" style=""margin:3px;"" title=""Co pay " & drgNm & """ onclick=""PushToCopay('" & drg & "', 'DRUGSALE', '" & drgNm & " [" & FormatNumber(otSponsor, 2) & "]" & "')"">[Copay]</button></td>"
                                Else
                                    ''If UCase(jSchd)=UCase(uName) Then
                                    If UCase(jSchd) = UCase("SystemAdmin") Or UCase(jSchd) = UCase(uName) Or UCase(GetComboNameFld("SystemUser", uName, "StaffID")) = UCase("STF001") Then
                                    strBillDesc = strBillDesc & "<td class=""no-print"" align=""left""><button class=""btn-warning no-print"" style=""margin:3px;"" title=""Undo Co pay " & drgNm & """ onclick=""PushToCopay('" & drg & "', 'DRUGSALE-UNDO', '" & drgNm & " [" & FormatNumber(otSponsor, 2) & "]" & "')"">[Undo Copay]</button></td>"
                                    End If
                                End If
                            End If
                        End If
                    End If
                    strBillDesc = strBillDesc & "</tr>"
                End If

                .movenext
            Loop

            strBillDesc = strBillDesc & "<tr height=""20"" class=""sub-total"">"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""3"">Total " & hdr & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""2"">" & (FormatNumber(CStr(subTot), 2, , , -1)) & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            Else
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""1"">" & (FormatNumber(CStr(subTotSpon), 2, , , -1)) & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""1"">" & (FormatNumber(CStr(subTotCash), 2, , , -1)) & "</td>"
            End If
            strBillDesc = strBillDesc & "</tr>"
        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub

'AddDrug2
Sub AddDrug2(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, rQty, drg, fQty, rNo
    Dim subTot, subTotSpon, subTotCash

    Set rst = CreateObject("ADODB.Recordset")
    ' ' sql = "select drugid, sum(dispenseamt1) as qty,avg(unitcost) as unt,sum(dispenseamt2) as tot from drugsaleitems2 "
    ' sql = "select drugid, sum(dispenseamt1) as qty,avg(unitcost) as unt, sum(dispenseamt2) as tot, sum(MainInfo1) as MainInfo1, sum(DispenseInfo1) as DispenseInfo1 from drugsaleitems2 "
    sql = "select drugid, sum(dispenseamt1) as qty,avg(unitcost) as unt, sum(dispenseamt2) as tot, sum(CAST(MainInfo1 AS FLOAT)) as MainInfo1, sum(CAST(DispenseInfo1 AS FLOAT)) as DispenseInfo1 from drugsaleitems2 "
    sql = sql & " where visitationid='" & vst & "' group by drugid"
    cnt = pos

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            subTot = 0
            subTotSpon = 0
            subTotCash = 0

            'Compile Receipt No for Credit Patients
            If UCase(rTyp) = "R002" Then 'Compile Receipt No for Credit Patients
                rNo = Trim(GetDrugReceiptNo(vst))

                If Len(rNo) > 0 Then
                    recNo = recNo & "," & rNo
                End If
            End If

            'Pharmacy
            hdr = "Medical Items"
            If InStr(1, vst, "-") > 1 Then
                hdr = hdr & " [Sponsor Not Covered (Paying)]"
            End If
            strBillDesc = strBillDesc & "<tr>"
            ' strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"

            strBillDesc = strBillDesc & "<td colspan=""8"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&emsp;<h6>" & hdr & "</h6></td>"
            strBillDesc = strBillDesc & "</tr>"

            ''@bless - 03 Nov 2023 ''Calculate unit cost from total and quantity. Average of total if unit costs are different and return is done
            ' Do While Not .EOF
            '     cnt = cnt + 1
            '     unt = .fields("unt")
            '     qty = .fields("qty")
            '     drg = .fields("drugid")
            '     rQty = GetReturnQty2(vst, drg)
            '     fQty = qty - rQty

            '     If fQty > 0 Then
            '         tot = fQty * unt '.Fields("tot")
            '         gTot = gTot + tot
            '         subTot = subTot + tot
            '         strBillDesc = strBillDesc & "<tr>"
            '         strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
            '         strBillDesc = strBillDesc & "<td>" & GetComboName("drug", drg) & "</td>"
            '         strBillDesc = strBillDesc & "<td align=""right"">" & CStr(fQty) & "</td>"
            '         strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
            '         strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
            '         strBillDesc = strBillDesc & "</tr>"
            '     End If

            '     .MoveNext
            ' Loop
            Do While Not .EOF
                cnt = cnt + 1
                amt = .fields("tot")
                qty = .fields("qty")
                drg = .fields("drugid")
                rAmt = GetReturnAmt2(vst, drg)
                tot = amt - rAmt
                rQty = GetReturnQty2(vst, drg)
                fQty = qty - rQty

                If fQty > 0 Then
                    unt = tot / fQty
                    copaySpon = tot
                    copayCash = 0
                    ''@bless - 24 Dec 2023 //Override defaults //Get/Set Copayment amounts
                    copayCash = .fields("MainInfo1")
                    ' If otCash > tot Then otCash = tot ''25 Dec 2023
                    copaySpon = .fields("DispenseInfo1")
                    drgNm = GetComboName("drug", drg)

                    subTot = subTot + tot
                    subTotSpon = subTotSpon + copaySpon
                    subTotCash = subTotCash + copayCash

                    gTot = gTot + tot
                    gTotSpon = gTotSpon + copaySpon
                    gTotCash = gTotCash + copayCash

                    strBillDesc = strBillDesc & "<tr>"
                    strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
                    strBillDesc = strBillDesc & "<td>" & drgNm & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & CStr(fQty) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                    If hasDiscount Then
                    strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                    Else
                    strBillDesc = strBillDesc & "<td align=""right""></td>"
                    End If

                    If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
                    strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                    Else
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(copaySpon), 2, , , -1)) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(copayCash), 2, , , -1)) & "</td>"
                        ''@bless  //link for copay
                        If allowCopay Then
                            If (CDbl(amtAllowCopay) >= CDbl(copaySpon)) Then
                                If (copaySpon > 0) Then
                                    strBillDesc = strBillDesc & "<td class=""no-print"" align=""left""><button class=""btn-info no-print"" style=""margin:3px;"" title=""Co pay " & drgNm & """ onclick=""PushToCopay('" & drg & "', 'DRUGSALE', '" & drgNm & " [" & FormatNumber(otSponsor, 2) & "]" & "')"">[Copay]</button></td>"
                                Else
                                    ''If UCase(jSchd)=UCase(uName) Then
                                    If UCase(jSchd) = UCase("SystemAdmin") Or UCase(jSchd) = UCase(uName) Or UCase(GetComboNameFld("SystemUser", uName, "StaffID")) = UCase("STF001") Then
                                    strBillDesc = strBillDesc & "<td class=""no-print"" align=""left""><button class=""btn-warning no-print"" style=""margin:3px;"" title=""Undo Co pay " & drgNm & """ onclick=""PushToCopay('" & drg & "', 'DRUGSALE-UNDO', '" & drgNm & " [" & FormatNumber(otSponsor, 2) & "]" & "')"">[Undo Copay]</button></td>"
                                    End If
                                End If
                            End If
                        End If
                    End If
                    strBillDesc = strBillDesc & "</tr>"
                End If

                .movenext
            Loop

            strBillDesc = strBillDesc & "<tr height=""20"" class=""sub-total"">"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""3"">Total " & hdr & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""2"">" & (FormatNumber(CStr(subTot), 2, , , -1)) & "</td>"
            If hasDiscount Then
                strBillDesc = strBillDesc & "<td align=""right""></td>"
            Else
                strBillDesc = strBillDesc & "<td align=""right""></td>"
            End If

            If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            Else
            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(subTotSpon), 2, , , -1)) & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(subTotCash), 2, , , -1)) & "</td>"
            End If
            strBillDesc = strBillDesc & "</tr>"

        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub

Function IsPrescriptionDrug(ins, drg, sTyp)
  Dim ot, sql, rstDst
  Set rstDst = CreateObject("ADODB.Recordset")
  ot = False
  sql = "select * from drugpricematrix2 where insurancetypeid='" & ins & "'"
  sql = sql & " and drugstoreTypeid='" & sTyp & "'"
  sql = sql & " and drugid='" & drg & "' and PermuteStatusID='P001' "
  rstDst.open qryPro.FltQry(sql), conn, 3, 4

  If rstDst.RecordCount > 0 Then
    rstDst.movefirst
    ot = True
  End If
  rstDst.Close
  Set rstDst = Nothing
  IsPrescriptionDrug = ot
End Function

Function IsPrescriptionDrugNHIS(drg, sTyp)
  Dim ot, sql, rstDst
  Set rstDst = CreateObject("ADODB.Recordset")
  ot = False
  sql = "select * from drugpricematrix2 where (InsuranceTypeID IN ('I101','I300','I500','I004') Or (InsuranceTypeID LIKE 'NHIS-%'))  "
  sql = sql & " and drugstoreTypeid='" & sTyp & "'"
  sql = sql & " and drugid='" & drg & "' and PermuteStatusID='P001' "
  rstDst.open qryPro.FltQry(sql), conn, 3, 4

  If rstDst.RecordCount > 0 Then
    rstDst.movefirst
    ot = True
  End If
  rstDst.Close
  Set rstDst = Nothing
  IsPrescriptionDrugNHIS = ot
End Function


'GetReturnQty
Function GetReturnQty(vst, dg)
    Dim rstTblSql, sql, ot
    Set rstTblSql = CreateObject("ADODB.Recordset")
    ot = 0
    With rstTblSql
        sql = "select sum(returnqty) as sm from drugreturnitems where visitationid='" & vst & "' and drugid='" & dg & "'"
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
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
'GetReturnQty
Function GetReturnAmt(vst, dg)
    Dim rstTblSql, sql, ot
    Set rstTblSql = CreateObject("ADODB.Recordset")
    ot = 0
    With rstTblSql
        sql = "select sum(FinalAmt) as sm from drugreturnitems where visitationid='" & vst & "' and drugid='" & dg & "'"
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .movefirst
            ot = .fields("sm")
            If IsNull(ot) Then
                ot = 0
            End If
        End If
        .Close
    End With
    Set rstTblSql = Nothing
    GetReturnAmt = ot
End Function
Function GetReturnAmt3(vst, dg, drgSl)
    Dim rstTblSql, sql, ot
    Set rstTblSql = CreateObject("ADODB.Recordset")
    ot = 0
    With rstTblSql
        sql = "select sum(FinalAmt) as sm from drugreturnitems where visitationid='" & vst & "' and drugid='" & dg & "' and DrugSaleID='" & drgSl & "'"
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .movefirst
            ot = .fields("sm")
            If IsNull(ot) Then
                ot = 0
            End If
        End If
        .Close
    End With
    Set rstTblSql = Nothing
    GetReturnAmt3 = ot
End Function



'GetReturnQty2
Function GetReturnQty2(vst, dg)
    Dim rstTblSql, sql, ot
    Set rstTblSql = CreateObject("ADODB.Recordset")
    ot = 0
    With rstTblSql
        sql = "select sum(returnqty) as sm from drugreturnitems2 where visitationid='" & vst & "' and drugid='" & dg & "'"
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .movefirst
            ot = .fields("sm")
            If IsNull(ot) Then
                ot = 0
            End If
        End If
        .Close
    End With
    Set rstTblSql = Nothing
    GetReturnQty2 = ot
End Function

Function GetReturnAmt2(vst, dg)
    Dim rstTblSql, sql, ot
    Set rstTblSql = CreateObject("ADODB.Recordset")
    ot = 0
    With rstTblSql
        sql = "select sum(MainItemValue1) as sm from drugreturnitems2 where visitationid='" & vst & "' and drugid='" & dg & "'"
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .movefirst
            ot = .fields("sm")
            If IsNull(ot) Then
                ot = 0
            End If
        End If
        .Close
    End With
    Set rstTblSql = Nothing
    GetReturnAmt2 = ot
End Function
Function GetReturnAmt4(vst, dg, drgSl)
    Dim rstTblSql, sql, ot
    Set rstTblSql = CreateObject("ADODB.Recordset")
    ot = 0
    With rstTblSql
        sql = "select sum(MainItemValue1) as sm from drugreturnitems2 where visitationid='" & vst & "' and drugid='" & dg & "' and DrugSaleID='" & drgSl & "'"
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .movefirst
            ot = .fields("sm")
            If IsNull(ot) Then
                ot = 0
            End If
        End If
        .Close
    End With
    Set rstTblSql = Nothing
    GetReturnAmt4 = ot
End Function


'AddNonDrug
Sub AddNonDrug(vst)

    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot

    Dim subTot

    Set rst = CreateObject("ADODB.Recordset")
    sql = "select itemid,sum(qty) as qty,avg(retailunitcost) as unt,sum(finalamt) as tot from stockissueitems where visitationid='" & vst & "' group by itemid"
    ' sql = "select itemid,sum(qty) as qty,avg(retailunitcost) as unt,sum(finalamt) as tot from stockissueitems "
    ' sql = sql & " where (visitationid='" & vst & "' Or visitationid='" & vst & "-C') group by itemid "
    cnt = pos

    'Compile Receipt No for Credit Patients
    If UCase(rTyp) = "R002" Then 'Compile Receipt No for Credit Patients
        rNo = Trim(GetStockReceiptNo(vst))

        If Len(rNo) > 0 Then
            recNo = recNo & "," & rNo
        End If
    End If

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            subTot = 0
            'Non Drug
            hdr = "Consummables"
            If InStr(1, vst, "-") > 1 Then
                hdr = "Consummables [Sponsor Not Covered (Paying)]"
            End If
            strBillDesc = strBillDesc & "<tr>"
            ' strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
            strBillDesc = strBillDesc & "<td colspan=""8"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&emsp;<h6>" & hdr & "</h6></td>"
            strBillDesc = strBillDesc & "</tr>"

            Do While Not .EOF
                cnt = cnt + 1
                itm = .fields("itemid")
                unt = .fields("unt")
                qty = .fields("qty")
                tot = .fields("tot")
                rQty = GetItemReturnQty(vst, itm)
                fQty = qty - rQty

                If fQty > 0 Then
                    gTot = gTot + tot
                    subTot = subTot + tot

                    strBillDesc = strBillDesc & "<tr>"
                    strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
                    strBillDesc = strBillDesc & "<td>" & GetComboName("items", .fields("itemid")) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & CStr(qty) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                    strBillDesc = strBillDesc & "</tr>"
                End If
                .movenext
            Loop

            strBillDesc = strBillDesc & "<tr height=""20"" class=""sub-total"">"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""3"">Total " & hdr & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""2"">" & (FormatNumber(CStr(subTot), 2, , , -1)) & "</td>"
            strBillDesc = strBillDesc & "</tr>"
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
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        If Not IsNull(rst.fields("qty")) Then
            ot = rst.fields("qty")
        End If
    End If
    rst.Close
    Set rst = Nothing
    GetItemReturnQty = ot
End Function

'AddNHISInsurance
Sub AddNHISInsurance(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, typ, ag, gen, cat, dsp

    Set rst = CreateObject("ADODB.Recordset")
    ' sql = "select  treatmentcost,diseaseTypeid,DiseaseCategoryID,genderid,agegroupid from diagnosis where insurancetypeid='I101' and visitationid='" & vst & "' order by treatmentcost desc"
    sql = "select  treatmentcost,diseaseTypeid,DiseaseCategoryID,genderid,agegroupid from diagnosis where "
    sql = sql & " (InsuranceTypeID IN ('I101', 'I300', 'I500', 'I004') or InsuranceTypeID LIKE 'NHIS-%') "
    sql = sql & " and visitationid='" & vst & "' order by treatmentcost desc"
    cnt = pos

    'rst.MaxRecords = 1 'Now multiple procedure
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            'Insurance
            hdr = "Insurance"
            hdr = "<h6>" & hdr & "</h6>"
            dsp = "Patient Insurance Coverage"

            If Not IsNull(.fields("treatmentcost")) Then
                '  If cnt = pos Then
                '    strBillDesc = strBillDesc & "<tr>"
                '    strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
                '    strBillDesc = strBillDesc & "</tr>"
                '  End If
                'cnt = cnt + 1
                tot = .fields("treatmentcost")
                gIns = gIns + tot

                'Check Exempt list
                typ = Trim(.fields("DiseaseTypeID"))
                gen = Trim(.fields("GenderID"))
                ag = Trim(.fields("AgegroupID"))
                '19 Jun 2016 'No more NHIS waiver
                '  If (UCase(typ) = "OBGY32") Or (UCase(typ) = "OBGY34") Or (UCase(typ) = "OBGY06") Or (UCase(typ) = "OBGY09") Then
                '    If UCase(gen) = "GEN02" Then
                '      wvPer = 100
                '    End If
                '  ElseIf (Left(UCase(typ), 4) = "PAED") And (UCase(typ) <> "PAED06") Then
                '    If UCase(ag) = "A001" Then
                '      wvPer = 100
                '    End If
                '  End If
            End If

            'Multiple Procedure
            .movenext

            Do While Not .EOF

                If Not IsNull(.fields("treatmentcost")) Then
                    tot = .fields("treatmentcost")
                    cat = Trim(.fields("DiseaseCategoryID"))
                    typ = Trim(.fields("DiseaseTypeID"))
                    gen = Trim(.fields("GenderID"))
                    ag = Trim(.fields("AgegroupID"))

                    If (UCase(cat) = "ASUR") Or (UCase(cat) = "PSUR") Or (UCase(cat) = "RSUR") Then
                        dsp = "Patient Insurance Coverage [Multiple Procedure]"
                        gIns = gIns + (tot * 0.2)
                    ElseIf (UCase(cat) = "ORTH") Or (UCase(cat) = "DENT") Then
                        dsp = "Patient Insurance Coverage [Multiple Procedure]"
                        gIns = gIns + (tot * 0.2)
                    End If
                End If

                .movenext
            Loop

        End If

        .Close
    End With

    strBillDesc = strBillDesc & "<tr>"
    strBillDesc = strBillDesc & "<td></td>"
    strBillDesc = strBillDesc & "<td>" & dsp & "</td>"
    strBillDesc = strBillDesc & "<td align=""right""></td>"
    strBillDesc = strBillDesc & "<td align=""right""></td>"
    strBillDesc = strBillDesc & "<td align=""right"">-" & (FormatNumber(CStr(gIns), 2, , , -1)) & "</td>"
    strBillDesc = strBillDesc & "</tr>"
    Set rst = Nothing
    pos = cnt
End Sub

'AddWaiver
Sub AddWaiver(vst)

    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot

    Set rst = CreateObject("ADODB.Recordset")
    sql = "select sum(PaidAmount) as tot from patientwaiveritems where visitationid='" & vst & "'"
    cnt = pos

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            'Waiver
            hdr = "Waiver"
            hdr = "<h6>" & hdr & "</h6>"

            Do While Not .EOF

                If Not IsNull(.fields("tot")) Then
                    '  If cnt = pos Then
                    '    strBillDesc = strBillDesc & "<tr>"
                    '    strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
                    '    strBillDesc = strBillDesc & "</tr>"
                    '  End If
                    'cnt = cnt + 1
                    tot = .fields("tot")
                    gWv = gWv + tot
                    strBillDesc = strBillDesc & "<tr>"
                    strBillDesc = strBillDesc & "<td></td>"
                    strBillDesc = strBillDesc & "<td>Patient Bill Waiver</td>"
                    strBillDesc = strBillDesc & "<td align=""right""></td>"
                    strBillDesc = strBillDesc & "<td align=""right""></td>"
                    strBillDesc = strBillDesc & "<td align=""right"">-" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                    strBillDesc = strBillDesc & "</tr>"
                End If

                .movenext
            Loop

        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub

'AddWaiver2
Sub AddWaiver2(vst)

    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, typ

    Set rst = CreateObject("ADODB.Recordset")
    sql = "select woundtypeid,sum(woundAmt2) as tot from patientwound where visitationid='" & vst & "' and woundStatusID='W001'"
    sql = sql & " group by woundTypeID"
    cnt = pos

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            'Waiver
            hdr = "Waiver / Rebate"
            hdr = "<h6>" & hdr & "</h6>"

            Do While Not .EOF

                If Not IsNull(.fields("tot")) Then
                    '  If cnt = pos Then
                    '    strBillDesc = strBillDesc & "<tr>"
                    '    strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
                    '    strBillDesc = strBillDesc & "</tr>"
                    '  End If
                    'cnt = cnt + 1
                    tot = .fields("tot")
                    typ = .fields("WoundTypeID")
                    gWv = gWv + tot
                    ' response.write "<tr>"
                    ' response.write "<td></td>"
                    ' response.write "<td>Patient " & GetComboName("WoundType", typ) & "</td>"
                    ' response.write "<td align=""right""></td>"
                    ' response.write "<td align=""right""></td>"
                    ' response.write "<td align=""right"">-" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                    ' response.write "</tr>"
                    '@bless - 11 May 2020'
                    strBillDesc = strBillDesc & "<tr>"
                    strBillDesc = strBillDesc & "<td></td>"
                    Select Case UCase(typ)
                        Case "T006"
                            strBillDesc = strBillDesc & "<td colspan=""3"">Patient " & GetComboName("WoundType", typ) & " ends " & FormatDate(GetUndertakingDate(vst)) & " " & "</td>"
                            ' response.write "<td align=""right""></td>"
                            ' response.write "<td align=""right""></td>"
                            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                        Case Else
                            strBillDesc = strBillDesc & "<td>Patient " & GetComboName("WoundType", typ) & "</td>"
                            strBillDesc = strBillDesc & "<td align=""right""></td>"
                            strBillDesc = strBillDesc & "<td align=""right""></td>"
                            strBillDesc = strBillDesc & "<td align=""right"">-" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                    End Select
                    strBillDesc = strBillDesc & "</tr>"
                End If

                .movenext
            Loop

        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub

'AddBloodBankNotice
Sub AddBloodBankNotice(vst)

    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, typ

    Set rst = CreateObject("ADODB.Recordset")
    sql = "select treatmentid,sum(qty) as qty,avg(unitcost) as unt,sum(initamt) as tot,max(MainValue2) as dlyChr "
    sql = sql & " from treatcharges where TreatmentID='NDR065' and (visitationid='" & vst & "' or visitationid='" & vst & "-C') group by treatmentid"
    cnt = pos

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            'Waiver
            hdr = "Owed Blood Bank"
            hdr = "<h6>" & hdr & "</h6>"

            Do While Not .EOF

                If Not IsNull(.fields("qty")) Then
                    If cnt = pos Then
                       strBillDesc = strBillDesc & "<tr style=""color:red;font-weight:bold;"">"
                       strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
                       strBillDesc = strBillDesc & "</tr>"
                    End If
                    cnt = cnt + 1
                    qty = .fields("qty")
                    typ = .fields("treatmentid")

                    strBillDesc = strBillDesc & "<tr style=""color:red;font-weight:bold;"">"
                    strBillDesc = strBillDesc & "<td><!-- " & cnt & " --></td>"
                    strBillDesc = strBillDesc & "<td>" & UCase(GetComboName("Treatment", typ)) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(qty), 0, , , -1)) & "</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">-</td>"
                    strBillDesc = strBillDesc & "<td align=""right"">-</td>"
                    strBillDesc = strBillDesc & "</tr>"
                End If

                .movenext
            Loop

        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub


'AddLab
Sub AddLab(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, rNo
    Dim subTot, subTotSpon, subTotCash

    Set rst = CreateObject("ADODB.Recordset")
    ' sql = "select TestGroupID, LabTestID, sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from investigation where visitationid='" & vst & "'"
    sql = "select TestGroupID, LabTestID, sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot, sum(ReceiptAmt1) as CoPayCash, sum(ReceiptAmt2) as CoPaySponsor from investigation where visitationid='" & vst & "'"
    sql = sql & " group by LabTestID, TestGroupID "
    sql = sql & " order by TestGroupID "
    cnt = pos
  ' sql = "select labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot, sum(ReceiptAmt1) as CoPayCash, sum(ReceiptAmt2) as CoPaySponsor from ("

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            subTot = 0
            subTotSpon = 0
            subTotCash = 0
            TestGroupID = "-"

            'Compile Receipt No for Credit Patients
            If UCase(rTyp) = "R002" Then 'Compile Receipt No for Credit Patients
                rNo = Trim(GetLabReceiptNo(vst))

                If Len(rNo) > 0 Then
                    recNo = recNo & "," & rNo
                End If
            End If

            'Investigations
            hdr = "Investigations"
            If InStr(1, vst, "-") > 1 Then
                hdr = hdr & " [Sponsor Not Covered (Paying)]"
            End If
            ' hdr = "<h6>" & hdr & "</h6>"
            strBillDesc = strBillDesc & "<tr>"
            ' strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
            strBillDesc = strBillDesc & "<td colspan=""8"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&emsp;<h6>" & hdr & "</h6></td>"
            strBillDesc = strBillDesc & "</tr>"

            Do While Not .EOF
                cnt = cnt + 1
                unt = .fields("unt")
                qty = .fields("qty")
                tot = .fields("tot")

                If UCase(TestGroupID) <> UCase(.fields("TestGroupID")) Then
                    TestGroupID = .fields("TestGroupID")
                    strBillDesc = strBillDesc & "<tr>"
                    ' strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
                    strBillDesc = strBillDesc & "<td colspan=""8"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&emsp;" & GetComboName("TestGroup", TestGroupID) & "</td>"
                    strBillDesc = strBillDesc & "</tr>"
                End If

                ' copaySpon = tot
                ' copayCash = 0
                copaySpon = .fields("CoPaySponsor")
                copayCash = .fields("CoPayCash")

                subTot = subTot + tot
                subTotSpon = subTotSpon + copaySpon
                subTotCash = subTotCash + copayCash

                gTot = gTot + tot
                gTotSpon = gTotSpon + copaySpon
                gTotCash = gTotCash + copayCash

                strBillDesc = strBillDesc & "<tr>"
                strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
                strBillDesc = strBillDesc & "<td>" & GetComboName("labtest", .fields("labtestid")) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & CStr(qty) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                If hasDiscount Then
                strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                Else
                strBillDesc = strBillDesc & "<td align=""right""></td>"
                End If

                If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
                strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                Else
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(copaySpon), 2, , , -1)) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(copayCash), 2, , , -1)) & "</td>"
                End If
                strBillDesc = strBillDesc & "</tr>"
                .movenext
            Loop

            strBillDesc = strBillDesc & "<tr height=""20"" class=""sub-total"">"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""3"">Total " & hdr & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""2""><u>" & (FormatNumber(CStr(subTot), 2, , , -1)) & "</u></td>"
            If hasDiscount Then
                strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            Else
                strBillDesc = strBillDesc & "<td align=""right""></td>"
            End If

            If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            Else
            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(subTotSpon), 2, , , -1)) & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(subTotCash), 2, , , -1)) & "</td>"
            End If
            strBillDesc = strBillDesc & "</tr>"
        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub


'AddLab
Sub AddLab2(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, rNo
    Dim subTot, subTotSpon, subTotCash

    Set rst = CreateObject("ADODB.Recordset")
    ' sql = "select TestGroupID, labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from investigation2 where visitationid='" & vst & "'"
    sql = "select TestGroupID, labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot, sum(ReceiptAmt1) as CoPayCash, sum(ReceiptAmt2) as CoPaySponsor from investigation2 where visitationid='" & vst & "'"
    sql = sql & " group by LabTestID, TestGroupID "
    sql = sql & " order by TestGroupID "
    cnt = pos

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            subTot = 0
            subTotSpon = 0
            subTotCash = 0
            TestGroupID = "-"


            'Compile Receipt No for Credit Patients
            If UCase(rTyp) = "R002" Then 'Compile Receipt No for Credit Patients
                rNo = Trim(GetLabReceiptNo(vst))

                If Len(rNo) > 0 Then
                    recNo = recNo & "," & rNo
                End If
            End If

            'Investigations
            hdr = "Investigations"
            If InStr(1, vst, "-") > 1 Then
                hdr = hdr & " [Sponsor Not Covered (Paying)]"
            End If
            ' hdr = "<h6>" & hdr & "</h6>"
            strBillDesc = strBillDesc & "<tr>"
            ' strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
            strBillDesc = strBillDesc & "<td colspan=""8"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&emsp;<h6>" & hdr & "</h6></td>"
            strBillDesc = strBillDesc & "</tr>"

            Do While Not .EOF
                cnt = cnt + 1
                unt = .fields("unt")
                qty = .fields("qty")
                tot = .fields("tot")

                ' copaySpon = tot
                ' copayCash = 0
                copaySpon = .fields("CoPaySponsor")
                copayCash = .fields("CoPayCash")

                subTot = subTot + tot
                subTotSpon = subTotSpon + copaySpon
                subTotCash = subTotCash + copayCash

                gTot = gTot + tot
                gTotSpon = gTotSpon + copaySpon
                gTotCash = gTotCash + copayCash

                If UCase(TestGroupID) <> UCase(.fields("TestGroupID")) Then
                    TestGroupID = .fields("TestGroupID")
                    strBillDesc = strBillDesc & "<tr>"
                    ' strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
                    strBillDesc = strBillDesc & "<td colspan=""8"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&emsp;" & GetComboName("TestGroup", TestGroupID) & "</td>"
                    strBillDesc = strBillDesc & "</tr>"
                End If

                strBillDesc = strBillDesc & "<tr>"
                strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
                strBillDesc = strBillDesc & "<td>" & GetComboName("labtest", .fields("labtestid")) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & CStr(qty) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                If hasDiscount Then
                strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                Else
                strBillDesc = strBillDesc & "<td align=""right""></td>"
                End If

                If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
                strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                Else
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(copaySpon), 2, , , -1)) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(copayCash), 2, , , -1)) & "</td>"
                End If
                strBillDesc = strBillDesc & "</tr>"
                .movenext
            Loop

            strBillDesc = strBillDesc & "<tr height=""20"" class=""sub-total"">"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""3"">Total " & hdr & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""2"">" & (FormatNumber(CStr(subTot), 2, , , -1)) & "</td>"
            If hasDiscount Then
                strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            Else
                strBillDesc = strBillDesc & "<td align=""right""></td>"
            End If

            If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            Else
            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(subTotSpon), 2, , , -1)) & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(subTotCash), 2, , , -1)) & "</td>"
            End If
            strBillDesc = strBillDesc & "</tr>"
        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub



'AddRegister
Sub AddRegister(pat, dy)

    Dim rst, sql, ot, cnt, hdr, chg

    Set rst = CreateObject("ADODB.Recordset")
    sql = "select * from Patient where patientid='" & pat & "' and firstdayid='" & dy & "'"
    cnt = pos

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            'Registration
            hdr = "Registration"
            strBillDesc = strBillDesc & "<tr>"
            strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
            strBillDesc = strBillDesc & "</tr>"

            withReg = True

            Do While Not .EOF
                cnt = cnt + 1
                adm = .fields("Patientid")
                chg = .fields("PatientValue1")

                'Compile Receipt No for Credit Patients
                If UCase(rTyp) = "R002" Then 'Compile Receipt No for Credit Patients
                    If Not IsNull(.fields("PatientInfo1")) Then
                        If Len(Trim(.fields("PatientInfo1"))) > 0 Then
                            recNo = recNo & "," & Trim(.fields("PatientInfo1"))
                        End If
                    End If
                End If

                gTot = gTot + chg
                strBillDesc = strBillDesc & "<tr>"
                strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
                strBillDesc = strBillDesc & "<td>Folder Registration</td>"
                strBillDesc = strBillDesc & "<td align=""right"">1</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
                strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
                strBillDesc = strBillDesc & "<td align=""right"">&nbsp;</td>"
                strBillDesc = strBillDesc & "</tr>"
                .movenext
            Loop

        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub

'AddTreat
Sub AddTreat(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, rNo, dailyChr, trId, trNm
    Dim subTot, disc, copayCash, copaySpon, otSponsor, otCash
    Set rst = CreateObject("ADODB.Recordset")
    ' sql = "select treatmentid,sum(qty) as qty,avg(unitcost) as unt,sum(initamt) as tot,max(MainValue2) as dlyChr from treatcharges where visitationid='" & vst & "' group by treatmentid"
    ' sql = sql & " where (visitationid='" & vst & "' or visitationid='" & vst & "-C') group by treatmentid"
    sql = "select TreatmentID, TreatTypeID, sum(qty) as qty, avg(unitcost) as unt, Sum(FinalAmt) as FinalAmt, Sum(DiscAmt) as DiscAmt, sum(InitAmt) as InitAmt, Sum(MainValue2) as CopayCash, Sum(MainValue1) as CopaySponsor  "
    sql = sql & " from TreatCharges "
    sql = sql & " where (VisitationID='" & vst & "') " '' " And TreatTypeID<>'T008' " ''Not Inpatient
    sql = sql & " group by TreatmentID,TreatTypeID "
    cnt = pos
    TreatTypeID = "-"

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            subTot = 0

            'Compile Receipt No for Credit Patients
            If UCase(rTyp) = "R002" Then 'Compile Receipt No for Credit Patients
                rNo = Trim(GetTreatReceiptNo(vst))

                If Len(rNo) > 0 Then
                    recNo = recNo & "," & rNo
                End If
            End If

            'treatment
            hdr = "Administrative Charges"
            If InStr(1, vst, "-") > 1 Then
                hdr = hdr & " [Sponsor Not Covered (Paying)]"
            End If
            ' hdr = "<h6>" & hdr & "</h6>"
            strBillDesc = strBillDesc & "<tr>"
            ' strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
            strBillDesc = strBillDesc & "<td colspan=""8"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&emsp;<h6>" & hdr & "</h6></td>"
            strBillDesc = strBillDesc & "</tr>"

            Do While Not .EOF
                cnt = cnt + 1
                dailyChr = 0
                unt = .fields("unt")
                qty = .fields("qty")
                tot = .fields("FinalAmt")
                copaySpon = .fields("CopaySponsor")
                disc = .fields("DiscAmt")
                copayCash = .fields("CopayCash")

                If UCase(TreatTypeID) <> UCase(.fields("TreatTypeID")) Then
                    TreatTypeID = .fields("TreatTypeID")
                    strBillDesc = strBillDesc & "<tr>"
                    ' strBillDesc = strBillDesc & "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & hdr & "</td>"
                    strBillDesc = strBillDesc & "<td colspan=""8"" style=""font-weight: bold"" height=""20"" valign=""bottom"">&nbsp;&nbsp;&nbsp;&nbsp;" & GetComboName("TreatType", TreatTypeID) & "</td>"
                    strBillDesc = strBillDesc & "</tr>"
                End If

                subTot = subTot + tot
                subTotDisc = subTotDisc + disc
                subTotSpon = subTotSpon + copaySpon
                subTotCash = subTotCash + copayCash

                gTot = gTot + tot
                gTotDisc = gTotDisc + disc
                gTotSpon = gTotSpon + copaySpon
                gTotCash = gTotCash + copayCash

                trId = .fields("treatmentid")
                trNm = GetComboName("treatment", trId)

                strBillDesc = strBillDesc & "<tr>"
                strBillDesc = strBillDesc & "<td>" & CStr(cnt) & "</td>"
                strBillDesc = strBillDesc & "<td>" & trNm & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & CStr(qty) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"" style=""padding-left: 6px;"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"" style=""padding-left: 6px;"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                If hasDiscount And (disc > 0) Then
                strBillDesc = strBillDesc & "<td align=""right"">" & (FormatNumber(CStr(disc), 2, , , -1)) & "</td>"
                Else
                strBillDesc = strBillDesc & "<td align=""right""></td>"
                End If

                If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
                strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
                Else
                strBillDesc = strBillDesc & "<td align=""right"" style=""padding-left: 6px;"">" & (FormatNumber(CStr(copaySpon), 2, , , -1)) & "</td>"
                strBillDesc = strBillDesc & "<td align=""right"" style=""padding-left: 6px;"">" & (FormatNumber(CStr(copayCash), 2, , , -1)) & "</td>"
                End If
                strBillDesc = strBillDesc & "</tr>"

                .movenext
            Loop

            strBillDesc = strBillDesc & "<tr height=""20"" class=""sub-total"">"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""3"">Total " & hdr & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"" colspan=""2"">" & (FormatNumber(CStr(subTot), 2, , , -1)) & "</td>"
            If hasDiscount Then
                strBillDesc = strBillDesc & "<td align=""right"">&nbsp;" & (FormatNumber(CStr(subTotDisc), 2, , , -1)) & "</td>"
            Else
                strBillDesc = strBillDesc & "<td align=""right""></td>"
            End If

            If (UCase(insGrp) = UCase("CASH") Or UCase(rTyp) = UCase("R001")) Then '' Paying/Cash
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">" & "" & "</td>"
            Else
            strBillDesc = strBillDesc & "<td align=""right"">&nbsp;" & (FormatNumber(CStr(subTotSpon), 2, , , -1)) & "</td>"
            strBillDesc = strBillDesc & "<td align=""right"">&nbsp;" & (FormatNumber(CStr(subTotCash), 2, , , -1)) & "</td>"
            End If
            strBillDesc = strBillDesc & "</tr>"
        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub


Sub AddBillReportHeader()
    Dim str

    str = ""
    str = str & "<tr><td align=""center"" style=""font-weight:bold;font-size:112pt"">"
    str = str & "<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""0"" style=""font-size:12pt"">"
    str = str & "<tr><td rowspan=""4""><img src=""images/banner1.bmp""></td>"
    str = str & "<td align=""center"">GREATER ACCRA REGIONAL HOSPITAL</td>"
    str = str & "<td rowspan=""4""><img src=""images/banner2.bmp""></td></tr>"

    str = str & "<tr><td align=""center"">ACCOUNTS DEPARTMENT</td></tr>"
    str = str & "<tr><td align=""center"">P. O. BOX 473, ACCRA</td></tr>"
    str = str & "<tr><td align=""center"">TEL: 0302-2283-82/15/48 EXT:126</td></tr>"
    str = str & "</table>"
    str = str & "</td></tr>"

    str = ""

    response.write str
End Sub

Function CleanReceiptNos(recNo)

    Dim arr, ul, num, lst, ot, arr2, ul2, num2, lst2, rec, cnt, recOk

    lst = Trim(recNo)
    lst2 = ""
    cnt = 0

    If Len(lst) > 0 Then
        arr = Split(lst, ",")
        ul = UBound(arr)

        For num = 0 To ul
            rec = Trim(arr(num))

            If Len(rec) > 0 Then
                If ReceiptCharValid(rec) Then
                    arr2 = Split(lst2, ",")
                    ul2 = UBound(arr2)
                    recOk = True

                    For num2 = 0 To ul2

                        If UCase(Trim(arr2(num2))) = UCase(Trim(rec)) Then
                            recOk = False

                            Exit For

                        End If

                    Next

                    If recOk Then
                        cnt = cnt + 1

                        If cnt > 1 Then
                            lst2 = lst2 & ","
                        End If

                        lst2 = lst2 & rec
                    End If
                End If
            End If

        Next

    End If

    CleanReceiptNos = lst2
End Function

Function ReceiptCharValid(rec)

    Dim ot, lst, lth, num, ch, pos

    ot = True
    lst = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890/\_- "
    lth = Len(rec)

    For num = 1 To lth
        ch = Mid(rec, num, 1)
        pos = InStr(1, UCase(lst), UCase(ch))

        If pos < 1 Then
            ot = False

            Exit For

        End If

    Next

    ReceiptCharValid = ot
End Function

Sub SetBillCompleted(vst)

    Dim rs, ot, sql, usr, dy, dt

    Set rs = CreateObject("ADODB.Recordset")
    ot = ""
    cnt = 0
    ' sql = "select SystemUserID,WorkingDayID,TransProcessDate1 from AdmissionPro where Visitationid='" & vst & "' and TransProcessStat2ID='T013'" 'Completed T010
    '' @bless - 27 Sept 2019 >> Exclude ward nurisng from displaying as Completing Bill...usually for Discharge-But-In'
    sql = "select SystemUserID,WorkingDayID,TransProcessDate1 from AdmissionPro where Visitationid='" & vst & "' "
    sql = sql & " and TransProcessStat2ID='T013'" 'Completed Review T013
    sql = sql & " and (JobScheduleID NOT LIKE 'W%' and TransProcessDate1 > '21 May 2019 00:00:00') "
    sql = sql & " order by TransProcessDate1 Desc "

    With rs
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            dy = Trim(.fields("WorkingDayID"))
            dt = Trim(.fields("TransProcessDate1"))
            usr = Trim(.fields("SystemUserID"))
            cmptdBy = GetComboName("Staff", GetComboNameFld("SystemUser", usr, "StaffID"))
            cmptdDt = FormatDateDetail(dt)
            cmptdSig = "Electronically Signed"
        End If

        .Close
    End With

    Set rs = Nothing
End Sub

Sub SetBillReviewed(vst)

    Dim rs, ot, sql, usr, dy, dt

    Set rs = CreateObject("ADODB.Recordset")
    ot = ""
    cnt = 0
    ' sql = "select SystemUserID,WorkingDayID,TransProcessDate1 from AdmissionPro where Visitationid='" & vst & "' and TransProcessStat2ID='T010'" 'Completed T010
    '' @bless - 27 Sept 2019 >> Exclude ward nurisng from displaying as Reviewer...usually for Discharge-But-In'
    sql = "select SystemUserID,WorkingDayID,TransProcessDate1 from AdmissionPro where Visitationid='" & vst & "' "
    sql = sql & " and TransProcessStat2ID='T010'" 'Completed T010
    sql = sql & " and (JobScheduleID NOT LIKE 'W%' and TransProcessDate1 > '21 May 2019 00:00:00') "
    sql = sql & " order by TransProcessDate1 Desc "

    With rs
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            dy = Trim(.fields("WorkingDayID"))
            dt = Trim(.fields("TransProcessDate1"))
            usr = Trim(.fields("SystemUserID"))

            If UCase(cmptdBy) = UCase("Bill Not Completed") Then ' 21 May 2019 Old Approval Process
                cmptdBy = GetComboName("Staff", GetComboNameFld("SystemUser", usr, "StaffID"))
                cmptdDt = FormatDateDetail(dt)
                cmptdSig = "Electronically Signed"

                revBy = ""
                revDt = ""
                revSig = ""
            Else
                revBy = GetComboName("Staff", GetComboNameFld("SystemUser", usr, "StaffID"))
                revDt = FormatDateDetail(dt)
                revSig = "Electronically Signed"
            End If
        End If

        .Close
    End With

    Set rs = Nothing
End Sub



' 'Clickable Url Link
'         lnkCnt = lnkCnt + 1
'         lnkID = "lnk" & CStr(lnkCnt)
'         lnkText = "<b>CLICK FOR SUMMARIZED BILL</b>"
'         'lnkUrl = "wpgSelectPrintLayout.asp?PositionForTableName=Admission&AdmissionID=" & adm
'         ' lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=Admission1&PositionForTableName=Admission&AdmissionID=" & adm
'         lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=Admission3&PositionForTableName=Admission&AdmissionID=" & adm
'         navPop = "POP"
'         inout = "IN"
'         fntSize = "9"
'         fntColor = "#ff0000"
'         bgColor = clr
'         wdth = ""
'         AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth



Sub addJS()
    Dim js
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
    js = js & "   let url = ""wpgPrtPrintLayoutAll.asp?PrintLayoutName=Admission2&PositionForTableName=Admission&AdmissionID=" & adm & """; " & vbNewLine
    js = js & "   url = url + ""&ProcessCoPay=YES&ProcessCoPayAmount=" & amtAllowCopay & """; " & vbNewLine
    js = js & "   url = url + ""&TableID="" + tbl + ""&VisitationID=" & vst & "&ItemID="" + drg; " & vbNewLine
    js = js & "   window.location.href = url;  " & vbNewLine
    js = js & "  }   " & vbNewLine
    js = js & " }   " & vbNewLine
    js = js & "</script>" & vbNewLine
    response.write js
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
End Sub
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
    Dim js
    js = "<script language=""javascript"">"
    js = js & "  function _clearEle() { " & vbNewLine
    js = js & "   var ele; " & vbNewLine
    js = js & "   ele = document.getElementById('tblMultiSelect'); " & vbNewLine
    js = js & "   if (ele) { var ele1 = ele.parentElement; if (ele1) {ele1.style.height='0px'; } } " & vbNewLine
    js = js & "  } " & vbNewLine
    js = js & "  window.onbeforeprint = _clearEle(); " & vbNewLine
    js = js & "</script>" & vbNewLine
    ' response.write js

' response.write "<tr>"
' response.write "<tr>"
' response.write "<td align=""center"">"
' response.write "<table id=""tblHiddenFields"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
' response.write "<tr>"
' response.write "<td align=""center"">"
' response.write "<table id=""tblFooter"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
' response.write "<tr>"
' response.write "<td colspan=""7"" bgcolor=""#FFFFFF"" height=""10"" style=""font-size: 8pt"" align=""right"">"
' response.write "Ridge Hospital @2017</td>"
' response.write "</tr>"
' response.write "</table>"
' response.write "</td>"
' response.write "</tr>"
' response.write "</table>"
' response.write "</td>"
' response.write "</tr>"
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
