'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim sql, rstPrn1, rstPrn2, cat, catNm, tot, vst, hdr, pos, gTot, gPaid, gUsed, recNo, pat, gWv
Set rstPrn1 = CreateObject("ADODB.Recordset")
Set rstPrn2 = CreateObject("ADODB.Recordset")

tot = 0
gTot = 0
gPaid = 0
gUsed = 0
gWv = 0
vst = Trim(GetRecordField("VisitationID"))
pat = Trim(GetRecordField("PatientID"))
recNo = Trim(GetRecordField("AdmissionInfo1"))
sql = GetTableSql("Visitation")
sql = sql & " and  Visitation.Visitationid='" & Trim(vst) & "'"

With rstPrn1
    .open qryPro.FltQry(sql), conn, 3, 4

    If .RecordCount > 0 Then
        .MoveFirst
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
        response.write "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"">"
        response.write "PATIENT BILL</td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center"">"
        response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""2"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 9pt; font-family: Arial"">"
        response.write "       <tr>"
        response.write "<td name=""tdLabelInpVisitationID"" id=""tdLabelInpVisitationID"" style=""font-weight: bold"">Visit No.</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdInputInpVisitationID"" id=""tdInputInpVisitationID"">" & (.fields("VisitationID")) & "</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdLabelInpVisitTypeID"" id=""tdLabelInpVisitTypeID"" style=""font-weight: bold"">VisitType</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdInputInpVisitTypeID"" id=""tdInputInpVisitTypeID"">" & (.fields("VisitTypeName")) & "</td>"
        response.write "<td width=""20""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpPatientID"" id=""tdLabelInpPatientID"" style=""font-weight: bold"">Patient No.</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdInputInpPatientID"" id=""tdInputInpPatientID"">" & (.fields("PatientID")) & "</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdLabelInpVisitTypeID"" id=""tdLabelInpVisitTypeID"" style=""font-weight: bold"">Patient Name</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdInputInpVisitTypeID"" id=""tdInputInpID"">" & (.fields("PatientName")) & "</td>"
        response.write "<td width=""20""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpGenderID"" id=""tdLabelInpGenderID"" style=""font-weight: bold"">Gender</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & (.fields("GenderName")) & "</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdLabelInpPatientAge"" id=""tdLabelInpPatientAge"" style=""font-weight: bold"">Age</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdInputInpPatientAge"" id=""tdInputInpPatientAge"">" & CStr(Int(CInt(.fields("PatientAge")))) & "</td>"
        response.write "<td width=""20""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpInsuredPatientID"" id=""tdLabelInpInsuredPatientID"" style=""font-weight: bold"">Billing Account No.</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdInputInpInsuredPatientID"" id=""tdInputInpInsuredPatientID"">" & (.fields("InsuredPatientID")) & "</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdLabelInpInsuranceSchemeID"" id=""tdLabelInpInsuranceSchemeID"" style=""font-weight: bold"">Billing Info.</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdInputInpInsuranceSchemeID"" id=""tdInputInpInsuranceSchemeID"">" & (.fields("InsuranceSchemeName")) & "</td>"
        response.write "<td width=""20""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpInsuranceNo"" id=""tdLabelInpInsuranceNo"" style=""font-weight: bold"">InsuranceNo</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdInputInpInsuranceNo"" id=""tdInputInpInsuranceNo"">" & (.fields("InsuranceNo")) & "</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdLabelInpVisitDate"" id=""tdLabelInpVisitDate"" style=""font-weight: bold"">Visit Date</td>"
        response.write "<td width=""20""></td>"
        response.write "<td name=""tdInputInpVisitDate"" id=""tdInputInpVisitDate"">" & (FormatDateDetail(.fields("VisitDate"))) & "</td>"
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
        response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1"" height=""20""></td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td style=""font-weight: bold""><u>NO.</u></td>"
        response.write "<td style=""font-weight: bold""><u>SERVICE DESCRIPTION</u></td>"
        response.write "<td align=""right"" style=""font-weight: bold""><u>QTY  </u></td>"
        response.write "<td align=""right"" style=""font-weight: bold""><u>UNIT  </u></td>"
        response.write "<td align=""right"" style=""font-weight: bold""><u>  TOTAL</u></td>"
        response.write "</tr>"
        'Consultation
        hdr = (.fields("SpecialistTypeName")) & " [Consultation]"
        response.write "<tr>"
        response.write "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">    Consultation</td>"
        response.write "</tr>"

        gTot = gTot + .fields("Visitcost")
        response.write "<tr>"
        response.write "<td>1.</td>"
        response.write "<td>" & hdr & "</td>"
        response.write "<td align=""right"">1</td>"
        response.write "<td align=""right"">" & (FormatNumber(CStr(.fields("Visitcost")), 2, , , -1)) & "</td>"
        response.write "<td align=""right"">" & (FormatNumber(CStr(.fields("Visitcost")), 2, , , -1)) & "</td>"
        response.write "</tr>"

        pos = 1
        AddAdmission vst
        AddDrug vst
        AddNonDrug vst
        AddLab vst
        'AddXray vst
        AddTreat vst

        response.write "<tr>"
        response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
        response.write "</tr>"

        'Grand Total
        response.write "<tr>"
        response.write "<td></td>"
        response.write "<td colspan=""3"" align=""left"">TOTAL BILL</td>"
        response.write "<td align=""right"">" & (FormatNumber(CStr(gTot), 2, , , -1)) & "</td>"
        response.write "</tr>"

        'Add Waivers
        AddWaiver vst

        If gWv > 0 Then
            response.write "<tr>"
            response.write "<td></td>"
            response.write "<td colspan=""3"" align=""left"">FINAL TOTAL BILL</td>"
            response.write "<td align=""right"">" & (FormatNumber(CStr(gTot - gWv), 2, , , -1)) & "</td>"
            response.write "</tr>"
        End If

        response.write "<tr>"
        response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
        response.write "</tr>"

        'Payments
        AddPayments pat, recNo

        'UsedPayments
        AddUsedPayments pat, recNo

        'Grand Total
        response.write "<tr>"
        response.write "<td></td>"
        response.write "<td colspan=""3"" align=""left""><b>OUTSTANDING BILL</b></td>"
        response.write "<td align=""right""><b>" & (FormatNumber(CStr(gTot - gPaid + gUsed - gWv), 2, , , -1)) & "</b></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td colspan=""5"" align=""center""><hr color=""#999999"" size=""1""></td>"
        response.write "</tr>"

        response.write "</table>"
        response.write "</td>"
        response.write "</tr>"
    End If

    .Close
End With

'AddPayments
Sub AddPayments(pat, recNo)
    Dim rst, sql, ot, cnt, hdr, dsc, pd, cn, bal, dt, sDt, eDt, vst, dDt, cn2
    Dim arr, ul, num, whcls, r, rCnt, sqlOk, sql2
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

    If (UCase(pat) = "P1") Or (UCase(pat) = "P2") Then
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

        sql = "select * from Receipt where (receiptID in (" & sql2 & "))"
        sql = sql & " " & whcls
        sql = sql & " order by receiptDate"
    End If

    cnt = 0

    If sqlOk Then

        With rst
            .open qryPro.FltQry(sql), conn, 3, 4

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

    End If 'Sql

    Set rst = Nothing
End Sub

'AddUsedPayments
Sub AddUsedPayments(pat, recNo)
    Dim rst, rst2, sql, ot, cnt, hdr, dsc, pd, cn, bal, dt, sDt, eDt, vst, dDt
    Dim cnt2, cn2, rec, usd, uCnt, sql2
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

    If (UCase(pat) = "P1") Or (UCase(pat) = "P2") Then
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

        sql2 = "select distinct receiptid from patientreceipt2 where visitationid='" & vst & "' and PatientID='" & pat & "'"
        sql = "select * from Receipt where (receiptID in (" & sql2 & "))"
        sql = sql & " " & whcls
        sql = sql & " order by receiptDate"
    End If 'P1,P2

    If sqlOk Then
        cnt = 0

        With rst
            .open qryPro.FltQry(sql), conn, 3, 4

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
                    rst2.open qryPro.FltQry(sql), conn, 3, 4

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

    End If 'SqlOk

    Set rst = Nothing
    Set rst2 = Nothing
End Sub
'AddAdmission
Sub AddAdmission(vst)
    Dim rst, sql, ot, cnt, hdr, adm, chg, dys
    Dim aDt, dDt, tmpDDt
    
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select * from admission where visitationid='" & vst & "'"
    cnt = pos
    wards = ""
    
    aDt = Now 'CDate("28 Sep 2022")
    dDt = CDate("28 Sep 2022")
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            cnt = cnt + 1
            .MoveFirst
            Do While Not .EOF
                'cnt = cnt + 1
                adm = .fields("admissionid")
                chg = .fields("bedcharge")
                dys = .fields("noofdays")
                
                'gTot = gTot + (chg * dys)
                If IsDate(.fields("AdmissionDate")) Then
                    If CDate(.fields("AdmissionDate")) < aDt Then
                        aDt = CDate(.fields("AdmissionDate"))
                    End If
                End If
                
                tmpDDt = .fields("DischargeDate")
                If IsNull(tmpDDt) Then
                    tmpDDt = Now
                End If
                
                If CDate(tmpDDt) > dDt Then
                    dDt = CDate(tmpDDt)
                End If
                
                If Len(wards) > 0 Then
                    wards = wards & ", "
                End If
                wards = wards & GetComboName("Ward", .fields("wardid"))
                .MoveNext
            Loop
        End If
        .Close
    End With

    'Admission
    hdr = "Admission"
    response.write "<tr>"
    response.write "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">" & hdr & "</td>"
    response.write "</tr>"
    
    response.write "<tr>"
    response.write "<td>" & CStr(cnt) & "</td>"
    response.write "<td>" & wards & "</td>"
    
    
    dys = DateDiff("d", aDt, dDt)
    If dys = 0 Then
        If FormatDate(aDt) = FormatDate(dDt) Then
            dys = 1
        End If
    End If
    
    gTot = gTot + (chg * dys)
    If dys > 0 Then
        response.write "<td align=""right"">" & CStr(dys) & " Days</td>"
    Else
        response.write "<td align=""right""> </td>"
    End If

    response.write "<td align=""right"">" & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
    response.write "<td align=""right"">" & (FormatNumber(CStr(chg * dys), 2, , , -1)) & "</td>"
    response.write "<td align=""right""> </td>"
    response.write "<td align=""right""> </td>"
    response.write "</tr>"
    
    Set rst = Nothing
    
    
    Exit Sub
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            
            'Admission
            hdr = "Admission"
            response.write "<tr>"
            response.write "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">" & hdr & "</td>"
            response.write "</tr>"

            Do While Not .EOF
                cnt = cnt + 1
                adm = .fields("admissionid")
                chg = .fields("bedcharge")
                dys = .fields("noofdays")
                gTot = gTot + (chg * dys)
                
                response.write "<tr>"
                response.write "<td>" & CStr(cnt) & "</td>"
                response.write "<td>" & GetComboName("Ward", .fields("wardid")) & " [" & GetComboName("AdmissionType", .fields("AdmissionTypeid")) & "]</td>"

                If dys > 0 Then
                    response.write "<td align=""right"">" & CStr(dys) & " Days</td>"
                Else
                    response.write "<td align=""right""> </td>"
                End If

                response.write "<td align=""right"">" & (FormatNumber(CStr(chg), 2, , , -1)) & "</td>"
                response.write "<td align=""right"">" & (FormatNumber(CStr(chg * dys), 2, , , -1)) & "</td>"
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
'AddDrug
Sub AddDrug(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, rQty, drg, fQty, tot2, rTot, sTot
    Set rst = CreateObject("ADODB.Recordset")

    'sql = "select drugid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from drugsaleitems where visitationid='" & vst & "' group by drugid"
    sql = "select drugid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from ("
    sql = sql & "select drugid, qty, unitcost, finalamt from drugsaleitems where visitationid='" & vst & "' "
    sql = sql & " union all "
    sql = sql & "select drugid, DispenseAmt1 as qty, unitcost, dispenseAmt2 as finalamt from drugsaleitems2 where visitationid='" & vst & "') as t "
    sql = sql & " group by drugid"

    cnt = pos

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            'Pharmacy
            hdr = "Medical Items"
            response.write "<tr>"
            response.write "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">    " & hdr & "</td>"
            response.write "</tr>"

            Do While Not .EOF
                unt = .fields("unt")
                qty = .fields("qty")
                drg = .fields("drugid")

                tot2 = .fields("tot")    'Addedd 1 Oct 2015

                rQty = GetReturnQty(vst, drg)
                fQty = qty - rQty

                If fQty > 0 Then
                    cnt = cnt + 1

                    If rQty > 0 Then 'Addedd 1 Oct 2015
                        rTot = GetReturnTot(vst, drg)
                        unt = (tot2 - rTot) / fQty
                    End If

                    tot = fQty * unt  '.Fields("tot")
                    gTot = gTot + tot
                    sTot = sTot + tot
                    response.write "<tr>"
                    response.write "<td>" & CStr(cnt) & "</td>"
                    response.write "<td>" & GetComboName("drug", drg) & "</td>"
                    response.write "<td align=""right"">" & CStr(fQty) & "</td>"
                    response.write "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
                    response.write "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                    response.write "</tr>"
                End If

                .MoveNext
            Loop
If .RecordCount > 1 Then
  response.write "<tr>"
  response.write "<td></td>"
  response.write "<td></td>"
  response.write "<td colspan=""2"" align=""right"" style='font-weight:bold;font-style:italic;'>Sub Total [" & hdr & "]</td>"
  response.write "<td align=""right"" style='font-weight:bold;font-style:italic;'>" & (FormatNumber(CStr(sTot), 2, , , -1)) & "</td>"
  response.write "</tr>"
End If
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
        sql = sql & "union all select FinalAmt, returnqty from drugreturnitems2 where visitationid='" & vst & "' and drugid='" & dg & "' "
        sql = sql & ") as t"

        .open qryPro.FltQry(sql), conn, 3, 4

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

        .open qryPro.FltQry(sql), conn, 3, 4

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
'AddNonDrug
Sub AddNonDrug(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, sTot
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select itemid,sum(qty) as qty,avg(retailunitcost) as unt,sum(finalamt) as tot from stockissueitems where visitationid='" & vst & "' group by itemid"
    cnt = pos

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .MoveFirst
            'Non Drug
            hdr = "Non Drug Consummables"
            response.write "<tr>"
            response.write "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">    " & hdr & "</td>"
            response.write "</tr>"

            Do While Not .EOF
                cnt = cnt + 1
                unt = .fields("unt")
                qty = .fields("qty")
                tot = .fields("tot")
                gTot = gTot + tot
                sTot = sTot + tot
                response.write "<tr>"
                response.write "<td>" & CStr(cnt) & "</td>"
                response.write "<td>" & GetComboName("items", .fields("itemid")) & "</td>"
                response.write "<td align=""right"">" & CStr(qty) & "</td>"
                response.write "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
                response.write "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
If .RecordCount > 1 Then
  response.write "<tr>"
  response.write "<td></td>"
  response.write "<td></td>"
  response.write "<td colspan=""2"" align=""right"" style='font-weight:bold;font-style:italic;'>Sub Total [" & hdr & "]</td>"
  response.write "<td align=""right"" style='font-weight:bold;font-style:italic;'>" & (FormatNumber(CStr(sTot), 2, , , -1)) & "</td>"
  response.write "</tr>"
End If
        End If

        .Close
    End With

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
            .MoveFirst
            'Waiver
            hdr = "Waiver"

            Do While Not .EOF

                If Not IsNull(.fields("tot")) Then
                    '  If cnt = pos Then
                    '    response.write "<tr>"
                    '    response.write "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">    " & hdr & "</td>"
                    '    response.write "</tr>"
                    '  End If
                    'cnt = cnt + 1
                    tot = .fields("tot")
                    gWv = gWv + tot
                    response.write "<tr>"
                    response.write "<td></td>"
                    response.write "<td>Patient Bill Waiver</td>"
                    response.write "<td align=""right""></td>"
                    response.write "<td align=""right""></td>"
                    response.write "<td align=""right"">-" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
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

'AddLab
Sub AddLab(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, sTot
    Set rst = CreateObject("ADODB.Recordset")

    'sql = "select labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from investigation where visitationid='" & vst & "'"
    '''sql = sql & " and testcategoryid<>'T006' and testcategoryid<>'T007' and testcategoryid<>'T008' group by labtestid"
    'sql = sql & " group by labtestid"

    sql = "select labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from ("
    sql = sql & "select labtestid, qty, unitcost, finalamt from investigation where visitationid='" & vst & "'"
    sql = sql & "union all "
    sql = sql & "select labtestid, qty, unitcost, finalamt from investigation2 where visitationid='" & vst & "') as t"

    'sql = sql & " and testcategoryid<>'T006' and testcategoryid<>'T007' and testcategoryid<>'T008' group by labtestid"
    sql = sql & " group by labtestid"

    cnt = pos

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .MoveFirst
            'Investigations
            hdr = "Investigations"
            response.write "<tr>"
            response.write "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">    " & hdr & "</td>"
            response.write "</tr>"

            Do While Not .EOF
                cnt = cnt + 1
                unt = .fields("unt")
                qty = .fields("qty")
                tot = .fields("tot")
                gTot = gTot + tot
                sTot = sTot + tot
                response.write "<tr>"
                response.write "<td>" & CStr(cnt) & "</td>"
                response.write "<td>" & GetComboName("labtest", .fields("labtestid")) & "</td>"
                response.write "<td align=""right"">" & CStr(qty) & "</td>"
                response.write "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
                response.write "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
If .RecordCount > 1 Then
  response.write "<tr>"
  response.write "<td></td>"
  response.write "<td></td>"
  response.write "<td colspan=""2"" align=""right"" style='font-weight:bold;font-style:italic;'>Sub Total [" & hdr & "]</td>"
  response.write "<td align=""right"" style='font-weight:bold;font-style:italic;'>" & (FormatNumber(CStr(sTot), 2, , , -1)) & "</td>"
  response.write "</tr>"
End If
        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub

'AddXRay
Sub AddXray(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, sTot
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from investigation where visitationid='" & vst & "'"
    sql = sql & " and (testcategoryid='T006' or testcategoryid='T007' or testcategoryid='T008') group by labtestid"
    cnt = pos

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .MoveFirst
            'X-Ray
            hdr = "X-Ray/Scan/ECG"
            response.write "<tr>"
            response.write "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">    " & hdr & "</td>"
            response.write "</tr>"

            Do While Not .EOF
                cnt = cnt + 1
                unt = .fields("unt")
                qty = .fields("qty")
                tot = .fields("tot")
                gTot = gTot + tot
                sTot = sTot + tot
                response.write "<tr>"
                response.write "<td>" & CStr(cnt) & "</td>"
                response.write "<td>" & GetComboName("labtest", .fields("labtestid")) & "</td>"
                response.write "<td align=""right"">" & CStr(qty) & "</td>"
                response.write "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
                response.write "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
If .RecordCount > 1 Then
  response.write "<tr>"
  response.write "<td></td>"
  response.write "<td></td>"
  response.write "<td colspan=""2"" align=""right"" style='font-weight:bold;font-style:italic;'>Sub Total [" & hdr & "]</td>"
  response.write "<td align=""right"" style='font-weight:bold;font-style:italic;'>" & (FormatNumber(CStr(sTot), 2, , , -1)) & "</td>"
  response.write "</tr>"
End If
        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub

'AddTreat
Sub AddTreat(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, sTot
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select treatmentid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from treatcharges where visitationid='" & vst & "' group by treatmentid"
    cnt = pos

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .MoveFirst
            'treatment
            hdr = "Services and Consumables"
            response.write "<tr>"
            response.write "<td colspan=""5"" style=""font-weight: bold"" height=""20"" valign=""bottom"">    " & hdr & "</td>"
            response.write "</tr>"

            Do While Not .EOF
                cnt = cnt + 1
                unt = .fields("unt")
                qty = .fields("qty")
                tot = .fields("tot")
                gTot = gTot + tot
                sTot = sTot + tot
                response.write "<tr>"
                response.write "<td>" & CStr(cnt) & "</td>"
                response.write "<td>" & GetComboName("treatment", .fields("treatmentid")) & "</td>"
                response.write "<td align=""right"">" & CStr(qty) & "</td>"
                response.write "<td align=""right"">" & (FormatNumber(CStr(unt), 2, , , -1)) & "</td>"
                response.write "<td align=""right"">" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
If .RecordCount > 1 Then
  response.write "<tr>"
  response.write "<td></td>"
  response.write "<td></td>"
  response.write "<td colspan=""2"" align=""right"" style='font-weight:bold;font-style:italic;'>Sub Total [" & hdr & "]</td>"
  response.write "<td align=""right"" style='font-weight:bold;font-style:italic;'>" & (FormatNumber(CStr(sTot), 2, , , -1)) & "</td>"
  response.write "</tr>"
End If
        End If

        .Close
    End With

    Set rst = Nothing
    pos = cnt
End Sub
'DisplayAdmitDate
Sub DisplayAdmitDate(vst)
    Dim rst, sql, ot, cnt, hdr, adm, chg, dys, aDt, dDt, onAdm
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select * from admission where visitationid='" & vst & "' and admissionstatusid<>'A003' order by admissiondate"

    aDt = Now
    dDt = CDate("01 Sep 2022")
    onAdm = False
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .MoveFirst
            'Admission
            cnt = 0

            Do While Not .EOF
                cnt = cnt + 1
                'aDt = ""
                'dDt = ""

                If Not IsNull(.fields("AdmissionDate")) Then
                    If IsDate(.fields("AdmissionDate")) Then
                        If .fields("AdmissionDate") < aDt Then
                            aDt = (.fields("AdmissionDate"))
                        End If
                    End If

                End If
                
                If Not IsNull(.fields("DischargeDate")) Then
                    If IsDate(.fields("DischargeDate")) Then
                        If .fields("DischargeDate") > dDt Then
                            dDt = (.fields("DischargeDate"))
                        End If
                    End If

                End If
                If .fields("AdmissionStatusID") = "A001" Then
                    onAdm = True
                End If

'                response.write "<tr>"
'                response.write "<td name=""tdLabelInpInsuranceNo"" id=""tdLabelInpInsuranceNo"" style=""font-weight: bold"">" & CStr(cnt) & ".  Admission Date</td>"
'                response.write "<td width=""20""></td>"
'                response.write "<td name=""tdInputInpInsuranceNo"" id=""tdInputInpInsuranceNo"">" & aDt & "</td>"
'                response.write "<td width=""20""></td>"
'                response.write "<td name=""tdLabelInpVisitDate"" id=""tdLabelInpVisitDate"" style=""font-weight: bold"">Discharge Date</td>"
'                response.write "<td width=""20""></td>"
'                response.write "<td name=""tdInputInpVisitDate"" id=""tdInputInpVisitDate"">" & dDt & "</td>"
'                response.write "<td width=""20""></td>"
'                response.write "</tr>"

                .MoveNext
            Loop

        End If

        .Close
    End With
    
    response.write "<tr>"
    response.write "<td name=""tdLabelInpInsuranceNo"" id=""tdLabelInpInsuranceNo"" style=""font-weight: bold"">" & CStr(cnt) & ".  Admission Date</td>"
    response.write "<td width=""20""></td>"
    response.write "<td name=""tdInputInpInsuranceNo"" id=""tdInputInpInsuranceNo"">" & FormatDate(aDt) & "</td>"
    response.write "<td width=""20""></td>"
    response.write "<td name=""tdLabelInpVisitDate"" id=""tdLabelInpVisitDate"" style=""font-weight: bold"">Discharge Date</td>"
    response.write "<td width=""20""></td>"
    If onAdm Then
        response.write "<td name=""tdInputInpVisitDate"" id=""tdInputInpVisitDate"">-</td>"
    Else
        response.write "<td name=""tdInputInpVisitDate"" id=""tdInputInpVisitDate"">" & FormatDate(dDt) & "</td>"
    End If
    response.write "<td width=""20""></td>"
    response.write "</tr>"

    Set rst = Nothing
End Sub

response.write "<tr>"
response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table id=""tblHiddenFields"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table id=""tblFooter"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
response.write "<tr>"
response.write "<td colspan=""7"" bgcolor=""#FFFFFF"" height=""10"" style=""font-size: 8pt"" align=""right"">"
response.write "" & GetComboName("Branch", "B001") & " @" & Year(Now()) & "</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
