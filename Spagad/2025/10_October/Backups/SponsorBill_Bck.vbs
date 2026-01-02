'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
'GetDrug
Function GetDrug(vst)
    Dim rst, sql, ot, cnt
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select sum(finalamt) as sm from drugsaleitems where visitationid='" & vst & "'"
    ot = 0

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst

            If Not IsNull(.fields("sm")) Then
                ot = .fields("sm")
            End If

        End If

        .Close
    End With

    Set rst = Nothing
    GetDrug = ot
End Function

'GetLab
Function GetLab(vst)
    Dim rst, sql, ot, cnt
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select sum(finalamt) as sm from investigation where visitationid='" & vst & "'"
    ot = 0

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst

            If Not IsNull(.fields("sm")) Then
                ot = .fields("sm")
            End If

        End If

        .Close
    End With

    Set rst = Nothing
    GetLab = ot
End Function

'GetXRay
Function GetXRay(vst)
    Dim rst, sql, ot, cnt
    Set rst = server.CreateObject("ADODB.Recordset")
    sql = "select sum(finalamt) as sm from investigation where visitationid='" & vst & "'"
    sql = sql & " and (testcategoryid='T006' or testcategoryid='T007' or testcategoryid='T008')"
    ot = 0

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst

            If Not IsNull(.fields("sm")) Then
                ot = .fields("sm")
            End If

        End If

        .Close
    End With

    Set rst = Nothing
    GetXRay = ot
End Function

'GetReturnAmt
Function GetReturnAmt(vst)
    Dim rstTblSql, sql, ot
    Set rstTblSql = CreateObject("ADODB.Recordset")
    ot = 0

    With rstTblSql
        'sql = "select sum(finalamt) as sm from drugreturnitems where visitationid='" & vst & "'"

        sql = "select sum(finalamt) as sm from ( "
        sql = sql & "select FinalAmt, returnqty from drugreturnitems where visitationid='" & vst & "' "
        sql = sql & "union all select FinalAmt, returnqty from drugreturnitems2 where visitationid='" & vst & "' "
        sql = sql & ") as t"

        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
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

'GetDept
Function GetDept(dpt)
    Dim ot, arr, num, ul, pt, fnd, pos, cnt
    ot = ""
    pos = -1
    fnd = False

    If Not IsNull(dpt) Then
        arr = Split(dpt, "/")
        ul = UBound(arr)

        For num = 0 To ul
            pt = Trim(arr(num))

            If Len(pt) > 0 Then

                If IsNumeric(Right(pt, 1)) Then
                    pos = num
                    fnd = True
                    Exit For
                End If

            End If

        Next

        cnt = 0

        For num = 0 To ul

            If num <> pos Then
                cnt = cnt + 1

                If cnt > 1 Then
                    ot = ot & "/"
                End If

                ot = ot & arr(num)
            End If

        Next

    End If

    GetDept = ot
End Function

Sub SetClaimItems(vst)
    Dim ot, rst, sql, ky, pos, tot, cnt, unicost, gTot, rAmt
    Dim rQty, nQty, aDt, dy, tot2, sbCst
    Set rst = CreateObject("ADODB.Recordset")
    ot = 0
    tot = 0
    pos = 0
    gTot = 0
    unicost = 0
    'Consultation
    clmCon = GetComboNameFld("Visitation", vst, "VisitCost")
    'LABTEST

    With rst
        'sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from investigation where visitationid='" & vst & "' and TestCategoryID<>'B19'"

        sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from ("
        sql = sql & "select labtestid, qty, unitcost, finalamt from investigation where visitationid='" & vst & "' and TestCategoryID<>'B19' "
        sql = sql & "union all "
        sql = sql & "select labtestid, qty, unitcost, finalamt from investigation2 where visitationid='" & vst & "' and TestCategoryID<>'B19' ) as t"

        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst
            ot = .fields("amt")
            tot = 0

            If IsNumeric(ot) Then
                tot = CDbl(ot)
                clmLab = clmLab + tot
            End If

        End If

        .Close
    End With

    ''LABTEST BY PRESC.
    'With rst
    '  sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from investigation2 where visitationid='" & vst & "'"
    '  .Open qryPro.FltQry(sql), conn, 3, 4
    '  If .RecordCount > 0 Then
    '    .movefirst
    '    ot = .fields("amt")
    '    tot = 0
    '    If IsNumeric(ot) Then
    '      tot = CDbl(ot)
    '      clmLab = clmLab + tot
    '    End If
    '  End If
    '  .Close
    'End With

    'XRAY

    With rst
        'sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from investigation where visitationid='" & vst & "' and TestCategoryID='B19'"

        sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from ("
        sql = sql & "select labtestid, qty, unitcost, finalamt from investigation where visitationid='" & vst & "' and TestCategoryID = 'B19' "
        sql = sql & "union all "
        sql = sql & "select labtestid, qty, unitcost, finalamt from investigation2 where visitationid='" & vst & "' and TestCategoryID = 'B19' ) as t"

        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst
            ot = .fields("amt")
            tot = 0

            If IsNumeric(ot) Then
                tot = CDbl(ot)
                clmXrSc = clmXrSc + tot
            End If

        End If

        .Close
    End With

    ''X-Ray/Scan BY PRESC.
    'With rst
    '  sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from LabByDoctor2 where visitationid='" & vst & "'"
    '  sql = sql & " and (TestCategoryID='T006' or TestCategoryID='T007') and JobScheduleID='FrontDeskCashier'"
    '  .Open qryPro.FltQry(sql), conn, 3, 4
    '  If .RecordCount > 0 Then
    '    .movefirst
    '    ot = .fields("amt")
    '    tot = 0
    '    If IsNumeric(ot) Then
    '      tot = CDbl(ot)
    '      clmXrSc = clmXrSc + tot
    '    End If
    '  End If
    '  .Close
    'End With

    'DRUGS

    With rst
        'sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from drugsaleitems where visitationid='" & vst & "'"

        sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from ("
        sql = sql & "select drugid, qty, unitcost, finalamt from drugsaleitems where visitationid='" & vst & "' "
        sql = sql & " union all "
        sql = sql & "select drugid, DispenseAmt1 as qty, unitcost, dispenseAmt2 as finalamt from drugsaleitems2 where visitationid='" & vst & "') as t "

        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst
            ot = .fields("amt")
            tot = 0

            If IsNumeric(ot) Then
                tot = CDbl(ot)
                clmDrg = clmDrg + tot
            End If

            rAmt = GetReturnAmt(vst)
            clmDrg = clmDrg - rAmt
        End If

        .Close
    End With

    ''DRUGS BY PRESC.
    'With rst
    '  sql = "select sum(DispenseAmt2) as amt, avg(unitcost) as unicost from drugsaleitems2 where visitationid='" & vst & "'"
    '  .Open qryPro.FltQry(sql), conn, 3, 4
    '  If .RecordCount > 0 Then
    '    ot = .fields("amt")
    '    tot = 0
    '    If IsNumeric(ot) Then
    '      tot = CDbl(ot)
    '      clmDrg = clmDrg + tot
    '    End If
    '  End If
    '  .Close
    'End With
    'ADMISSION

    With rst
        sql = "select * from Admission where visitationid='" & vst & "'"
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst

            Do While Not .EOF
                unt = .fields("bedcharge")
                cnt = .fields("noofdays")
                aDt = .fields("AdmissionDate")

                If Not IsNumeric(cnt) Then
                    cnt = DateDiff("d", aDt, Now())
                    dy = DateDiff("h", CDate(aDt), Now())
                    dy = dy / 24
                    cnt = Int(dy)

                    If (dy - Int(dy)) > 0 Then
                        cnt = Int(dy) + 1
                    End If

                End If

                If cnt > 0 Then
                    tot = unt * cnt
                    clmAdm = 0 ' clmAdm + tot
                End If

                .MoveNext
            Loop

        End If

        .Close
    End With

    'NON-DRUG CONSUMABLES

    With rst
        sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from stockIssueItems where visitationid='" & vst & "'"
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst
            ot = .fields("amt")
            tot = 0

            If IsNumeric(ot) Then
                tot = CDbl(ot)
                clmOth = clmOth + tot
            End If

        End If

        .Close
    End With

    'EYE

    With rst
        sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from treatcharges where visitationid='" & vst & "' and  treatModeid='B80'"
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst
            ot = .fields("amt")
            tot = 0

            If IsNumeric(ot) Then
                tot = CDbl(ot)
                clmEye = clmEye + tot

            End If

        End If

        .Close
    End With

    'ENT

    With rst
        sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from treatcharges where visitationid='" & vst & "' and  treatModeid='B90'"
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst
            ot = .fields("amt")
            tot = 0

            If IsNumeric(ot) Then
                tot = CDbl(ot)
                clmEnt = clmEnt + tot
            End If

        End If

        .Close
    End With

    'DENT

    With rst
        sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from treatcharges where visitationid='" & vst & "' and  treatModeid='B88'"
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst
            ot = .fields("amt")
            tot = 0

            If IsNumeric(ot) Then
                tot = CDbl(ot)
                clmDent = clmDent + tot
            End If

        End If

        .Close
    End With

    'PHYS

    With rst
        sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from treatcharges where visitationid='" & vst & "' and  treatModeid='B91'"
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst
            ot = .fields("amt")
            tot = 0

            If IsNumeric(ot) Then
                tot = CDbl(ot)
                clmPhys = clmPhys + tot
            End If

        End If

        .Close
    End With

    'PROCEDURE

    With rst
        sql = "select sum(finalamt) as amt, avg(unitcost) as unicost from treatcharges where visitationid='" & vst & "'"
        sql = sql & " and  treatModeid<>'B80' and  treatModeid<>'B88' and  treatModeid<>'B90' and  treatModeid<>'B91'"
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst
            ot = .fields("amt")
            tot = 0

            If IsNumeric(ot) Then
                tot = CDbl(ot)
                clmPro = clmPro + tot
            End If

        End If

        .Close
    End With

    clmAdmPro = clmAdm + clmPro

    clmTot = clmTot + clmCon
    clmTot = clmTot + clmDrg
    clmTot = clmTot + clmLab
    clmTot = clmTot + clmAdm
    clmTot = clmTot + clmXrSc
    clmTot = clmTot + clmOth
    clmTot = clmTot + clmPro
    clmTot = clmTot + clmEye
    clmTot = clmTot + clmDent
    clmTot = clmTot + clmEnt
    clmTot = clmTot + clmPhys

    Set rst = Nothing
End Sub

'SetInsPatientInfo
Sub SetInsPatientInfo(inspat)
    Dim rst, sql, ot, cnt, iDep, pPat
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select * from insuredpatient where insuredpatientid='" & inspat & "'"
    ot = 0

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst
            iDep = Trim(.fields("InitialDependantID"))
            gInsNo = Trim(.fields("InsuranceNo"))

            If UCase(iDep) <> "NONE" Then
                pPat = Trim(GetComboNameFld("InsuredPatient", iDep, "PatientID"))
                spDet = GetComboName("Patient", pPat)
            End If

        End If

        .Close
    End With

    Set rst = Nothing
End Sub

'DisplayHeaderFooter
Function DisplayHeaderFooter(modCnt, modPos, pgCnt)
    Dim ot
    ot = pgCnt

    If (modPos Mod modCnt) = 0 Then

        If ot > 0 Then
            response.write "<tr><td></td>"
            response.write "<td colspan=""17""><b>PAGE NO.&nbsp;:&nbsp;" & CStr(ot) & "</b></td>"
            response.write "</tr>"
            response.write "</table>"
            modCnt = 35
            modPos = 1
        End If

        response.write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""3"" style=""color:#000000;font-size: 8pt; font-family: Arial; border-collapse:collapse; page-break-after:always"">"

        response.write "<tr height=""1"">"
        response.write "<td colspan=""18""></td>"
        response.write "</tr>"

        response.write "<tr><td valign=""top""><u><b>NO.</b></u></td>"
        response.write "<td valign=""top""><u><b>DATE</b></u></td>"
        response.write "<td valign=""top""><u><b>HOSP#</b></u></td>"
        response.write "<td valign=""top""><u><b>NAME</b></u></td>"
        response.write "<td valign=""top""><u><b>SPONSOR DETAIL</b></u></td>"
        response.write "<td valign=""top""><u><b>CLAIM&nbsp;NO.</b></u></td>"
        response.write "<td valign=""top""><u><b>INS.&nbsp;NO.</b></u></td>"
        response.write "<td valign=""top"" align=""right""><u><b>OPD</b></u></td>"
        response.write "<td valign=""top"" align=""right""><u><b>EYE</b></u></td>"
        response.write "<td valign=""top"" align=""right""><u><b>DENT</b></u></td>"
        response.write "<td valign=""top"" align=""right""><u><b>ENT</b></u></td>"
        response.write "<td valign=""top"" align=""right""><u><b>PHYS</b></u></td>"
        response.write "<td valign=""top"" align=""right""><u><b>PHARM</b></u></td>"
        'response.write "<td valign=""top"" align=""right""><u><b>WARD</b></u></td>"
        response.write "<td valign=""top"" align=""right""><u><b>LAB</b></u></td>"
        response.write "<td valign=""top"" align=""right""><u><b>XRAY</b></u></td>"
        response.write "<td valign=""top"" align=""right""><u><b>WARD/THEATR</b></u></td>"
        response.write "<td valign=""top"" align=""right""><u><b>OTHERS</b></u></td>"
        response.write "<td valign=""top"" align=""right""><u><b>TOTALS</b></u></td></tr>"

        ot = ot + 1
    End If

    DisplayHeaderFooter = ot
End Function

'DisplayCompanyReport
Sub DisplayCompanyReport()
    Dim rst, sql, wkd, sp, rst1, tot, gSt, cnt, patNm
    Dim insr, aCnt, vst, vty, ag, insno, comTot, pNo
    Dim con, drg, lab, xr, vdt, dpt, cDpt, dptCnt

    Dim totCon, totDrg, totLab, totXrSc, totAdm, totPro, totOth, totTot
    Dim totEye, totEnt, totDent, totPhys, totAdmPro

    Dim dptCon, dptDrg, dptLab, dptXrSc, dptAdm, dptPro, dptOth, dptTot
    Dim dptEye, dptEnt, dptDent, dptPhys, dptAdmPro
    Dim hrf

    pgCnt = 0
    modCnt = 30
    modPos = -1
    gCoPayTot = 0

    aCnt = 0
    cnt = 0
    dptCnt = 0

    insr = Trim(GetRecordField("Sponsorid"))
    Set rst = CreateObject("ADODB.Recordset")
    Set rst1 = CreateObject("ADODB.Recordset")

    sql = "select * from visitation where billmonthid='" & mth & "'"
    sql = sql & " and Sponsorid='" & insr & "' order by insschememodeid,visitdate"

    totCon = 0
    totDrg = 0
    totLab = 0
    totAdm = 0
    totXrSc = 0
    totOth = 0
    totPro = 0
    totAdmPro = 0
    totTot = 0

    totEye = 0
    totEnt = 0
    totDent = 0
    totPhys = 0

    'Dept
    dptCon = 0
    dptDrg = 0
    dptLab = 0
    dptAdm = 0
    dptXrSc = 0
    dptOth = 0
    dptPro = 0
    dptAdmPro = 0
    dptTot = 0

    dptEye = 0
    dptEnt = 0
    dptDent = 0
    dptPhys = 0

    dpt = ""
    cDpt = ""

    With rst1
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst

            Do While Not .EOF
                'Header/Footer
                cnt = cnt + 1
                aCnt = aCnt + 1

                pNo = .fields("PatientID")
                vst = .fields("visitationid")
                vdt = .fields("visitdate")
                dpt = .fields("InsSchemeModeID")

                patNm = GetComboName("Patient", pNo) '.fields("VisitationName")
                insno = .fields("InsuranceNo")
                spDet = ""
                SetInsPatientInfo .fields("InsuredPatientID")

                'Dept

                If UCase(cDpt) <> UCase(dpt) Then
                    dptCnt = dptCnt + 1

                    If dptCnt > 1 Then
                        modPos = modPos + 1
                        pgCnt = DisplayHeaderFooter(modCnt, modPos, pgCnt)

                        response.write "<tr><td></td>"
                        response.write "<td colspan=""6""><b>Sub Total</b></td>"
                        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptCon), 2, , , -1)) & "</b></td>"
                        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptEye), 2, , , -1)) & "</b></td>"
                        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptDent), 2, , , -1)) & "</b></td>"
                        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptEnt), 2, , , -1)) & "</b></td>"
                        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptPhys), 2, , , -1)) & "</b></td>"
                        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptDrg), 2, , , -1)) & "</b></td>"
                        'response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptAdm), 2, , , -1)) & "</b></td>"
                        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptLab), 2, , , -1)) & "</b></td>"
                        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptXrSc), 2, , , -1)) & "</b></td>"
                        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptAdmPro), 2, , , -1)) & "</b></td>"
                        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptOth), 2, , , -1)) & "</b></td>"
                        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptTot), 2, , , -1)) & "</b></td>"
                        response.write "</tr>"
                    End If

                    If ((modPos + 2) Mod modCnt) = 0 Then
                        modPos = modPos + 2 'Jump one line if Dept is about to start
                        pgCnt = DisplayHeaderFooter(modCnt, modPos, pgCnt)
                    Else
                        modPos = modPos + 1
                        pgCnt = DisplayHeaderFooter(modCnt, modPos, pgCnt)
                    End If

                    response.write "<tr><td></td>"
                    response.write "<td colspan=""17""><b>" & UCase(GetComboName("InsSchemeMode", dpt)) & "</b></td>"
                    response.write "</tr>"
                    cDpt = dpt

                    dptCon = 0
                    dptDrg = 0
                    dptLab = 0
                    dptAdm = 0
                    dptXrSc = 0
                    dptOth = 0
                    dptPro = 0
                    dptTot = 0
                    dptAdmPro = 0

                    dptEye = 0
                    dptEnt = 0
                    dptDent = 0
                    dptPhys = 0
                End If

                modPos = modPos + 1
                pgCnt = DisplayHeaderFooter(modCnt, modPos, pgCnt)

                response.write "<tr><td valign=""top"">" & CStr(aCnt) & "</td>"
                response.write "<td valign=""top"">" & Day(vdt) & "/" & Month(vdt) & "/" & Right(Year(vdt), 2) & "</td>"
                response.write "<td valign=""top"">" & UCase(pNo) & "</td>"

                If Len(patNm) > 25 Then
                    response.write "<td style=""font-size:7pt"" valign=""top"">"
                Else
                    response.write "<td valign=""top"">"
                End If

                response.write Replace(UCase(patNm), " ", "&nbsp;") & "</td>"
                'response.write "<td valign=""top"">" & UCase(patNm) & "</td>"

                If Len(spDet) > 25 Then
                    response.write "<td style=""font-size:7pt"" valign=""top"">"
                Else
                    response.write "<td valign=""top"">"
                End If

                response.write Replace(UCase(spDet), " ", "&nbsp;") & "</td>"
                'response.write "<td valign=""top"">" & UCase(spDet) & "</td>"

                hrf = "wpgPrtPrintLayoutAll.asp?PositionForTableName=Visitation&PrintLayoutName=VisitationClaim&VisitationID=" & vst
                response.write "<td valign=""top""><a target=""_Blank"" href=""" & hrf & """>" & UCase(vst) & "</a></td>"

                response.write "<td valign=""top"">" & Replace(UCase(gInsNo), " ", "&nbsp;") & "</td>"

                clmCon = 0
                clmDrg = 0
                clmLab = 0
                clmAdm = 0
                clmXrSc = 0
                clmOth = 0
                clmPro = 0
                clmTot = 0
                clmAdmPro = 0

                clmEye = 0
                clmEnt = 0
                clmDent = 0
                clmPhys = 0

                SetClaimItems vst

                clmAdmPro = clmAdm + clmPro

                totCon = totCon + clmCon
                totDrg = totDrg + clmDrg
                totLab = totLab + clmLab
                totAdm = totAdm + clmAdm
                totXrSc = totXrSc + clmXrSc
                totOth = totOth + clmOth
                totPro = totPro + clmPro
                totTot = totTot + clmTot
                totAdmPro = totAdmPro + clmAdm + clmPro

                totEye = totEye + clmEye
                totEnt = totEnt + clmEnt
                totDent = totDent + clmDent
                totPhys = totPhys + clmPhys

                'Depart
                dptCon = dptCon + clmCon
                dptDrg = dptDrg + clmDrg
                dptLab = dptLab + clmLab
                dptAdm = dptAdm + clmAdm
                dptXrSc = dptXrSc + clmXrSc
                dptOth = dptOth + clmOth
                dptPro = dptPro + clmPro
                dptTot = dptTot + clmTot
                dptAdmPro = dptAdmPro + clmAdm + clmPro

                dptEye = dptEye + clmEye
                dptEnt = dptEnt + clmEnt
                dptDent = dptDent + clmDent
                dptPhys = dptPhys + clmPhys

                If clmCon <> 0 Then
                    response.write "<td valign=""top"" align=""right"">" & UCase(FormatNumber(CStr(clmCon), 2, , , -1)) & "</td>"
                Else
                    response.write "<td valign=""top"" align=""right"">-</td>"
                End If

                'EYE

                If clmEye <> 0 Then
                    response.write "<td valign=""top"" align=""right"">" & UCase(FormatNumber(CStr(clmEye), 2, , , -1)) & "</td>"
                Else
                    response.write "<td valign=""top"" align=""right"">-</td>"
                End If

                'DENT

                If clmDent <> 0 Then
                    response.write "<td valign=""top"" align=""right"">" & UCase(FormatNumber(CStr(clmDent), 2, , , -1)) & "</td>"
                Else
                    response.write "<td valign=""top"" align=""right"">-</td>"
                End If

                'ENT

                If clmEnt <> 0 Then
                    response.write "<td valign=""top"" align=""right"">" & UCase(FormatNumber(CStr(clmEnt), 2, , , -1)) & "</td>"
                Else
                    response.write "<td valign=""top"" align=""right"">-</td>"
                End If

                'PHYS

                If clmPhys <> 0 Then
                    response.write "<td valign=""top"" align=""right"">" & UCase(FormatNumber(CStr(clmPhys), 2, , , -1)) & "</td>"
                Else
                    response.write "<td valign=""top"" align=""right"">-</td>"
                End If

                If clmDrg <> 0 Then
                    response.write "<td valign=""top"" align=""right"">" & UCase(FormatNumber(CStr(clmDrg), 2, , , -1)) & "</td>"
                Else
                    response.write "<td valign=""top"" align=""right"">-</td>"
                End If

                'If clmAdm <> 0 Then
                'response.write "<td valign=""top"" align=""right"">" & UCase(FormatNumber(CStr(clmAdm), 2, , , -1)) & "</td>"
                'Else
                'response.write "<td valign=""top"" align=""right"">-</td>"
                'End If

                If clmLab <> 0 Then
                    response.write "<td valign=""top"" align=""right"">" & UCase(FormatNumber(CStr(clmLab), 2, , , -1)) & "</td>"
                Else
                    response.write "<td valign=""top"" align=""right"">-</td>"
                End If

                If clmXrSc <> 0 Then
                    response.write "<td valign=""top"" align=""right"">" & UCase(FormatNumber(CStr(clmXrSc), 2, , , -1)) & "</td>"
                Else
                    response.write "<td valign=""top"" align=""right"">-</td>"
                End If

                If clmAdmPro <> 0 Then
                    response.write "<td valign=""top"" align=""right"">" & UCase(FormatNumber(CStr(clmAdmPro), 2, , , -1)) & "</td>"
                Else
                    response.write "<td valign=""top"" align=""right"">-</td>"
                End If

                If clmOth <> 0 Then
                    response.write "<td valign=""top"" align=""right"">" & UCase(FormatNumber(CStr(clmOth), 2, , , -1)) & "</td>"
                Else
                    response.write "<td valign=""top"" align=""right"">-</td>"
                End If

                response.write "<td valign=""top"" align=""right""><b>" & UCase(FormatNumber(CStr(clmTot), 2, , , -1)) & "</b></td>"

                response.write "</tr>"
                .MoveNext
            Loop

        End If

        .Close
    End With

    If dptCnt > 1 Then
        modPos = modPos + 1
        pgCnt = DisplayHeaderFooter(modCnt, modPos, pgCnt)

        response.write "<tr><td></td>"
        response.write "<td colspan=""6""><b>Sub Total</b></td>"
        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptCon), 2, , , -1)) & "</b></td>"
        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptEye), 2, , , -1)) & "</b></td>"
        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptDent), 2, , , -1)) & "</b></td>"
        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptEnt), 2, , , -1)) & "</b></td>"
        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptPhys), 2, , , -1)) & "</b></td>"
        'response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptDrg), 2, , , -1)) & "</b></td>"
        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptAdm), 2, , , -1)) & "</b></td>"
        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptLab), 2, , , -1)) & "</b></td>"
        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptXrSc), 2, , , -1)) & "</b></td>"
        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptAdmPro), 2, , , -1)) & "</b></td>"
        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptOth), 2, , , -1)) & "</b></td>"
        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(dptTot), 2, , , -1)) & "</b></td>"
        response.write "</tr>"
    End If

    modPos = modPos + 1
    pgCnt = DisplayHeaderFooter(modCnt, modPos, pgCnt)

    response.write "<tr><td></td>"
    response.write "<td colspan=""6""><b>GRAND TOTALS</b></td>"
    response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totCon), 2, , , -1)) & "</b></td>"
    response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totEye), 2, , , -1)) & "</b></td>"
    response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totDent), 2, , , -1)) & "</b></td>"
    response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totEnt), 2, , , -1)) & "</b></td>"
    response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totPhys), 2, , , -1)) & "</b></td>"
    response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totDrg), 2, , , -1)) & "</b></td>"
    'response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totAdm), 2, , , -1)) & "</b></td>"
    response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totLab), 2, , , -1)) & "</b></td>"
    response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totXrSc), 2, , , -1)) & "</b></td>"
    response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totAdmPro), 2, , , -1)) & "</b></td>"
    response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totOth), 2, , , -1)) & "</b></td>"
    response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totTot), 2, , , -1)) & "</b></td>"
    response.write "</tr>"

    response.write "<tr>"
    response.write "<td colspan=""19""><hr size=""1""></td>"
    response.write "</tr>"

    DisplayCompanyCoPay

    If gCoPayTot > 0 Then
        response.write "<tr><td></td>"
        response.write "<td colspan=""16""><b>FINAL TOTALS</b></td>"
        response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(totTot - gCoPayTot), 2, , , -1)) & "</b></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td colspan=""18""><hr size=""1""></td>"
        response.write "</tr>"
    End If

    'Close Last Page
    response.write "<tr><td></td>"
    response.write "<td colspan=""18""><b>PAGE NO.&nbsp;:&nbsp;" & CStr(pgCnt) & "</b></td>"
    response.write "</tr>"
    response.write "</table>"

    Set rst = Nothing
    Set rst1 = Nothing
End Sub

'DisplayCompanyCoPay
Sub DisplayCompanyCoPay()
    Dim rst, sql, wkd, sp, rst1, tot, gSt, cnt, patNm, pNo
    Dim insr, aCnt, vst, vty, ag, insno, comTot
    Dim con, drg, lab, xr, vdt, dpt, cDpt, dptCnt, coPay
    Dim totCon, totDrg, totLab, totXrSc, totAdm, totPro, totOth, totTot
    Dim hrf
    Dim dptCon, dptDrg, dptLab, dptXrSc, dptAdm, dptPro, dptOth, dptTot

    'pgCnt = 0
    'modCnt = 45
    'modPos = -1
    gCoPayTot = 0

    aCnt = 0
    cnt = 0
    dptCnt = 0

    insr = Trim(GetRecordField("Sponsorid"))
    Set rst = CreateObject("ADODB.Recordset")
    Set rst1 = CreateObject("ADODB.Recordset")

    sql = "select * from visitation where billmonthid='" & mth & "'"
    sql = sql & " and Sponsorid='" & insr & "' and VisitValue1>0 order by visitdate"

    totCon = 0
    totDrg = 0
    totLab = 0
    totAdm = 0
    totXrSc = 0
    totOth = 0
    totPro = 0
    totTot = 0

    dptCon = 0
    dptDrg = 0
    dptLab = 0
    dptAdm = 0
    dptXrSc = 0
    dptOth = 0
    dptPro = 0
    dptTot = 0

    dpt = ""
    cDpt = ""

    With rst1
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst

            Do While Not .EOF

                coPay = 0

                If Not IsNull(.fields("VisitValue1")) Then

                    If IsNumeric(.fields("VisitValue1")) Then
                        coPay = CDbl(.fields("VisitValue1"))
                    End If

                End If

                If coPay > 0 Then
                    cnt = cnt + 1
                    aCnt = aCnt + 1

                    If cnt = 1 Then

                        response.write "<tr><td valign=""top""><u><b>NO.</b></u></td>"
                        response.write "<td valign=""top""><u><b>DATE</b></u></td>"
                        response.write "<td valign=""top""><u><b>HOSP#</b></u></td>"
                        response.write "<td valign=""top""><u><b>NAME</b></u></td>"
                        response.write "<td valign=""top""><u><b>CLAIM&nbsp;NO.</b></u></td>"
                        response.write "<td valign=""top""><u><b>INS.&nbsp;NO.</b></u></td>"

                        response.write "<td valign=""top"" colspan=""9"">&nbsp;</td>"

                        response.write "<td><u><b>INITIAL.&nbsp;AMT</b></u></td>"
                        response.write "<td><u><b>FINAL.&nbsp;AMT</b></u></td>"
                        response.write "<td valign=""top"" align=""right""><u><b>CO-PAY</b></u></td>"
                        response.write "</tr>"
                    End If

                    pNo = .fields("PatientID")
                    vst = .fields("visitationid")
                    vdt = .fields("visitdate")
                    dpt = .fields("InsSchemeModeID")

                    patNm = GetComboName("Patient", pNo) '.fields("VisitationName")
                    insno = .fields("InsuranceNo")
                    spDet = ""
                    SetInsPatientInfo .fields("InsuredPatientID")

                    response.write "<tr><td valign=""top"">" & CStr(aCnt) & "</td>"
                    response.write "<td valign=""top"">" & Day(vdt) & "/" & Month(vdt) & "/" & Right(Year(vdt), 2) & "</td>"
                    response.write "<td valign=""top"">" & UCase(pNo) & "</td>"

                    If Len(patNm) > 25 Then
                        response.write "<td valign=""top"" style=""font-size:7pt"" valign=""top"">"
                    Else
                        response.write "<td valign=""top"">"
                    End If

                    response.write Replace(UCase(patNm), " ", "&nbsp;") & "</td>"

                    hrf = "wpgPrtPrintLayoutAll.asp?PositionForTableName=Visitation&PrintLayoutName=VisitationClaim&VisitationID=" & vst
                    response.write "<td valign=""top""><a target=""_Blank"" href=""" & hrf & """>" & UCase(vst) & "</a></td>"

                    response.write "<td valign=""top"">" & Replace(UCase(gInsNo), " ", "&nbsp;") & "</td>"

                    clmCon = 0
                    clmDrg = 0
                    clmLab = 0
                    clmAdm = 0
                    clmXrSc = 0
                    clmOth = 0
                    clmPro = 0
                    clmTot = 0

                    clmCon = 0
                    clmDrg = 0
                    clmLab = 0
                    clmAdm = 0
                    clmXrSc = 0
                    clmOth = 0
                    clmPro = 0
                    clmEye = 0
                    clmDent = 0
                    clmEnt = 0
                    clmPhys = 0

                    SetClaimItems vst

                    totCon = totCon + clmCon
                    totDrg = totDrg + clmDrg
                    totLab = totLab + clmLab
                    totAdm = totAdm + clmAdm
                    totXrSc = totXrSc + clmXrSc
                    totOth = totOth + clmOth
                    totPro = totPro + clmPro
                    totTot = totTot + clmTot

                    gCoPayTot = gCoPayTot + coPay

                    dptCon = dptCon + clmCon
                    dptDrg = dptDrg + clmDrg
                    dptLab = dptLab + clmLab
                    dptAdm = dptAdm + clmAdm
                    dptXrSc = dptXrSc + clmXrSc
                    dptOth = dptOth + clmOth
                    dptPro = dptPro + clmPro
                    dptTot = dptTot + clmTot

                    response.write "<td colspan=""9"">&nbsp;</td>"

                    response.write "<td>" & UCase(FormatNumber(CStr(clmTot), 2, , , -1)) & "</td>"
                    response.write "<td>" & UCase(FormatNumber(CStr(clmTot - coPay), 2, , , -1)) & "</td>"
                    response.write "<td valign=""top"" align=""right""><b>" & UCase(FormatNumber(CStr(coPay), 2, , , -1)) & "</b></td>"

                    response.write "</tr>"
                End If 'coPay

                .MoveNext
            Loop

            If cnt > 0 Then
                response.write "<tr><td></td>"
                response.write "<td colspan=""16""><b>CO-PAYMENT TOTALS</b></td>"
                response.write "<td align=""right""><b>" & UCase(FormatNumber(CStr(gCoPayTot), 2, , , -1)) & "</b></td>"
                response.write "</tr>"

                response.write "<tr>"
                response.write "<td colspan=""18""><hr size=""1""></td>"
                response.write "</tr>"
            End If

        End If

        .Close
    End With

    Set rst = Nothing
    Set rst1 = Nothing
End Sub

'GetCompanyCoPay
Function GetCompanyCoPay()
    Dim rst, sql, wkd, sp, rst1, tot, gSt, cnt, patNm, pNo
    Dim insr, aCnt, vst, vty, ag, insno, comTot
    Dim con, drg, lab, xr, vdt, dpt, cDpt, dptCnt, coPay

    insr = Trim(GetRecordField("Sponsorid"))
    Set rst = CreateObject("ADODB.Recordset")
    Set rst1 = CreateObject("ADODB.Recordset")

    coPay = 0

    sql = "select sum(visitValue1) as amt from visitation where billmonthid='" & mth & "'"
    sql = sql & " and Sponsorid='" & insr & "' and VisitValue1>0"

    With rst1
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst

            If Not IsNull(.fields("amt")) Then

                If IsNumeric(.fields("amt")) Then
                    coPay = CDbl(.fields("amt"))
                End If

            End If

        End If

        .Close
    End With

    GetCompanyCoPay = coPay
    Set rst = Nothing
    Set rst1 = Nothing
End Function

Sub DisplayCoverLetterL()

    Dim wkDy, prtFlt, totAmt, totClaim
    Dim cellSty, cellStyB, coPayAmt, clmAmt
    coPayAmt = GetCompanyCoPay()
    clmAmt = GetSponsorBillAmt(GetRecordField("SponsorID"))
    totAmt = FormatNumber(CStr(clmAmt - coPayAmt), 2, , , -1)
    totClaim = FormatNumber(CStr(GetClaimCount(GetRecordField("SponsorID"))), 0, , , -1)

    cellSty = "border-top: 1px solid #808080; border-left: 1px solid #808080; border-right: 1px solid #808080; border-bottom: 1px solid #808080"
    cellStyB = "border-bottom: 1px solid #888888"

    response.write "<tr>"
    response.write "<td align=""center"">"

    DisplayHeader

    response.write "</td>"
    response.write "</tr>"
    response.write "<tr>"
    response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
    response.write "</tr>"
    'Content
    response.write "<tr>"
    response.write "<td align=""left"">"
    response.write "<table border=""0""  width=""100%"" cellspacing=""0"">"
    response.write "<tr>"
    response.write "<td align=""center"" width=""10%""></td>"
    response.write "<td valign=""top""  align=""left"" width=""90%"">"
    response.write "<table border=""0"" width=""100%"" cellspacing=""3"" cellpadding=""0"" style=""font-size: 13pt; font-family: Arial"">"

    response.write "<tr>"
    response.write "<td>"
    response.write "<br>Our Ref ----------------------------------------&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Your Ref ---------------------------------------<br><br>"
    response.write "</td>"
    response.write "</tr>"

    response.write "<tr>"
    response.write "<td>"
    response.write "INVOICE NO. " & Trim(GetRecordField("Sponsorid")) & Right(mth, 4) & "<br><br>"
    response.write "</td>"
    response.write "</tr>"

    '    response.write "<tr>"
    '      response.write "<td>"
    '      response.write "Your Ref ---------------------------------------<br><br>"
    '      response.write "</td>"
    '    response.write "</tr>"

    response.write "<tr>"
    response.write "<td>"
    response.write FormatDate(Now()) & "<br><br>"
    response.write "</td>"
    response.write "</tr>"

    response.write "<tr>"
    response.write "<td>"
    response.write "<u>SUBMISSION OF MEDICAL BILL</u><br><br>"
    response.write "</td>"
    response.write "</tr>"

    response.write "<tr>"
    response.write "<td>"
    response.write "<b>" & UCase(GetRecordField("SponsorName")) & "</b><br><br>"
    response.write "</td>"
    response.write "</tr>"

    response.write "<tr>"
    response.write "<td style=""font-size:12pt"">"
    response.write "We submit herewith Medical Bill for the Period <b>" & GetWorkingMonthName(mth) & ".</b><br><br>"
    response.write "Total amount for bill is <b>GHC&nbsp;" & totAmt & ".</b> Total number of Claims are&nbsp;<b>" & totClaim & ".</b><br><br>"
    response.write "Please find attached the details for your perusal. Counting on your cooperation.<br><br>"
    response.write "Thank you.<br><br>"
    response.write "Yours faithfully,<br><img src=""images/AccountManagerSign2.jpg""><br>"
    response.write "----------------------------------------<br>"
    response.write "ACCOUNTS MANAGER<br><br>"
    response.write "NOTE:<br>"
    response.write "</td>"
    response.write "</tr>"

    response.write "<tr>"
    response.write "<td style=""font-size:9pt"">"
    response.write "1. Bills should be paid within one month after date of submission.<br>"
    response.write "2. Please treat all Medical information and bills as extremely confidential. Take all necessary steps to guarantee the confidentiality of all information.<br><br>"
    response.write "Thank You.<br>"
    response.write "</td>"
    response.write "</tr>"

    response.write "<tr>"
    response.write "<td style=""font-size:11pt"" align=""center"">"
    response.write "CONTACTS"
    response.write "</td>"
    response.write "</tr>"
    response.write "</table>"

    response.write "</td>"
    response.write "</tr>"

    'Bottom
    response.write "<tr>"
    response.write "<td colspan=""2"" valign=""top"" align=""left"" width=""100%"" style=""" & cellSty & """>"
    response.write "<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""0"" style=""font-size: 7pt; font-family: Arial"">"

    response.write "<td align=""center"">"
    response.write "Hospital&nbsp;HOSPITAL&nbsp;-&nbsp;OSU <br>"
    response.write "35 CANTONMENT RD <br>"
    response.write "(OSU OXFORD STREET)<br>"
    response.write "Opposite Goil Filling Station <br>"
    response.write "OSU, ACCRA<br>"
    response.write "TEL:0302761976-7<br><br>"
    response.write "</td>"

    'response.write "<td>&nbsp;</td>"

    response.write "<td align=""center"">"
    response.write "Hospital SPECIALIST HOSPITAL<br>"
    response.write "OSU, ACCRA<br>"
    response.write "TEL:0302797147<br><br>"
    response.write "</td>"

    'response.write "<td>&nbsp;</td>"

    response.write "<td align=""center"">"
    response.write "Hospital MOTHER & CHILD HOSPITAL<br>"
    response.write "OSU, ACCRA <br>"
    response.write "TEL: 0302798290, 0231797953<br><br>"
    response.write "</td>"

    'response.write "<td>&nbsp;</td>"

    response.write "<td align=""center"">"
    response.write "Hospital CLINIC - LEGON<br>"
    response.write "LEGON SHOPPING CENTRE <br>"
    response.write "LEGON HOUSING<br>"
    response.write "LEGON, ACCRA<br>"
    response.write "TEL: 0236840627<br><br>"
    response.write "</td>"

    'response.write "<td>&nbsp;</td>"

    response.write "<td align=""center"">"
    response.write "Hospital CLINIC - LEGON<br>"
    response.write "LEGON, ACCRA<br>"
    response.write "TEL:0302761976-7<br><br>"
    response.write "</td>"

    'response.write "<td>&nbsp;</td>"

    response.write "<td align=""center"">"
    response.write "Hospital CLINIC - DOME<br>"
    response.write "DOME CFC<br>"
    response.write "DOME, ACCRA<br>"
    response.write "(Near Dome Vodafone Station)<br>"
    response.write "TEL:<br><br>"
    response.write "</td>"

    'response.write "<td>&nbsp;</td>"

    response.write "<td align=""center"">"
    response.write "Hospital CLINIC - PENSION HOUSE <br>"
    response.write "TEL: 0216840942 <br><br>"
    response.write "</td>"

    'response.write "<td>&nbsp;</td>"

    response.write "<td align=""center"">"
    response.write "Hospital CLINIC - LEGON <br>"
    response.write "OPPOSITE TENNIS CLUB <br>"
    response.write "LEGON, ACCRA<br>"
    response.write "TEL: 0303403861-2 <br><br>"
    response.write "</td>"

    'response.write "<td>&nbsp;</td>"

    response.write "<td align=""center"">"
    response.write "Hospital CLINIC - TEMA <br>"
    response.write "COMMUNITY TWO <br>"
    response.write "(NEAR MARKET)<br>"
    response.write "TEMA, ACCRA <br>"
    response.write "TEL: 0303212992 <br><br>"
    response.write "</td>"
    response.write "</tr>"

    response.write "</table>"
    response.write "</td>"
    response.write "</tr>"

    response.write "</table>"
    response.write "</td>"
    response.write "</tr>"
End Sub

Function GetWorkingMonthName(mth)
    Dim ot, ky
    ky = Trim(mth)
    ot = ""

    If Len(ky) = 9 Then

        If (UCase(Left(ky, 3)) = "MTH") And IsNumeric(Right(ky, 6)) Then
            ot = UCase(MonthName(CLng(Right(ky, 2)), False) & " " & Mid(ky, 4, 4))
        End If

    End If

    GetWorkingMonthName = ot
End Function

Function GetClaimCount(sp)
    Dim rst, sql, ot, tot, amt
    Set rst = CreateObject("ADODB.Recordset")
    ot = 0
    tot = 0
    amt = 0

    With rst
        sql = "select count(distinct visitationID) as amt from corporatebill where sponsorid='" & sp & "' and BillMonthid='" & mth & "'"
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst
            amt = .fields("amt")
            tot = 0

            If IsNumeric(amt) Then
                tot = CDbl(amt)
                ot = ot + tot
            End If

        End If

        .Close
    End With

    GetClaimCount = ot
    Set rst = Nothing

End Function

'GetSponsorBillAmt
Function GetSponsorBillAmt(sp)
    Dim rst, sql, ot, tot, amt
    Set rst = CreateObject("ADODB.Recordset")
    ot = 0
    tot = 0
    amt = 0

    With rst
        sql = "select sum(billAmt1) as amt, sum(billamt4) as cAmt from corporatebill where sponsorid='" & sp & "' and BillMonthid='" & mth & "'"
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            .movefirst
            amt = .fields("amt")
            tot = 0

            If IsNumeric(amt) Then
                tot = CDbl(amt)
                ot = ot + tot
            End If

            'Cancel
            amt = .fields("cAmt")
            tot = 0

            If IsNumeric(amt) Then
                tot = CDbl(amt)
                ot = ot - tot
            End If

        End If

        .Close
    End With

    GetSponsorBillAmt = ot
    Set rst = Nothing
End Function

Sub DisplayHeader()
    response.write "<table border=""0""cellspacing=""0"" cellpadding=""0"" width=""" & PrintWidth & """>"
    response.write "<tr><td>"
    response.write "<img src=""images/logo.jpg"" height=""60"" width=""60"">"
    response.write "</td>"
    response.write "<td>"

    response.write "<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">"
    response.write "<tr>"
    response.write "<td align=""center"" style=""font-size: 14pt; font-weight:bold"" colspan=""6""></td>"
    response.write "</tr>"
    response.write "<tr>"
    response.write "<td align=""center"" style=""font-size: 14pt"" colspan=""6"">Hospital</td>"
    response.write "</tr>"
    response.write "<tr>"
    response.write "<td align=""center"" style=""font-size: 10pt"" colspan=""6"">PMB 16, MINISTRY, ACCRA</td>"
    response.write "</tr>"

    response.write "<tr>"
    response.write "<td align=""center"" style=""font-size: 10pt"" colspan=""6"">FAX:&nbsp;+233-0302-777790,&nbsp;&nbsp;&nbsp;&nbsp;WEB:&nbsp;Hospital.com, EMAIL:&nbsp;info@Hospital.com</td>"
    response.write "</tr>"
    response.write "</table>"

    response.write "</td>"
    response.write "</tr>"
    response.write "</table>"
End Sub

Dim mth, wkDy, prtFlt
Dim clmCon, clmDrg, clmLab, clmXrSc, clmAdm, clmPro, clmOth, clmTot, clmAdmPro
Dim clmEye, clmEnt, clmDent, clmPhys, spDet, modCnt, pgCnt, modPos, gInsNo, gCoPayTot

server.scripttimeout = 1800

prtFlt = Request.QueryString("PrintFilter")
mth = prtFlt

'CoverLetter
response.write "<tr>"
response.write "<td align=""left"">"
response.write "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0""style=""page-break-after:always"">"
response.write "<tr><td valign=""top"">"
DisplayCoverLetterL
response.write "</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td align=""center"">"
DisplayHeader
response.write "</td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:10pt"">"
response.write (GetRecordField("SponsorName")) & " [Medical Bills for " & GetComboName("WorkingMonth", mth) & "]</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td align=""left"">"
response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 9pt; font-family: Arial"">"
response.write "<tr>"
response.write "<td align=""left"">"
response.write "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
response.write "<tr><td valign=""top"">"
DisplayCompanyReport
response.write "</td>"
response.write "</tr>"

'response.write "<tr>"
'response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'response.write "</tr>"

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
'None
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
