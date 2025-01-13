'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
addCSS
generateReport

Sub generateReport()
    Dim sql, rst, vst, pat, hrf, subTotal, grandTotal, arrearsTotal, gPaidTotal, InsuranceGroupID
    Dim arPeriod, periodStart, periodEnd, pcst, arrears, recNo, gPaid, gUsed, nsubTotal, ngrandTotal
    Set rst = server.CreateObject("ADODB.Recordset")

    arPeriod = getDatePeriodFromDelim(Trim(Request.QueryString("PrintFilter")))
    periodStart = FormatDate(arPeriod(0))
    periodEnd = FormatDate(arPeriod(1))
    grandTotal = 0
    subTotal = 0
    gPaidTotal = 0
    arrearsTotal = 0
    nsubTotal = 0
    ngrandTotal = 0
    recNo = ""
    
    sql = "SELECT Admission.patientID, Admission.VisitationID, Admission.admissionDate, "
    sql = sql & " Admission.DischargeDate, visitation.visitcost, visitation.InsuranceGroupID"
    sql = sql & " FROM Admission"
    sql = sql & " JOIN Visitation ON Visitation.VisitationID = Admission.VisitationID"
    sql = sql & " WHERE admissionDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    sql = sql & " ORDER BY Admission.admissionDate DESC"
     
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            response.write "<table class = 'anaesthesia' > "
            response.write "    <thead> "
            response.write "    <tr class = 'anaesthesia'>"
            response.write "        <th colspan = '8'>Generated Patients on Admission with Arrears REPORT Between " & periodStart & " and " & periodEnd & "</th>"
            response.write "    </tr>"
            response.write "  <tr class = 'anaesthesia'> "
            response.write "      <th>Patient Name</th> "
            response.write "      <th>Patient ID</th> "
            response.write "      <th>Visit Number</th> "
            response.write "      <th>Admission Date</th> "
            response.write "      <th>Discharge Date</th> "
            response.write "      <th>Total Bill</th> "
            response.write "      <th>Amount Paid</th> "
            response.write "      <th>Arrears</th> "
            response.write "  </tr> "
            response.write "    </thead><tbody> "
            .MoveFirst


                Do While Not .EOF
                    InsuranceGroupID = .fields("InsuranceGroupID")
                    If InsuranceGroupID = "CASH" Then
                        pcst = .fields("visitcost")
                        vst = .fields("VisitationID")
                        pat = .fields("patientID")
                        gPaid = FormatNumber(AddPayments(vst, pat, recNo))
                        gUsed = FormatNumber(AddUsedPayments(vst, pat, recNo))
                        subTotal = AddAdmission(vst) + AddDrug(vst) + AddNonDrug(vst) + AddLab(vst) + AddTreat(vst) + pcst - AddWaiver(vst)
                        grandTotal = FormatNumber(grandTotal + subTotal)
                        arrears = FormatNumber(CDbl(subTotal - gPaid + gUsed - gWv), 2, , , -1)
                        If arrears > 0 Then
                            arrearsTotal = arrearsTotal + arrears
                            gPaidTotal = gPaidTotal + gPaid
                            nsubTotal = subTotal
                            ngrandTotal = grandTotal
                            hrf = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationBill&PositionForTableName=Visitation&VisitationID=" & vst & "&PositionForCtxTableName=Visitation&PageMode=ProcessSelect&AddOtherSection=All"
                            response.write "  <tr class = 'queryData'> "
                            response.write "      <td>" & GetComboName("Patient", .fields("patientID")) & "</td> "
                            response.write "      <td>" & .fields("patientID") & "</td> "
                            response.write "      <td>" & vst & "</td> "
                            response.write "      <td>" & FormatDate(.fields("AdmissionDate")) & "</td> "
                            If Trim(.fields("DischargeDate")) = "" Then
                                response.write "      <td>No Discaharge Date</td> "
                            Else
                            response.write "      <td>" & FormatDate(.fields("DischargeDate")) & "</td> "
                            End If
                            ' response.write "      <td>" & FormatDate(.fields("DischargeDate")) & "</td> "
                            response.write "      <td>" & nsubTotal & "</td> "
                            response.write "      <td>" & gPaid & "</td> "
                            response.write "      <td><a href='" & hrf & "' target='_blank'><b>" & arrears & "</b></a></td> "
                            response.write "  </tr> "
                            subTotal = 0
                            gPaid = 0
                            gUsed = 0
                        End If
                    End If
                    .MoveNext
                Loop
            response.write "    <tr>"
            response.write "        <td colspan='5'><b> GRANDTOTAL </b></td>"
            response.write "        <td><b>" & ngrandTotal & "</b></td>"
            response.write "        <td><b>" & FormatNumber(gPaidTotal, 2, , , -1) & "</b></td>"
            response.write "        <td><b>" & FormatNumber(arrearsTotal, 2, , , -1) & "</b></td>"
            response.write "    </tr>"
            response.write "</tbody></table>"
        End If
        .Close
    End With
    
    Set rst = Nothing
End Sub

Function AddAdmission(vst)
    Dim rst, sql, ot, cnt, hdr, adm, chg, dys, sTot
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select * from admission where visitationid='" & vst & "'"
    cnt = pos
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            sTot = 0
            Do While Not .EOF
                cnt = cnt + 1
                adm = .fields("admissionid")
                chg = 0 '.fields("bedcharge")
                dys = 0 '.fields("noofdays")
                gTot = gTot + (chg * dys)
                sTot = sTot + tot
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    AddAdmission = gTot
End Function

Function AddDrug(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, rQty, drg, fQty, tot2, rTot, sTot
    Set rst = CreateObject("ADODB.Recordset")
    
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
            sTot = 0
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
                  sTot = sTot + tot
                  gTot = gTot + tot
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    pos = cnt
    AddDrug = gTot
End Function

Function AddNonDrug(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, sTot
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select itemid,sum(qty) as qty,avg(retailunitcost) as unt,sum(finalamt) as tot from stockissueitems where visitationid='" & vst & "' group by itemid"
    cnt = pos
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            sTot = 0
            Do While Not .EOF
                cnt = cnt + 1
                unt = .fields("unt")
                qty = .fields("qty")
                tot = .fields("tot")
                gTot = gTot + tot
                sTot = sTot + tot
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    pos = cnt
    AddNonDrug = gTot
End Function

'AddLab
Function AddLab(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, sTot
    Set rst = CreateObject("ADODB.Recordset")

    sql = "select labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from ( "
    sql = sql & "select labtestid, qty, unitcost, finalamt from investigation where visitationid='" & vst & "' "
    sql = sql & "union all "
    sql = sql & "select labtestid, qty, unitcost, finalamt from investigation2 where visitationid='" & vst & "') as t "
    sql = sql & " group by labtestid"
    
    cnt = pos
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            sTot = 0
            Do While Not .EOF
                cnt = cnt + 1
                unt = .fields("unt")
                qty = .fields("qty")
                tot = .fields("tot")
                gTot = gTot + tot
                sTot = sTot + tot
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    pos = cnt
    AddLab = sTot
End Function

'AddTreat
Function AddTreat(vst)
    Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, sTot
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select treatmentid,sum(qty) as qty, avg(unitcost) as unt,sum(finalamt) as tot from treatcharges where visitationid='" & vst & "' group by treatmentid"
    cnt = pos
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            sTot = 0
            Do While Not .EOF
                cnt = cnt + 1
                unt = .fields("unt")
                qty = .fields("qty")
                tot = .fields("tot")
                gTot = gTot + tot
                sTot = sTot + tot
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    pos = cnt
    AddTreat = gTot
End Function

Function AddWaiver(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select sum(PaidAmount) as tot from patientwaiveritems where visitationid='" & vst & "'"
  cnt = pos
  With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  .MoveFirst
  Do While Not .EOF
  If Not IsNull(.fields("tot")) Then
    tot = .fields("tot")
    gWv = gWv + tot
  End If
  .MoveNext
  Loop
  End If
  .Close
  End With
  Set rst = Nothing
  pos = cnt
  AddWaiver = gWv
End Function

Function GetReturnTot(vst, dg)
    Dim rstTblSql, sql, ot
    Set rstTblSql = CreateObject("ADODB.Recordset")
    ot = 0
    With rstTblSql

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

Function GetReturnQty(vst, dg)
    Dim rstTblSql, sql, ot
    Set rstTblSql = CreateObject("ADODB.Recordset")
    ot = 0
    With rstTblSql

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

Function getDatePeriodFromDelim(strDelimPeriod)
        
    Dim arPeriod, periodStart, periodEnd

    Dim arOut(1)

    arPeriod = Split(strDelimPeriod, "||")

    If UBound(arPeriod) >= 0 Then
        periodStart = arPeriod(0)
    End If

    If UBound(arPeriod) >= 1 Then
        periodEnd = arPeriod(1)
    End If

    periodStart = makeDatePeriod(Trim(periodStart), periodEnd, "0:00:00")
    periodEnd = makeDatePeriod(Trim(periodEnd), periodStart, "23:59:59")

    arOut(0) = periodStart
    arOut(1) = periodEnd

    getDatePeriodFromDelim = arOut

End Function

Function makeDatePeriod(strDateStart, defaultDate, strTime)

    If IsDate(strDateStart) Then
        makeDatePeriod = FormatDate(strDateStart) & " " & Trim(strTime)
    Else

        If IsDate(defaultDate) Then
            makeDatePeriod = FormatDate(defaultDate) & " " & Trim(strTime)
        Else
            makeDatePeriod = FormatDate(Now()) & " " & Trim(strTime)
        End If
    End If

End Function

Function AddPayments(vst, pat, recNo)
    Dim rst, sql, ot, cnt, hdr, dsc, pd, cn, bal, dt, sDt, eDt, dDt, cn2, gPaid
    Dim arr, ul, num, whcls, r, rCnt, sqlOk, sql2
    Set rst = CreateObject("ADODB.Recordset")
    gPaid = 0
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
      .MoveNext
      Loop
      End If
      .Close
      End With
    End If 'Sql
    Set rst = Nothing
    AddPayments = gPaid
End Function

Function AddUsedPayments(vst, pat, recNo)
  Dim rst, rst2, sql, ot, cnt, hdr, dsc, pd, cn, bal, dt, sDt, eDt, dDt
  Dim cnt2, cn2, rec, usd, uCnt, sql2
  Dim arr, ul, num, whcls, r, rCnt, sqlOk
  Set rst = CreateObject("ADODB.Recordset")
  Set rst2 = CreateObject("ADODB.Recordset")
  uCnt = 0
  
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
              dsc = dsc & "[V#&nbsp;:&nbsp;" & rst2.fields("VisitationID") & "]&nbsp;" & GetComboName("PaymentType", rst2.fields("PaymentTypeID"))
              rst2.MoveNext
            Loop
          End If
          rst2.Close
          gUsed = gUsed + usd
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
  AddUsedPayments = gUsed
End Function

Sub addCSS()
  With response
    .write " <style> "
    .write "    .anaesthesia, .anaesthesia th, .anaesthesia td{ "
    .write "        border: 1px solid silver; "
    .write "        border-collapse: collapse; "
    .write "        padding: 5px; "
    .write "    } "
    .write "    .anaesthesia{ "
    .write "        width: 80vw; "
    .write "        margin: 0 auto; "
    .write "        font-family: sans-serif; "
    .write "        font-size: 13px; "
    .write "        box-sizing: border-box; "
    .write "    }"
    .write "    .anaesthesia tr{page-break-inside:avoid; "
    .write "        page-break-after:auto "
    .write "    } "
    .write "    .anaesthesia th, .anaesthesia td { "
    .write "        border: 1px solid silver; "
    .write "        text-align: center; "
    .write "        padding: 5px; "
    .write "        font-size:13px; "
    .write "        margin: 0 auto; "
    .write "    } "
    .write "    .tHead{ "
    .write "        position: sticky; top: 0; "
    .write "    }  "
    .write "    .queryData td{ "
    .write "        font-size: 12; "
    .write "    }  "
    .write "    .anaesthesia th{ "
    .write "        background-color: blanchedalmond; "
    .write "        text-align: center; "
    .write "        font-weight: bold;"
    .write "        font-size: 14px;color:#000;"
    .write "   } "
    .write " </style> "
  End With
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
