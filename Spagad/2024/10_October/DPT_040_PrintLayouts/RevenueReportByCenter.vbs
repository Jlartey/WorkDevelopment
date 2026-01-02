'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Sub GenerateRevenueReport()
  Dim sql, rst, ot, cnt, rst0, sto, stoNm, amt, totAmt, schAmt, rAmt
  Dim bGp, bGpCnt, bGpNum, bGpAmt, fnd, cnt2, ky, kyNm, kyCnt, kyAmt
  Dim rst1, brAmt
  Dim arrBgp(1000, 2)
  Set rst0 = CreateObject("ADODB.Recordset")
  Set rst = CreateObject("ADODB.Recordset")
  Set rst1 = CreateObject("ADODB.Recordset")
  
  amt = 0
  totAmt = 0
  kyAmt = 0
  schAmt = 0
  brAmt = 0
  sql = "select * from Branch where BranchID='" & cBr & "'" '<>'NONE' order by Branchid"
  rst1.open sql, conn, 3, 4
  If rst1.RecordCount > 0 Then
    rst1.MoveFirst
    Do While Not rst1.EOF
      cBr = rst1.fields("BranchID")
      cBrNm = rst1.fields("BranchName")
  brAmt = 0
  amt = 0
  kyAmt = 0
  schAmt = 0
  sql = "select * from RevenueCenter order by RevenueCenterid"
  rst0.open sql, conn, 3, 4
  If rst0.RecordCount > 0 Then
    rst0.MoveFirst
    cnt = 0
    
    response.write "<table border=""1"" cellpadding=""3"" cellspacing=""0"" style=""font-size: 10pt; font-family: Arial; border-collapse:collapse;page-break-after:always"">"
    response.write "<tr>"
    response.write "<td><b>NO.</b></td>"
    response.write "<td><b>CODE NO.</b></td>"
    response.write "<td valign=""top"" align=""Center""><b>REVENUE CENTER [" & cBrNm & "]</b></td>"
    'response.write "<td valign=""top"" align=""right""><b>BILL GROUP</b></td>"
    response.write "<td valign=""top"" align=""right""><b>BILL AMOUNT</b></td>"
    response.write "</tr>"
    Do While Not rst0.EOF
      sto = rst0.fields("RevenueCenterID")
      amt = 0
      kyCnt = 0
      kyAmt = 0
      schAmt = 0
      'Visitation
      sql = "select specialistTypeid, sum(visitcost) as amt from visitation "
      sql = sql & " where RevenueCenterid='" & sto & "' and sponsorid<>'CANCEL'"
      sql = sql & " and BillMonthID='" & inpFlt & "' and BranchID='" & cBr & "' group by SpecialistTypeID"
      rst.open sql, conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
            amt = .fields("amt")
            ky = .fields("specialistTypeID")
            If amt > 0 Then
              kyCnt = kyCnt + 1
              totAmt = totAmt + amt
              schAmt = schAmt + amt
              brAmt = brAmt + amt
              If kyCnt = 1 Then
                stoNm = GetComboName("RevenueCenter", sto)
                response.write "<tr>"
                response.write "<td><b>-</b></td>"
                response.write "<td><b>" & UCase(sto) & "</b></td>"
                response.write "<td><b>" & UCase(stoNm) & "</b></td>"
                response.write "<td align=""right""><b>-</b></td>" 'Price
                response.write "</tr>"
              End If
              kyNm = GetComboName("SpecialistType", ky)
              response.write "<tr>"
              response.write "<td><b>" & CStr(kyCnt) & "</b></td>"
              response.write "<td>" & UCase(ky) & "</td>"
              response.write "<td>" & UCase(kyNm) & "</td>"
              response.write "<td align=""right"">" & FormatNumber(amt, 2, , , -1) & "</td>" 'Price
              response.write "</tr>"
            End If
            .MoveNext
          Loop
        End If
      End With
      rst.Close
      
      'NHIS Diagnosis
      sql = "select diagnosis.visitationid, max(diagnosis.treatmentcost) as amt from diagnosis,visitation,consultReview "
      sql = sql & " where diagnosis.visitationid=visitation.visitationID"
      sql = sql & " and diagnosis.consultReviewid=consultReview.consultReviewID"
      sql = sql & " and diagnosis.RevenueCenterid='" & sto & "' and diagnosis.sponsorid<>'CANCEL'"
      sql = sql & "  and visitation.BillMonthID='" & inpFlt & "' and consultreview.BranchID='" & cBr & "' group by diagnosis.visitationid"
      rst.open sql, conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          kyAmt = 0
          Do While Not .EOF
            amt = .fields("amt")
            ky = .fields("visitationID")
            If amt > 0 Then
              totAmt = totAmt + amt
              schAmt = schAmt + amt
              kyAmt = kyAmt + amt
              brAmt = brAmt + amt
            End If
            .MoveNext
          Loop
          If kyAmt > 0 Then
            kyCnt = kyCnt + 1
            If kyCnt = 1 Then
              stoNm = GetComboName("RevenueCenter", sto)
              response.write "<tr>"
              response.write "<td><b>-</b></td>"
              response.write "<td><b>" & UCase(sto) & "</b></td>"
              response.write "<td><b>" & UCase(stoNm) & "</b></td>"
              response.write "<td align=""right""><b>-</b></td>" 'Price
              response.write "</tr>"
            End If
            kyNm = "NHIS CONSULTATION"
            response.write "<tr>"
            response.write "<td><b>" & CStr(kyCnt) & "</b></td>"
            response.write "<td>CONNHIS</td>"
            response.write "<td>" & UCase(kyNm) & "</td>"
            response.write "<td align=""right"">" & FormatNumber(kyAmt, 2, , , -1) & "</td>" 'Price
            response.write "</tr>"
          End If
        End If
      End With
      rst.Close
      
      'Lab
      sql = "select investigation.labtestid, sum(investigation.finalamt) as amt from investigation,visitation "
      sql = sql & " where investigation.visitationid=visitation.visitationID"
      sql = sql & " and investigation.RevenueCenterid='" & sto & "' and investigation.sponsorid<>'CANCEL'"
      sql = sql & " and visitation.BillMonthID='" & inpFlt & "' and investigation.BranchID='" & cBr & "' group by investigation.labtestid"
      rst.open sql, conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
            amt = .fields("amt")
            ky = .fields("labTestID")
            If amt > 0 Then
              kyCnt = kyCnt + 1
              totAmt = totAmt + amt
              schAmt = schAmt + amt
              brAmt = brAmt + amt
              If kyCnt = 1 Then
                stoNm = GetComboName("RevenueCenter", sto)
                response.write "<tr>"
                response.write "<td><b>-</b></td>"
                response.write "<td><b>" & UCase(sto) & "</b></td>"
                response.write "<td><b>" & UCase(stoNm) & "</b></td>"
                response.write "<td align=""right""><b>-</b></td>" 'Price
                response.write "</tr>"
              End If
              kyNm = GetComboName("labTest", ky)
              response.write "<tr>"
              response.write "<td><b>" & CStr(kyCnt) & "</b></td>"
              response.write "<td>" & UCase(ky) & "</td>"
              response.write "<td>" & UCase(kyNm) & "</td>"
              response.write "<td align=""right"">" & FormatNumber(amt, 2, , , -1) & "</td>" 'Price
              response.write "</tr>"
            End If
            .MoveNext
          Loop
        End If
      End With
      rst.Close
      'Lab Walkin
      sql = "select labtestid, sum(finalamt) as amt from investigation "
      sql = sql & " where RevenueCenterid='" & sto & "' and sponsorid<>'CANCEL'"
      sql = sql & " and WorkingMonthID='" & inpFlt & "' and visitationid='E01' and investigation.BranchID='" & cBr & "' group by labtestid"
      rst.open sql, conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
            amt = .fields("amt")
            ky = .fields("labTestID")
            If amt > 0 Then
              kyCnt = kyCnt + 1
              totAmt = totAmt + amt
              schAmt = schAmt + amt
              brAmt = brAmt + amt
              If kyCnt = 1 Then
                stoNm = GetComboName("RevenueCenter", sto)
                response.write "<tr>"
                response.write "<td><b>-</b></td>"
                response.write "<td><b>" & UCase(sto) & "</b></td>"
                response.write "<td><b>" & UCase(stoNm) & "</b></td>"
                response.write "<td align=""right""><b>-</b></td>" 'Price
                response.write "</tr>"
              End If
              kyNm = GetComboName("labTest", ky)
              response.write "<tr>"
              response.write "<td><b>" & CStr(kyCnt) & "</b></td>"
              response.write "<td>" & UCase(ky) & "</td>"
              response.write "<td>" & UCase(kyNm) & "</td>"
              response.write "<td align=""right"">" & FormatNumber(amt, 2, , , -1) & "</td>" 'Price
              response.write "</tr>"
            End If
            .MoveNext
          Loop
        End If
      End With
      rst.Close
      
      'DrugSale
      sql = "select drugsaleitems.drugid, sum(drugsaleitems.finalamt) as amt from drugsaleitems,visitation "
      sql = sql & " where drugsaleitems.visitationid=visitation.visitationID"
      sql = sql & " and drugsaleitems.RevenueCenterid='" & sto & "' and drugsaleitems.sponsorid<>'CANCEL'"
      sql = sql & "  and visitation.BillMonthID='" & inpFlt & "' and drugsaleitems.BranchID='" & cBr & "' group by drugsaleitems.drugid"
      rst.open sql, conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
            ky = .fields("DrugID")
            amt = .fields("amt")
            rAmt = GetDrugRetAmtSch(ky, sto, inpFlt)
            amt = amt - rAmt
            If amt > 0 Then
              kyCnt = kyCnt + 1
              totAmt = totAmt + amt
              schAmt = schAmt + amt
              brAmt = brAmt + amt
              If kyCnt = 1 Then
                stoNm = GetComboName("RevenueCenter", sto)
                response.write "<tr>"
                response.write "<td><b>-</b></td>"
                response.write "<td><b>" & UCase(sto) & "</b></td>"
                response.write "<td><b>" & UCase(stoNm) & "</b></td>"
                response.write "<td align=""right""><b>-</b></td>" 'Price
                response.write "</tr>"
              End If
              kyNm = GetComboName("Drug", ky)
              response.write "<tr>"
              response.write "<td><b>" & CStr(kyCnt) & "</b></td>"
              response.write "<td>" & UCase(ky) & "</td>"
              response.write "<td>" & UCase(kyNm) & "</td>"
              response.write "<td align=""right"">" & FormatNumber(amt, 2, , , -1) & "</td>" 'Price
              response.write "</tr>"
            End If
            .MoveNext
          Loop
        End If
      End With
      rst.Close
      
      'Drug Walkin
      sql = "select drugid, sum(finalamt) as amt from drugsaleitems "
      sql = sql & " where RevenueCenterid='" & sto & "' and sponsorid<>'CANCEL'"
      sql = sql & " and WorkingMonthID='" & inpFlt & "' and visitationid='E01' and drugsaleitems.BranchID='" & cBr & "' group by drugid"
      rst.open sql, conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
            amt = .fields("amt")
            ky = .fields("drugID")
            rAmt = GetDrugRetAmtSch2(ky, sto, inpFlt)
            amt = amt - rAmt
            If amt > 0 Then
              kyCnt = kyCnt + 1
              totAmt = totAmt + amt
              schAmt = schAmt + amt
              brAmt = brAmt + amt
              If kyCnt = 1 Then
                stoNm = GetComboName("RevenueCenter", sto)
                response.write "<tr>"
                response.write "<td><b>-</b></td>"
                response.write "<td><b>" & UCase(sto) & "</b></td>"
                response.write "<td><b>" & UCase(stoNm) & "</b></td>"
                response.write "<td align=""right""><b>-</b></td>" 'Price
                response.write "</tr>"
              End If
              kyNm = GetComboName("drug", ky)
              response.write "<tr>"
              response.write "<td><b>" & CStr(kyCnt) & "</b></td>"
              response.write "<td>" & UCase(ky) & "</td>"
              response.write "<td>" & UCase(kyNm) & "</td>"
              response.write "<td align=""right"">" & FormatNumber(amt, 2, , , -1) & "</td>" 'Price
              response.write "</tr>"
            End If
            .MoveNext
          Loop
        End If
      End With
      rst.Close
      
      'Treatment
      sql = "select TreatCharges.Treatmentid, sum(TreatCharges.finalamt) as amt from TreatCharges,visitation,consultReview "
      sql = sql & " where TreatCharges.visitationid=visitation.visitationID"
      sql = sql & " and TreatCharges.consultReviewid=consultReview.consultReviewID"
      sql = sql & " and TreatCharges.RevenueCenterid='" & sto & "' and TreatCharges.sponsorid<>'CANCEL'"
      sql = sql & "  and visitation.BillMonthID='" & inpFlt & "' and consultreview.BranchID='" & cBr & "' group by TreatCharges.Treatmentid"
      rst.open sql, conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
            amt = .fields("amt")
            ky = .fields("TreatmentID")
            If amt > 0 Then
              kyCnt = kyCnt + 1
              totAmt = totAmt + amt
              schAmt = schAmt + amt
              brAmt = brAmt + amt
              If kyCnt = 1 Then
                stoNm = GetComboName("RevenueCenter", sto)
                response.write "<tr>"
                response.write "<td><b>-</b></td>"
                response.write "<td><b>" & UCase(sto) & "</b></td>"
                response.write "<td><b>" & UCase(stoNm) & "</b></td>"
                response.write "<td align=""right""><b>-</b></td>" 'Price
                response.write "</tr>"
              End If
              kyNm = GetComboName("Treatment", ky)
              response.write "<tr>"
              response.write "<td><b>" & CStr(kyCnt) & "</b></td>"
              response.write "<td>" & UCase(ky) & "</td>"
              response.write "<td>" & UCase(kyNm) & "</td>"
              response.write "<td align=""right"">" & FormatNumber(amt, 2, , , -1) & "</td>" 'Price
              response.write "</tr>"
            End If
            .MoveNext
          Loop
        End If
      End With
      rst.Close
      
      'Service Walkin
      sql = "select treatcharges.treatmentid, sum(finalamt) as amt from treatcharges,consultReview "
      sql = sql & " where treatcharges.RevenueCenterid='" & sto & "' and treatcharges.sponsorid<>'CANCEL'"
      sql = sql & " and TreatCharges.consultReviewid=consultReview.consultReviewID"
      sql = sql & " and treatcharges.WorkingMonthID='" & inpFlt & "' and treatcharges.visitationid='E01' and consultreview.BranchID='" & cBr & "' group by treatcharges.treatmentid"
      rst.open sql, conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
            amt = .fields("amt")
            ky = .fields("treatmentID")
            If amt > 0 Then
              kyCnt = kyCnt + 1
              totAmt = totAmt + amt
              schAmt = schAmt + amt
              brAmt = brAmt + amt
              If kyCnt = 1 Then
                stoNm = GetComboName("RevenueCenter", sto)
                response.write "<tr>"
                response.write "<td><b>-</b></td>"
                response.write "<td><b>" & UCase(sto) & "</b></td>"
                response.write "<td><b>" & UCase(stoNm) & "</b></td>"
                response.write "<td align=""right""><b>-</b></td>" 'Price
                response.write "</tr>"
              End If
              kyNm = GetComboName("treatment", ky)
              response.write "<tr>"
              response.write "<td><b>" & CStr(kyCnt) & "</b></td>"
              response.write "<td>" & UCase(ky) & "</td>"
              response.write "<td>" & UCase(kyNm) & "</td>"
              response.write "<td align=""right"">" & FormatNumber(amt, 2, , , -1) & "</td>" 'Price
              response.write "</tr>"
            End If
            .MoveNext
          Loop
        End If
      End With
      rst.Close
      If schAmt > 0 Then
        response.write "<tr>"
        response.write "<td>-</td>"
        response.write "<td colspan=""2"" align=""center""><b>Sub Totals [" & stoNm & "]</b></td>"
        response.write "<td align=""right""><b>" & FormatNumber(schAmt, 2, , , -1) & "</b></td>" 'Price
        response.write "</tr>"
      End If
      rst0.MoveNext
    Loop
    'G Total
    response.write "<tr>"
    response.write "<td>-</td>"
    response.write "<td colspan=""2"" align=""center""><b>TOTALS [" & cBrNm & "]</b></td>"
    response.write "<td align=""right""><b>" & FormatNumber(brAmt, 2, , , -1) & "</b></td>" 'Price
    response.write "</tr>"
    'response.write "</table>"
  End If 'rst0
  
  rst0.Close
  
  rst1.MoveNext
    Loop
    'G Total
    response.write "<tr>"
    response.write "<td>-</td>"
    response.write "<td colspan=""2"" align=""center""><b>TOTALS</b></td>"
    response.write "<td align=""right""><b>" & FormatNumber(totAmt, 2, , , -1) & "</b></td>" 'Price
    response.write "</tr>"
    response.write "</table>"
  End If 'rst0
  
  rst1.Close
  Set rst = Nothing
  Set rst1 = Nothing
  Set rst0 = Nothing
End Sub

Function GetDrugRetAmtSch(drg, sto, mth)
  Dim sql, amt, rst
  Set rst = CreateObject("ADODB.Recordset")
  'DrugReturn
  amt = 0
  sql = "select drugReturnitems.drugid, sum(drugReturnitems.finalamt) as amt from drugReturnitems,visitation "
  sql = sql & " where drugReturnitems.visitationid=visitation.visitationID and drugReturnitems.drugid='" & drg & "'"
  sql = sql & " and drugReturnitems.RevenueCenterid='" & sto & "' and drugReturnitems.sponsorid<>'CANCEL'"
  sql = sql & "  and visitation.billmonthid='" & mth & "' and drugreturnitems.BranchID='" & cBr & "' group by drugReturnitems.drugid"
  rst.open sql, conn, 3, 4
  With rst
    If .RecordCount > 0 Then
      .MoveFirst
      If Not IsNull(.fields("amt")) Then
        If IsNumeric(.fields("amt")) Then
          amt = .fields("amt")
        End If
      End If
    End If
  End With
  GetDrugRetAmtSch = amt
  rst.Close
End Function
Function GetDrugRetAmtSch2(drg, sto, mth)
  Dim sql, amt, rst
  Set rst = CreateObject("ADODB.Recordset")
  'DrugReturn
  amt = 0
  sql = "select drugid, sum(finalamt) as amt from drugreturnitems "
      sql = sql & " where RevenueCenterid='" & sto & "' and sponsorid<>'CANCEL' and drugid='" & drg & "'"
      sql = sql & " and workingmonthid='" & mth & "' and visitationid='E01' and BranchID='" & cBr & "' group by drugid"
  rst.open sql, conn, 3, 4
  With rst
    If .RecordCount > 0 Then
      .MoveFirst
      If Not IsNull(.fields("amt")) Then
        If IsNumeric(.fields("amt")) Then
          amt = .fields("amt")
        End If
      End If
    End If
  End With
  GetDrugRetAmtSch2 = amt
  rst.Close
End Function
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
'////////////////////////////////////START SCRIPT //////////////////////////////////
Dim dat1, dat2, arr, num, ul, inpFlt, cBr, cBrNm

dat1 = FormatDate(Now()) & " 0:00:00"
dat2 = FormatDate(Now()) & " 23:59:59"

server.scripttimeout = 1800
inpFlt = Trim(Request.QueryString("printfilter0"))
cBr = Trim(Request.QueryString("printfilter1"))
arr = Split(inpFlt, "||")
ul = UBound(arr)
If ul >= 0 Then
  For num = 0 To ul
    If num = 0 Then
      dat1 = Trim(arr(0))
    ElseIf num = 1 Then
      dat2 = Trim(arr(1))
    End If
  Next
  If IsDate(dat1) Then
    If IsDate(dat2) Then
    
    Else 'No Dat2
      dat2 = FormatDate(CDate(dat1)) & " 23:59:59"
    End If
  Else 'No Dat1
    If IsDate(dat2) Then
      dat1 = FormatDate(CDate(dat2)) & " 0:00:00"
    Else 'No Dat2
      fnd = False
    End If
  End If
End If

response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""0"" cellpadding=""0"">"
response.write "<tr>"
AddReportHeader
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
response.write "</tr>"

'Title
response.write "<tr>"
response.write "<td align=""center"" bgcolor=""#FFFFFF"" >"
response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" bgcolor=""White"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:9pt"">"
response.write "<tr>"
response.write "<td>TITLE :&nbsp;&nbsp;&nbsp;&nbsp;</td>"
response.write "<td>REVENUE REPORT BY CENTERS</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td>CLINIC :&nbsp;&nbsp;&nbsp;&nbsp;</td>"
response.write "<td>" & GetComboName("Branch", cBr) & "</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td>FOR :&nbsp;&nbsp;&nbsp;&nbsp;</td>"
response.write "<td>" & GetWorkingMonthName(inpFlt) & "</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td>DATE :&nbsp;&nbsp;&nbsp;&nbsp;</td>"
response.write "<td> AS AT&nbsp;&nbsp;&nbsp;&nbsp;" & (FormatDateDetail(Now())) & "</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td></tr>"


response.write "<tr>"
response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""left"">"
response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 9pt; font-family: Arial"">"
response.write "<tr>"
response.write "<td align=""center"">"

response.write "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
'response.write "<tr><td valign=""top"" colspan=""3""><hr color=""#999999"" size=""1"">"
'response.write "</td></tr>"
response.write "<tr><td valign=""top"" colspan=""3"" align=""center"">"
GenerateRevenueReport
response.write "</td>"
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
'None
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
