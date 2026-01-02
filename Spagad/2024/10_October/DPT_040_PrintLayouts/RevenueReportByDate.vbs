'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Sub GenerateRevenueReport()
  Dim sql, rst, ot, cnt, rst0, sto, stoNm, amt, totAmt, hrf, insGrp, cPay, cPayTot, whC
  Set rst0 = CreateObject("ADODB.Recordset")
  Set rst = CreateObject("ADODB.Recordset")

  amt = 0
  totAmt = 0
  cPay = 0
  cPayTot = 0
  sql = "select * from Sponsor where sponsorid<>'CANCEL' order by sponsorid"
  rst0.open qryPro.FltQry(sql), conn, 3, 4
  If rst0.RecordCount > 0 Then
    rst0.MoveFirst
    cnt = 0
    
    response.write "<table border=""1"" cellpadding=""3"" cellspacing=""0"" style=""font-size: 10pt; font-family: Arial; border-collapse:collapse;page-break-after:always"">"
    response.write "<tr>"
    response.write "<td><b>NO.</b></td>"
    response.write "<td><b>CODE NO.</b></td>"
    response.write "<td valign=""top"" align=""Center""><b>SPONSOR</b></td>"
    response.write "<td valign=""top"" align=""right""><b>BILL AMOUNT</b></td>"
    response.write "<td valign=""top"" align=""right""><b>CO-PAY AMT</b></td>"
    response.write "<td valign=""top"" align=""right""><b>FINAL AMT</b></td>"
    response.write "<td valign=""top"" align=""right""><b>DETAIL</b></td>"
    response.write "</tr>"
    Do While Not rst0.EOF
      sto = rst0.fields("SponsorID")
      stoNm = rst0.fields("SponsorName")
      insGrp = Trim(GetVettingGroup(sto))
      amt = 0
      cPay = 0
      whC = ""
      'CorporateBill
      If (UCase(sto) = "NDUOM") And (UCase(inpFlt) <= "MTH201510") Then 'GM MEDICAL
        sql = "select sum(billAmt1) as amt,sum(billAmt4) as cAmt from corporatebill "
        sql = sql & " where sponsorid='" & sto & "'"
        'sql = sql & " and BillDate3 between '" & dat1 & "' and '" & dat2 & "'"
        sql = sql & " and BillMonthID='" & inpFlt & "'"
        whC = ""
      ElseIf (UCase(insGrp) = "INS") Or (UCase(insGrp) = "NHIS") Then
        sql = "select sum(billAmt1) as amt,sum(billAmt4) as cAmt from corporatebill "
        sql = sql & " where sponsorid='" & sto & "' and corpbillstatusid='P003'"
        'sql = sql & " and BillDate3 between '" & dat1 & "' and '" & dat2 & "'"
        sql = sql & " and BillMonthID='" & inpFlt & "'"
        whC = " and visitmodeid='V002'"
      Else
        sql = "select sum(billAmt1) as amt,sum(billAmt4) as cAmt from corporatebill "
        sql = sql & " where sponsorid='" & sto & "'"
        'sql = sql & " and BillDate3 between '" & dat1 & "' and '" & dat2 & "'"
        sql = sql & " and BillMonthID='" & inpFlt & "'"
        whC = ""
      End If
      rst.open qryPro.FltQry(sql), conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          If Not IsNull(.fields("amt")) Then
            If IsNumeric(.fields("amt")) Then
              amt = amt + CDbl(.fields("amt"))
            End If
          End If
          If Not IsNull(.fields("cAmt")) Then
            If IsNumeric(.fields("cAmt")) Then
              amt = amt - CDbl(.fields("cAmt"))
            End If
          End If
        End If
      End With
      
      rst.Close
      'PatientBill 'SELF PAY
      If UCase(sto) = "SELF" Then
        sql = "select sum(billAmt1) as amt,sum(billAmt4) as cAmt from Patientbill "
        'sql = sql & " where BillDate3 between '" & dat1 & "' and '" & dat2 & "'"
        sql = sql & " where BillMonthID='" & inpFlt & "'"
        rst.open qryPro.FltQry(sql), conn, 3, 4
        With rst
          If .RecordCount > 0 Then
            .MoveFirst
            If Not IsNull(.fields("amt")) Then
              If IsNumeric(.fields("amt")) Then
                amt = amt + CDbl(.fields("amt"))
              End If
            End If
            If Not IsNull(.fields("cAmt")) Then
            If IsNumeric(.fields("cAmt")) Then
              amt = amt - CDbl(.fields("cAmt"))
            End If
          End If
          End If
        End With
        rst.Close
      Else
        cPay = GetCoPayAmount(sto, inpFlt, whC)
      End If
      
      If amt > 0 Then
        cnt = cnt + 1
        response.write "<tr>"
        response.write "<td><b>" & CStr(cnt) & "</b></td>"
        response.write "<td>" & UCase(sto) & "</td>"
        response.write "<td>" & UCase(stoNm) & "</td>"
        response.write "<td align=""right"">" & FormatNumber(amt, 2, , , -1) & "</td>" 'Price
        
        If cPay > 0 Then
          hrf = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=CoPaymentList&PositionForTableName=WorkingDay&BillMonthID=" & inpFlt & "&SponsorID=" & sto
          response.write "<td align=""right""><a target=""_Blank"" href=""" & hrf & """>" & FormatNumber(cPay, 2, , , -1) & "</a></td>" 'coPay
        Else
          response.write "<td align=""right""></td>"
        End If
        response.write "<td align=""right"">" & FormatNumber(amt - cPay, 2, , , -1) & "</td>" 'FinalAmt
        
        hrf = "wpgSelectPrintLayout.asp?PositionForTableName=Sponsor&SponsorID=" & sto
        response.write "<td align=""center""><a target=""_Blank"" href=""" & hrf & """>Detail</a></td>"
        response.write "</tr>"
      End If
      totAmt = totAmt + amt
      cPayTot = cPayTot + cPay
      rst0.MoveNext
    Loop
    'G Total
    response.write "<tr>"
    response.write "<td>-</td>"
    response.write "<td colspan=""2"" align=""center""><b>TOTALS</b></td>"
    response.write "<td align=""right""><b>" & FormatNumber(totAmt, 2, , , -1) & "</b></td>" 'Price
    response.write "<td align=""right""><b>" & FormatNumber(cPayTot, 2, , , -1) & "</b></td>" 'Co-Pay
    response.write "<td align=""right""><b>" & FormatNumber(totAmt - cPayTot, 2, , , -1) & "</b></td>" 'FianlAmt
    response.write "<td>-</td>"
    response.write "</tr>"
    response.write "</table>"
  End If 'rst0
  
  rst0.Close
  Set rst = Nothing
  Set rst0 = Nothing
End Sub
'GetVettingGroup
Function GetVettingGroup(spn)
  Dim rst, sql, ot
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select VettingGroupID from InsuranceScheme where sponsorid='" & spn & "' order by vettinggroupid"
  ot = ""
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      ot = Trim(.fields("VettingGroupID"))
    End If
    .Close
  End With
  GetVettingGroup = ot
  Set rst = Nothing
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
'GetCoPayAmount
Function GetCoPayAmount(insr, mth, whC)
Dim rst, sql, wkd, sp, rst1, tot, gSt, cnt, patNm, pNo
Dim aCnt, vst, vty, ag, insno, comTot
Dim con, drg, lab, xr, vdt, dpt, cDpt, dptCnt, coPay

Set rst = CreateObject("ADODB.Recordset")
Set rst1 = CreateObject("ADODB.Recordset")

coPay = 0

sql = "select sum(visitValue1) as amt from visitation where billmonthid='" & mth & "'"
sql = sql & " and Sponsorid='" & insr & "' " & whC & " and VisitValue1>0"

With rst1
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
  .MoveFirst
  If Not IsNull(.fields("amt")) Then
    If IsNumeric(.fields("amt")) Then
      coPay = CDbl(.fields("amt"))
    End If
  End If
End If
.Close
End With
GetCoPayAmount = coPay
Set rst = Nothing
Set rst1 = Nothing
End Function
'////////////////////////////////////START SCRIPT //////////////////////////////////
Dim dat1, dat2, arr, num, ul, inpFlt

dat1 = FormatDate(Now()) & " 0:00:00"
dat2 = FormatDate(Now()) & " 23:59:59"
inpFlt = Trim(Request.QueryString("printfilter0"))
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
response.write "<td>REVENUE REPORT BY SPONSORS</td>"
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
