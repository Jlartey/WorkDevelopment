Dim reportHeader, staffID, userID, BranchID
Dim arrAggFldLst, aggrFldLst, arrFldLstCnt, majorTbl
Dim aggrStruct 'specifies whether detail report and aggregate report are turned on
Dim date1, date2, dateRange

InitValues
ReportHeaderDetails
ProcessCode

Sub InitValues()
Dim arr, num, ul, inpFlt
    reportHeader = "CASHIER RECEIPT REPORT"
    majorTbl = "Receipt"
    userID = UCase(Trim(uName))
    staffID = UCase(Trim(Session("StaffID")))
    BranchID = UCase(Trim(request("PrintFilter0")))
    reportHeader = "CASHIER RECEIPT REPORT - " & GetBranch()
    
    

date1 = FormatDate(Now()) & " 0:00:00"
date2 = FormatDate(Now()) & " 23:59:59"
inpFlt = Trim(request.queryString("printfilter1"))
arr = Split(inpFlt, "||")
ul = UBound(arr)
If ul >= 0 Then
  date1 = ""
  date2 = ""
  For num = 0 To ul
    If num = 0 Then
      date1 = Trim(arr(0))
    ElseIf num = 1 Then
      date2 = Trim(arr(1))
    End If
  Next
  If IsDate(date1) Then
    If IsDate(date2) Then
    
    Else 'No date2
      date2 = FormatDate(CDate(date1)) & " 23:59:59"
    End If
  Else 'No date1
    If IsDate(date2) Then
      date1 = FormatDate(CDate(date2)) & " 0:00:00"
    Else 'No date2
      date1 = FormatDate(Now()) & " 0:00:00"
      date2 = FormatDate(Now()) & " 23:59:59"
      fnd = False
    End If
  End If
End If
    
'    str = Split(UCase(Trim(request("PrintFilter"))), "||")
'    date1 = GetDateStartingToday(Now)
'    date2 = GetDateEndingToday(Now)
'    If UBound(str) >= 1 Then
'        If IsDate(str(0)) And IsDate(str(1)) Then
'            date1 = FormatDateDetail(str(0))
'            date2 = FormatDateDetail(str(1))
'        ElseIf IsDate(str(0)) Then
'            date1 = FormatDateDetail(str(0))
'            date1 = GetDateStartingToday(date1)
'            date2 = GetDateEndingToday(date1)
'        ElseIf IsDate(str(1)) Then
'            date2 = FormatDateDetail(str(1))
'            date2 = GetDateEndingToday(date2)
'            date1 = GetDateStartingToday(date2)
'        End If
'    ElseIf UBound(str) = 0 Then
'        If IsDate(str(0)) Then
'            date1 = FormatDateDetail(str(0))
'            date1 = GetDateStartingToday(date1)
'            date2 = GetDateEndingToday(date1)
'        End If
'    End If

    dateRange = "From " & FormatDateDetail(date1) & " To " & FormatDateDetail(date2)
End Sub

Function GetDateEndingToday(dt)
Dim rval
rval = Day(dt) & "-" & MonthName(Month(dt), True) & "-" & Year(dt) & " 23:59:59"
GetDateEndingToday = rval
End Function

Function GetDateStartingToday(dt)
Dim rval
rval = Day(dt) & "-" & MonthName(Month(dt), True) & "-" & Year(dt) & " 00:00:00"
GetDateStartingToday = rval
End Function

Function WriteDetailReport(brn)
Dim rs, rs2, sql, ky, kyNm, ky2, kyNm2, cnt, cnt2
Dim val1, val2, val3, val4, val5, val6, val7, val8, val9, val10
Dim tot1, tot2, tot3, tot4, tot5, tot6, tot7, tot8, tot9, tot10, tCnt
Dim stot1, stot2, stot3, sTot4, sTot5, sTot6, sTot7, sTot8, sTot9, sTot10, mRec, mLth
Set rs = CreateObject("ADODB.Recordset")
Set rs2 = CreateObject("ADODB.Recordset")
Dim maxLength
maxLength = 20
mLth = 20
sql = ""
response.write "<tr><td><table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""1"" bgcolor=""White""  style=""font-family: Arial; font-size: 9.5pt"">"
cnt = 0
tot8 = 0
tot7 = 0
sTot8 = 0
sTot7 = 0
With rs
    sql = "SELECT distinct PaymentmodeID from Receipt where (ReceiptDate Between '" & date1 & "' AND '" & date2 & "') "
    sql = sql & " AND BranchID='" & brn & "'"
    sql = sql & " ORDER BY PaymentModeID"
    .Open sql, conn, 3, 4
    If .RecordCount > 0 Then
        tCnt = .RecordCount
        Do While Not .EOF
            cnt = cnt + 1
            ky = .fields("paymentmodeid")
            kyNm = GetComboName("PaymentMode", ky)
            response.write "<tr>"
            response.write "<td><b>" & kyNm & "<b></td>"
            response.write "</tr>"
            response.write "<tr><td><table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""3"" bgcolor=""White"" style=""font-family: Arial; font-size: 8.5pt"">"
            response.write "<tr>"
            response.write "<td><b>No.</b></td>"
            response.write "<td><b>Pat No.</b></td>"
            response.write "<td><b>Name</b></td>"
            response.write "<td><b>Rec #</b></td>"
            response.write "<td><b>M.Rec</b></td>"
            response.write "<td><b>Date</b></td>"
            If ky = "P002" Then
                response.write "<td><b>Description</b></td>"
                response.write "<td><b>Cheq.</b></td>"
                response.write "<td><b>Bank</b></td>"
            Else
              response.write "<td colspan=""3""><b>Description</b></td>"
            End If
            response.write "<td align=""right""><b>Paid Amt</b></td>"
            response.write "<td align=""right""><b>Refun</b></td>"
            response.write "<td align=""right""><b>Final Amt</b></td>"
            response.write "<td align=""right""><b>Cashier</b></td>"
            response.write "</tr>"
            cnt2 = 0
            tot7 = 0
            tot8 = 0
            With rs2
                sql = "SELECT * from Receipt where (ReceiptDate Between '" & date1 & "' AND '" & date2 & "') "
                sql = sql & " AND BranchID='" & brn & "' and PaymentModeID='" & ky & "'"
                sql = sql & " ORDER BY ReceiptID"
                .Open sql, conn, 3, 4
                If .RecordCount > 0 Then
                    Do While Not .EOF
                        cnt2 = cnt2 + 1
                        'set values
                        If ky = "P002" Then
                            maxLength = 20
                            mLth = 20
                        Else
                            maxLength = 30
                            mLth = 25
                        End If
                        val1 = .fields("PatientID")
                        If UCase(Trim(val1)) = "P1" Then
                          val1 = val1 & "&nbsp;Walk&nbsp;In"
                        ElseIf UCase(Trim(val1)) = "P2" Then
                          val1 = val1 & "&nbsp;Walk&nbsp;In&nbsp;NG"
                        End If
                        val2 = Left(.fields("ReceiptName"), maxLength)
                        val3 = .fields("ReceiptID")
                        mRec = .fields("KeyPrefix")
                        val4 = GetShortDate(.fields("ReceiptDate"))
                        val10 = Left(Trim(.fields("Remarks")), mLth)
                        val5 = .fields("ChequeNo")
                        val6 = Left(GetComboName("CustomerBank", .fields("CustomerBankID")), 12)
                        
                        val7 = CDbl(.fields("ReceiptAmount1"))
                        val8 = CDbl(.fields("paidamounT"))
                        val9 = .fields("SystemUserID")
                        
                        tot7 = tot7 + val7
                        tot8 = tot8 + val8
                        
                        response.write "<tr>"
                        response.write "<td>" & cnt2 & ".</td>"
                        response.write "<td align=""left"">" & val1 & "</td>"
                        response.write "<td align=""left"">" & val2 & "</td>"
                        response.write "<td align=""left"">" & val3 & "</td>"
                        response.write "<td align=""left"">" & mRec & "</td>"
                        response.write "<td align=""left"">" & val4 & "</td>"
                        If ky = "P002" Then
                            response.write "<td align=""left"">" & val10 & "</td>"
                            response.write "<td align=""left"">" & val5 & "</td>"
                            response.write "<td align=""left"">" & val6 & "</td>"
                        Else
                          response.write "<td align=""left"" colspan=""3"">" & val10 & "</td>"
                        End If
                        response.write "<td align=""right"">" & (FormatNumber(CDbl(val7), 2, , , -1)) & "</td>"
                        If val8 = 0 Then
                          response.write "<td align=""right"">-</td>"
                        Else
                          response.write "<td align=""right"">" & (FormatNumber(CDbl(val8), 2, , , -1)) & "</td>"
                        End If
                        response.write "<td align=""right"">" & (FormatNumber(CDbl(val7 - val8), 2, , , -1)) & "</td>"
                        response.write "<td align=""right"">" & val9 & "</td>"
                        response.write "</tr>"
                        .MoveNext
                    Loop
                End If
                .Close
                sTot7 = sTot7 + tot7
                sTot8 = sTot8 + tot8
                response.write "<tr>"
                response.write "<td>&nbsp;</td>"
                response.write "<td colspan=""2""><b>Total</b></td>"
                response.write "<td align=""right"" colspan=""6""></td>"
                response.write "<td align=""right""><b>" & (FormatNumber(CDbl(tot7), 2, , , -1)) & "</b></td>"
                response.write "<td align=""right""><b>" & (FormatNumber(CDbl(tot8), 2, , , -1)) & "</b></td>"
                response.write "<td align=""right""><b>" & (FormatNumber(CDbl(tot7 - tot8), 2, , , -1)) & "</b></td>"
                response.write "<td>&nbsp;</td>"
                response.write "</tr>"
            End With
            If cnt = tCnt Then 'End of Loop
              response.write "<tr>"
              response.write "<td colspan=""13""><hr color=""#999999"" size=""1""></td>"
              response.write "</tr>"
              
              response.write "<tr>"
              response.write "<td colspan=""4""><b>OVERALL TOTALS </b></td>"
              response.write "<td align=""right"" colspan=""5""></td>"
              response.write "<td align=""right""><b>" & (FormatNumber(CDbl(sTot7), 2, , , -1)) & "</b></td>"
              response.write "<td align=""right""><b>" & (FormatNumber(CDbl(sTot8), 2, , , -1)) & "</b></td>"
              response.write "<td align=""right""><b>" & (FormatNumber(CDbl(sTot7 - sTot8), 2, , , -1)) & "</b></td>"
              response.write "<td>&nbsp;</td>"
              response.write "</tr>"
              
              response.write "<tr>"
              response.write "<td colspan=""13""><hr color=""#999999"" size=""1""></td>"
              response.write "</tr>"
            End If
            response.write "</table></td></tr>"
            .MoveNext
            
        Loop
    End If
    .Close
End With

response.write "</table></td></tr>"
WriteDetailReport = tot9

Set rs = Nothing
Set rs2 = Nothing
End Function

Sub ProcessCode()
    response.write "<tr><td><table width=""" & (PrintWidth) & """ border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"">"
    WriteDetailReport BranchID
    response.write "</table></td></tr>"
End Sub





Function GetFieldValue(tbl, kfld, vfld, kval)
Dim rs, sql, rval
Set rs = CreateObject("ADODB.Recordset")
rval = ""
With rs
    sql = "Select " & vfld & " From " & tbl & " Where " & kfld & "='" & kval & "'"
    .Open sql, conn, 3, 4
    If .RecordCount > 0 Then
        .MoveFirst
        If Not IsNull(.fields(vfld)) Then
            rval = CStr(.fields(vfld))
        End If
    End If
    .Close
End With
GetFieldValue = rval
Set rs = Nothing
End Function

Function GetBranch()
    GetBranch = GetFieldValue("Branch", "BranchID", "BranchName", BranchID)
End Function

Function GetBranchName(brn)
    GetBranchName = GetFieldValue("Branch", "BranchID", "BranchName", brn)
End Function

Function GetBranchType()
    GetBranchType = GetFieldValue("SysBranchType", "SysBranchTypeID", "SysBranchTypeName", GetFieldValue("Branch", "BranchID", "SysBranchTypeID", BranchID))
End Function

Function GetBranchTypeName(brn)
    GetBranchTypeName = GetFieldValue("SysBranchType", "SysBranchTypeID", "SysBranchTypeName", GetFieldValue("Branch", "BranchID", "SysBranchTypeID", brn))
End Function

Function GetStaff()
    GetStaff = GetFieldValue("Staff", "StaffID", "StaffName", staffID)
End Function

Function GetStaffName(stf)
    GetStaffName = GetFieldValue("Staff", "StaffID", "StaffName", stf)
End Function
Function GetShortDate(dt)
Dim sDate, dDt
    If IsDate(dt) Then
        dDt = CDate(dt)
        sDate = Day(dDt) & "/" & Month(dDt) & "/" & Right(Year(dDt), 2)
    Else
        sDate = ""
    End If
    GetShortDate = sDate
End Function

Sub ReportHeaderDetails()
    response.write "<tr>"
    response.write "<td align=""center"">"
    response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""0"" cellpadding=""0"">"
    response.write "<tr>"
    AddReportHeader
    response.write "</tr>"
    response.write "</table>"
    response.write "</td>"
    response.write "</tr>"
'    response.write "<tr>"
'    response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'    response.write "</tr>"
'    response.write "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:12pt"">"
'    response.write "THE TRUST HOSPITAL</td>"
'    response.write "</tr>"
'    response.write "<tr>"
'    response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'    response.write "</tr>"
    response.write "<tr>"
    response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
    response.write "</tr>"
    response.write "<tr>"
    response.write "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:10pt"">"
    response.write "" & reportHeader & "</td>"
    response.write "</tr>"
    response.write "<tr>"
    response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
    response.write "</tr>"
    
    response.write "<tr>"
    response.write "<td><table width=""100%""><tr>"
    response.write "<td style=""font-family: Arial; color: #111111; font-size:10pt"">" & GetBranch() & "</td>"
    response.write "<td style=""font-family: Arial; color: #111111; font-size:10pt"">" & dateRange & "</td>"
    response.write "<td  style=""font-family: Arial; color: #111111; font-size:10pt"">"
    response.write "" & "Printed by " & GetStaff() & " on " & FormatDateDetail(Now) & "</td>"
    response.write "</tr></table></td>"
    response.write "</tr>"
    
'    response.write "<tr>"
'    response.write "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-size:10pt"">"
'    response.write "" & "Printed by " & GetStaff() & " on " & FormatDateDetail(Now) & "</td>"
'    response.write "</tr>"
    response.write "<tr>"
    response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
    response.write "</tr>"
End Sub
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
