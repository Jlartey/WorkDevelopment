'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
  Dim str, lnkCnt, vst, rowClmHdr, admID
  Dim recCnt ''@bless - Count no. of records in AddCompInfo2
  recCnt = 0
  lnkCnt = 0
  vst = Trim(Request.queryString(("VisitationID")))
  If Len(vst) > 0 Then
    SetPageVariable "CurrentVisit", vst
  Else
    vst = Trim(GetPageVariable("CurrentVisit"))
  End If
  If Len(vst) > 0 Then
    DisplaySummary
  End If
Sub DisplaySummary()
  InitPrintLayout
  response.write "<table width=""100%"" border=""1"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse:collapse"">"
  response.write "<tr bgcolor=""#eeeeee"">"
    response.write "<td align=""center"" valign=""top"" colspan=""2"">"
    response.write "<font size=""5""><b>Discharge Patient For Billing</b>"
    If Not IsDiagnosed(vst) Then
      response.write "<font size=""8"" style=""color:red""><b>POLICY ALERT:: PATIENT MUST BE DIAGNOSISED BEFORE DISCHARGED FOR BILLING</b>"
    End If
    response.write "</td>"
  response.write "</tr>"

  response.write "<tr>"
    response.write "<td align=""right"" valign=""top"">"
    response.write "<b>Principal Diagnosis/Discharge Diagnosis"
    response.write "</td>"
    response.write "<td>"
    DisplayEMRComp2 "TH060", "", "TH06008"
    DisplayEMRComp2 "TH060", "", "TH06008V"
    response.write "</td>"
  response.write "</tr>"
  response.write "<tr>"
    response.write "<td align=""right"" valign=""top"">"
    response.write "<b>Investigations"
    response.write "</td>"
    response.write "<td>"
    DisplayEMRComp2 "TH060", "", "TH06009"
    response.write "</td>"
  response.write "</tr>"
  response.write "<tr>"
    response.write "<td align=""right"" valign=""top"">"
    response.write "<b>Management Plans"
    response.write "</td>"
    response.write "<td>"
    DisplayEMRComp2 "TH060", "", "TH06010"
    response.write "</td>"
  response.write "</tr>"
  response.write "<tr>"
    response.write "<td align=""right"" valign=""top"">"
    response.write "<b>Drug Prescriptions"
    response.write "</td>"
    response.write "<td>"
    DisplayEMRComp2 "TH060", "", "TH06018"
    response.write "</td>"
  response.write "</tr>"
  response.write "<tr>"
    response.write "<td align=""right"" valign=""top"">"
    response.write "<b>Discharge Summary"
    response.write "</td>"
    response.write "<td>"
    recCnt = 0
    recCnt = 5 ''@bless - 4 Dec 2023 //allow
    DisplayEMRComp2 "TS003", "", ""
    response.write "</td>"
  response.write "</tr>"

  If CInt(recCnt) >= 3 Then
    DisplayOutcome
  Else
    response.write "<tr>"
      response.write "<td align=""right"" valign=""top"">"
      response.write "<b>Discharge Summary"
      response.write "</td>"
      response.write "<td>"
      response.write "<font size='6' color='red'>Discharge Summary must be filled by the doctor first before patient can be 'Discharged For Billing' </font>"
      response.write "</td>"
    response.write "</tr>"
  End If

  response.write "</table>"
End Sub
Sub DisplayOutcome()
  Dim jb, wdNm, outC, ot
  Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
  jb = jSchd
  wdNm = Trim(GetComboName("Ward", jb))
  outC = Trim(GetComboNameFld("Visitation", vst, "MedicalOutComeID"))
  response.write outC & "<br>"
  If Len(wdNm) > 0 Then 'Ward Profile
    Select Case UCase(outC)
      Case "M006" '
        response.write "<tr>"
        response.write "<td align=""center"" valign=""top"" colspan=""2"" style=""font-size:14pt;font-weight:bold;color:red"">"
        response.write "Discharge cannot be done.<br>Patient has already been discharged for Billing."
        response.write "</td>"
        response.write "</tr>"
      Case Else
        ot = Trim(Request("cmdApply"))
        If Len(ot) > 0 Then
          conn.execute qryPro.FltQry("update Visitation set MedicalOutComeID='M006' where VisitationID='" & vst & "'")

          response.write "<tr>"
          response.write "<td align=""center"" valign=""top"" colspan=""2"" style=""font-size:14pt;font-weight:bold;color:#44aa44"">"
          response.write "Patient Folder has been submitted to Billing Unit for final Billing."
          response.write "</td>"
          response.write "</tr>"
        Else
          'Response.write "<tr>"
          'Response.write "<td align=""center"" valign=""top"" colspan=""2""><br>"
          'Response.write "<input style=""font-size:14;font-weight:bold; height:25;background-color:#aaaaff"" type=""submit"" value=""Discharge Patient For Billing"" name=""cmdApply"" ID=""cmdApply"">"
          'Response.write "<br></td>"
          'Response.write "</tr>"

          '13 Aug 2018 Start AdmissionPro A001->A003
          admID = Trim(GetActiveAdmission(vst))
          If Len(admID) > 0 Then
            response.write "<td align=""center"" valign=""top"" colspan=""2"">"
            'Clickable Url Link
            lnkCnt = lnkCnt + 1
            lnkID = "lnk" & CStr(lnkCnt)
            lnkText = "<b>&nbsp;&nbsp;Click to Start Discharge Pending Billing</b>"
            ' lnkUrl = "wpgAdmissionPro.asp?PageMode=AddNew&PullupData=AdmissionID||" & admID & "&TransProcessVal2ID=AdmissionPro-T003"
            lnkUrl = "wpgAdmissionPro.asp?PageMode=AddNew&PullupData=AdmissionID||" & admID & "&TransProcessVal2ID=AdmissionPro-T009" ''Discharge for Billing
            navPop = "POP"
            inout = "IN"
            fntSize = "16"
            fntColor = "#44aa44"
            bgColor = ""
            wdth = ""
            AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            response.write "</td>"
          Else
            response.write "<tr color=""red"">"
            response.write "<td align=""center"" valign=""top"" colspan=""2"" style=""font-size:14pt;font-weight:bold;color:red"">"
            response.write "This Patient does not have any ACTIVE Admission to Discharge Pending Billing"
            response.write "</td>"
            response.write "</tr>"
          End If
        End If
    End Select
  ElseIf IsBilling(jSchd) Or IsCashierSupervisor(jSchd) Then
    Select Case UCase(outC)
      Case "M006" '
        ot = Trim(Request("cmdApply"))
        If Len(ot) > 0 Then
          conn.execute qryPro.FltQry("update Visitation set MedicalOutComeID='M007' where VisitationID='" & vst & "'")

          response.write "<tr color=""green"">"
          response.write "<td align=""center"" valign=""top"" colspan=""2"" style=""font-size:14pt;font-weight:bold;color:#44aa44"">"
          response.write "Patient Billing has been completed."
          response.write "</td>"
          response.write "</tr>"
        Else
          response.write "<tr>"
          response.write "<td align=""center"" valign=""top"" colspan=""2""><br>"
          response.write "<input style=""font-size:14;font-weight:bold; height:25;background-color:#aaaaff"" type=""submit"" value=""Complete Patient Billing."" name=""cmdApply"" ID=""cmdApply"">"
          response.write "<br></td>"
          response.write "</tr>"
        End If
      Case Else
        If UCase(outC) = "M007" Then
          response.write "<tr color=""red"">"
          response.write "<td align=""center"" valign=""top"" colspan=""2"" style=""font-size:14pt;font-weight:bold;color:red"">"
          response.write "Billing cannot be completed.<br>Patient Billing is already completed."
          response.write "</td>"
          response.write "</tr>"

        Else
          response.write "<tr color=""red"">"
          response.write "<td align=""center"" valign=""top"" colspan=""2"" style=""font-size:14pt;font-weight:bold;color:red"">"
          response.write "Billing cannot be completed.<br>Patient has not yet been discharged for Billing."
          response.write "</td>"
          response.write "</tr>"
        End If
    End Select
  End If
End Sub
Sub DisplayEMRComp2(cmpSrcKy, cmpTabKy, cmpKy)
  Dim pat, cmpTb, cmpFd, cmpSrcTb, cmpSrcFd, rst, rst2, imgSrc, ot, sql
  Dim rqKy, rqFd, rqTb, rlTb, cmpTabFd, cmpTabTb, cmpVarTb, rqDtFd, url
  Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, tb, tbKy
  Set rst = CreateObject("ADODB.Recordset")
  Set rst2 = CreateObject("ADODB.Recordset")
  tb = "Visitation"
  tbKy = "VisitationID"
  cmpFd = "EMRComponentID"
  'Component Data
  If Len(cmpFd) > 0 Then
    sql = "select * from CompTableKey where CompTableKeyid='" & Trim(cmpFd) & "'"
    With rst2
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        .MoveFirst
        'Tables
        cmpTb = Trim(.fields("TableID"))
        cmpSrcTb = Trim(.fields("CompSourceTable"))
        rqTb = Trim(.fields("CompRequestTable"))
        rlTb = Trim(.fields("CompResultTable"))
        cmpVarTb = Trim(.fields("TableDefName"))
        cmpTabTb = Left(cmpVarTb, Len(cmpVarTb) - 3) & "CompTab"
        'Fields
        cmpSrcFd = GetUniqueKey(cmpSrcTb)
        rqFd = GetUniqueKey(rqTb)
        cmpTabFd = GetUniqueKey(cmpTabTb)
        'Keys
        If HasCompData(rlTb, rqTb, rqFd, vst, cmpSrcFd, cmpSrcKy, cmpTabFd, cmpTabKy, cmpFd, cmpKy) Then
          pat = GetComboNameFld("Visitation", vst, "PatientID")
          If Len(pat) > 0 Then
            sTb = rlTb
            sTbNm = GetComboName(cmpSrcTb, cmpSrcKy)
            response.write "<table width=""100%"" border=""1"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse:collapse"">"
'            response.write "<tr>"
'              response.write "<td height=""5"">"
'              response.write "&nbsp;</td>"
'            response.write "</tr>"
              If HasAccessRight(uName, "frm" & rlTb, "View") Then
               response.write "<tr>"
               response.write "<td>"
               'OpenPanelSection
               response.write "<table border=""0"" cellpadding=""2"" cellspacing=""10"">"
                   'Header
                   'Detail
                   If UCase(cmpTb) = "TESTCOMPONENT" Then
                     rqDtFd = "RequestDate"
                   ElseIf UCase(cmpTb) = "EMRCOMPONENT" Then
                     rqDtFd = "EMRDate"
                   End If
                   AddCompInfo2 rlTb, rqTb, rqFd, vst, cmpSrcTb, cmpSrcFd, cmpSrcKy, cmpTabTb, cmpTabFd, cmpTabKy, cmpTb, cmpFd, cmpKy, rqDtFd
                 response.write "</table>" 'EMR
               response.write "</td>"
              response.write "</tr>"
             End If
            response.write "</table>"
          End If 'pat
        End If 'HasCompData
        .Close
      End If 'rst2
    End With
  End If 'cmpFd
  Set rst = Nothing
  Set rst2 = Nothing
End Sub

Sub LoadCSS()
  Dim str
  str = ""
  str = str & "<style type='text/css' id=""styPrt"">"
  str = str & ".cpHdrTd{font-size:14pt;font-weight:bold}"
  str = str & ".cpHdrTr{background-color:#ddffdd}"
  str = str & ".cpHdrTd2{font-size:12pt;font-weight:bold}"
  str = str & ".cpHdrTr2{background-color:#ddffdd}" 'fafafa
  str = str & "</style>"

  response.write str
End Sub
Sub AddUrlLink(lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth)
  Dim plusMinus, imgName, lnkOpClNavPop, align
   plusMinus = ""
   imgName = ""
   align = ""
   lnkOpClNavPop = inout & "||" & navPop & "||800||600||CLOSE"
  AddPrtNavLink lnkID, plusMinus, imgName, lnkText, lnkUrl, lnkOpClNavPop, fntSize, fntColor, bgColor, align, wdth
End Sub

Function GetConvertImageFile(fdVl2)
  Dim objStrm, fdVl
  fdVl = ""
  If Not IsNull(fdVl2) Then
    'Convert from Binary to Text
    Set objStrm = CreateObject("ADODB.Stream")
    objStrm.open
    objStrm.Type = 1 'Binary
    objStrm.write fdVl2
    objStrm.position = 0
    objStrm.Type = 2 'Text
    fdVl = objStrm.readtext
    objStrm.Close
  End If
  GetConvertImageFile = fdVl
End Function
Function GetRowColMax(emrDat)
  Dim ot, ky
  ot = 0
  ky = Trim(emrDat)
  Select Case UCase(ky)
    Case "EMR050" 'Vital Sign
      ot = 9
      rowClmHdr = "Temp.||Pulse||Resp.||BP(Systolic)||BP(Diastolic)||SPO2||Weight(Kg)||Height(cm)||BMI"
  End Select
  GetRowColMax = ot
End Function
Function GetRowColPos(emrDat, emrCmp, col)
  Dim ot, emr, cmp, cl
  ot = 0
  emr = Trim(emrDat)
  cmp = Trim(emrCmp)
  cl = Trim(col)
  Select Case UCase(emr)
    Case "EMR050"
      Select Case UCase(cmp)
        Case "EMR05001"
          Select Case UCase(cl)
            Case "COLUMN2"
              ot = 1
            Case "COLUMN4"
              ot = 2
            Case "COLUMN6"
              ot = 3
          End Select
        Case "EMR05002"
          Select Case UCase(cl)
            Case "COLUMN2"
              ot = 4
            Case "COLUMN4"
              ot = 5
            Case "COLUMN6"
              ot = 6
          End Select
        Case "EMR05003"
          Select Case UCase(cl)
            Case "COLUMN2"
              ot = 7
            Case "COLUMN4"
              ot = 8
            Case "COLUMN6"
              ot = 9
          End Select
      End Select
  End Select
  GetRowColPos = ot
End Function
Function GetDispType(jb)
  Dim ot
  ot = ""
  Select Case UCase(jb)
    Case "BILLINGHEAD"
      ot = "BILLING"
    Case "CLAIMMANAGER"
      ot = "BILLING"
    Case "REGVISITWARD"
      ot = "BILLING"
    Case "REGVISITCON"
      ot = "BILLING"
    Case "CONSULTINGROOMVISIT"
      ot = "BILLING"
  End Select
  GetDispType = ot
End Function
Function IsBilling(jb)
  Dim ot, lst, arr, ul, num
  ot = False
  lst = "S17||S18||S87||S97||Cashier"
  arr = Split(lst, "||")
  ul = UBound(arr)
  For num = 0 To ul
    If UCase(Trim(arr(num))) = UCase(Trim(jb)) Then
      ot = True
      Exit For
    End If
  Next
  IsBilling = ot
End Function

Function IsCashierSupervisor(jb)
  Dim arr, ul, num, lst, ot
  ot = False
  lst = "ChiefCashier||CreditControl||M11||M11Head"
  arr = Split(lst, "||")
  ul = UBound(arr)
  For num = 0 To ul
    If UCase(Trim(arr(num))) = UCase(Trim(jb)) Then
      ot = True
      Exit For
    End If
  Next
  IsCashierSupervisor = ot
End Function

Sub AddCompInfo2(rlTb, rqTb, rqFd, vst, cmpSrcTb, cmpSrcFd, cmpSrcKy, cmpTabTb, cmpTabFd, cmpTabKy, cmpTb, cmpFd, cmpKy, rqDtFd)
  Dim rst, sql, ot, cnt, hdr, recKy, clr, hasAcc, compDt, sUsr, stf, stfNm, isNum, isTrd
  Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, curCompRq
  Dim tdColSp, tdAlign, secCnt, compKey, rstCompFld, fdVl, fdVl2, fd, tdAttr
  Dim wd, curCompTab, compTab, curCompDat, compDat, compCmp, compRq, hdr2
  Dim arrAll, ulAll, numAll, arrTd, ulTd, numTd, sTd0, sTd1, sTd2, sTd3
  Dim arr, ul, num, src0, src1, src2, src3, src4, src5, iFUrl, tdDim, srcKy
  Dim cmpHdrClr, cmpAlt2Clr, cmpAlt1Clr, pat
  Dim tdBgCl, tdVAln, tdStyl, tdDisp, tdLbl, sTd4, sTd5, sTd6, sTd7
  Dim arrGrp, ulGrp

  Dim arrRowClm(2, 21)
  Dim arrRowPrm(2, 21)
  Dim rowClmMax, cmplRow, rowClmNum, rowClmPos, arrRowHdr, numRowHdr, ulRowHdr

  rowClmMax = 0
  rowClmPos = 0
  rowClmHdr = ""
  cmplRow = False
  For rowClmNum = 1 To 20
     arrRowClm(1, rowClmNum) = ""
     arrRowPrm(1, rowClmNum) = ""
  Next

  Set rst = CreateObject("ADODB.Recordset")
  Set rstCompFld = CreateObject("ADODB.Recordset")
  isTrd = False
  sql = "select " & rlTb & "." & rqFd & "," & rlTb & "." & cmpSrcFd & "," & rlTb & "." & cmpTabFd & "," & rlTb & "." & cmpFd
  sql = sql & " ," & rqTb & "." & rqDtFd & "," & rqTb & ".SystemUserID"
  sql = sql & " ," & rlTb & ".Column1," & rlTb & ".Column2," & rlTb & ".Column3"
  sql = sql & " ," & rlTb & ".Column4," & rlTb & ".Column5," & rlTb & ".Column6"
  sql = sql & " from " & rlTb & "," & rqTb
  sql = sql & " where " & rlTb & "." & rqFd & "=" & rqTb & "." & rqFd
  If Len(Trim(cmpKy)) > 0 Then
    sql = sql & " and " & rlTb & "." & cmpFd & "='" & cmpKy & "'"
    isTrd = True
  ElseIf Len(Trim(cmpTabKy)) > 0 Then
    sql = sql & " and " & rlTb & "." & cmpTabFd & "='" & cmpTabKy & "'"
    isTrd = True
  ElseIf Len(Trim(cmpSrcKy)) > 0 Then
    sql = sql & " and " & rlTb & "." & cmpSrcFd & "='" & cmpSrcKy & "'"
    isTrd = True
  End If
  sql = sql & " and " & rqTb & ".VisitationID='" & vst & "'"
  sql = sql & " order by " & rqTb & "." & rqDtFd & "," & rlTb & "." & cmpSrcFd & "," & rlTb & "." & cmpTabFd & "," & rlTb & ".CompPos"
cnt = 0
wd = "100%"
secCnt = 0
compKey = cmpFd
cmpHdrClr = "#eeeeee"
cmpAlt2Clr = "#ffffff"
cmpAlt1Clr = "#ffffff"
'#d0d0d0||#ffffff||#ececff||5
With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    recCnt = .RecordCount
    .MoveFirst
    curCompTab = ""
    compTab = ""
    curCompDat = ""
    compDat = ""
    curCompRq = ""
    sUsr = ""
    response.write "<style>"
    response.write ".cmpTdSty {"
    response.write "border:1px solid #cccccc;"
    response.write "border-collapse: collapse;"
    response.write "}"
    response.write "</style>"
    pat = GetComboNameFld("Visitation", vst, "PatientID")
    Do While Not .EOF
      compTab = Trim(.fields(cmpTabFd))
      compDat = Trim(.fields(cmpSrcFd))
      compCmp = Trim(.fields(cmpFd))
      compRq = Trim(.fields(rqFd))
      compDt = Trim(.fields(rqDtFd))
      sUsr = Trim(.fields("SystemUserID"))
      If UCase(cmpSrcFd) < UCase(rqFd) Then
        srcKy = compDat & "-" & compRq
      Else
        srcKy = compRq & "-" & compDat
      End If
      If ((UCase(compDat) <> UCase(curCompDat)) Or (UCase(compRq) <> UCase(curCompRq))) Then
        If rowClmMax > 0 Then
          cnt = cnt + 1
          'Hdr
          response.write "<tr>"
          response.write "<td class=""cmpTdSty"" valign=""top"" colspan=""7"">"
          response.write "<table class=""cmpTdSty"" cellspacing=""0"" cellpadding=""3"" width=""100%"" style=""font-size:11pt""><tr>"
          arrRowHdr = Split(rowClmHdr, "||")
          ulRowHdr = UBound(arrRowHdr)
          For numRowHdr = 0 To ulRowHdr
            response.write "<td class=""cmpTdSty"" valign=""top"" align=""center""><b>" & arrRowHdr(numRowHdr) & "</b></td>"
          Next
          response.write "</tr>"
          'Row
          response.write "<tr>"
          For rowClmNum = 1 To rowClmMax
            response.write "<td class=""cmpTdSty"" valign=""top"" align=""center"">" & arrRowClm(1, rowClmNum) & "</td>"
          Next
          response.write "</tr></table></td>"
          response.write "<td>"
          If Not isTrd Then
            'Clickable Url Link
            lnkCnt = lnkCnt + 1
            lnkID = "lnk" & CStr(lnkCnt)
            lnkText = "Trend"
            lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=DisplayCompTrend"
            lnkUrl = lnkUrl & "&PatientID=" & pat & "&CompTableKeyID=" & cmpFd & "&" & cmpSrcFd & "=" & compDat
            navPop = "NAV"
            inout = "IN"
            fntSize = "10"
            fntColor = "#444488"
            bgColor = ""
            wdth = ""
            AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
          ElseIf isNum Then
            'Clickable Url Link
            lnkCnt = lnkCnt + 1
            lnkID = "lnk" & CStr(lnkCnt)
            lnkText = "Graph"
            lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=GraphCompTrend"
            lnkUrl = lnkUrl & "&PatientID=" & pat & "&CompTableKeyID=" & cmpFd & "&" & cmpSrcFd & "=" & compDat
            navPop = "NAV"
            inout = "IN"
            fntSize = "10"
            fntColor = "#444488"
            bgColor = ""
            wdth = ""
            AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
          End If
          response.write "</td>"
          response.write "</tr>"
          'Reset compile row
          rowClmPos = 0
          cmplRow = False
          For rowClmNum = 1 To 20
             arrRowClm(1, rowClmNum) = ""
             arrRowPrm(1, rowClmNum) = ""
          Next
        End If
        rowClmMax = GetRowColMax(compDat)
        secCnt = secCnt + 1
        If secCnt > 1 Then
          response.write "</table>"
          'ClosePanelSectionEMR
          response.write "</td>"
          response.write "</tr>"
        End If
        response.write "<tr>"
        response.write "<td>"
        'OpenPanelSectionEMR
        hdr2 = ""
        response.write "<table class=""cmpTdSty"" cellspacing=""0"" width=""100%"" cellpadding=""3"" style=""font-size:11pt"">"
        If (UCase(cmpFd) <> "TESTCOMPONENTID") And (Not isTrd) Then
          hdr2 = "New&nbsp;&nbsp;&nbsp;&nbsp;Edit&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        stf = GetComboNameFld("SystemUser", sUsr, "StaffID")
        stfNm = GetComboName("Staff", stf)
        hdr = GetComboName(cmpTabTb, compTab)
        hdr2 = hdr2 & "<b>By</b>:&nbsp;" & stfNm & "&nbsp;&nbsp;&nbsp;&nbsp;<b>On</b>:&nbsp;" & FormatDateDetail(compDt)
        response.write "<tr  bgcolor=""" & cmpHdrClr & """>"
        response.write "<td colspan=""8"" align=""right"">" & hdr2 & "</td>"
        response.write "</tr>"
        response.write "<tr bgcolor=""" & cmpHdrClr & """>"
        response.write "<td colspan=""7"" height=""20"" valign=""bottom"" class=""cpHdrTd2""><u>" & CStr(secCnt) & ".&nbsp;&nbsp;" & hdr & "</u></td>"
        response.write "<td class=""cpHdrTd2"">"
        If Not isTrd Then
          'Clickable Url Link
          lnkCnt = lnkCnt + 1
          lnkID = "lnk" & CStr(lnkCnt)
          lnkText = "<u><b>Trend</b></u>"
          lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=DisplayCompTrend"
          lnkUrl = lnkUrl & "&PatientID=" & pat & "&CompTableKeyID=" & compKey & "&" & cmpSrcFd & "=" & compDat & "&" & cmpTabFd & "=" & compTab
          navPop = "NAV"
          inout = "IN"
          fntSize = "10"
          fntColor = "#444488"
          bgColor = ""
          wdth = ""
          AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
          response.write "</td>"
        End If
        response.write "</td>"
        response.write "</tr>"
        curCompTab = compTab
        curCompDat = compDat
        curCompRq = compRq
        cnt = 0
      End If
      If rowClmMax = 0 Then
        cnt = cnt + 1
        response.write "<tr>"
        response.write "<td class=""cmpTdSty"" valign=""top"">" & GetComboName(cmpTb, compCmp) & "</td>"
      End If
      'Open CompField
      If UCase(appDbType) = "ORACLEDB" Then
        compCmp = GetORACaseSen("RecordKey", compCmp)
      End If
      isNum = False
      sql = "select * from compfield where comptablekeyid='" & compKey & "' and recordkey='" & compCmp & "' order by compfieldid"
      rstCompFld.open qryPro.FltQry(sql), conn, 3, 4
      If rstCompFld.RecordCount > 0 Then
        'Write Field Header
        rstCompFld.MoveFirst
        Do While Not rstCompFld.EOF
          fdVl = ""
          fdVl2 = ""
          tdColSp = "1"
          tdAlign = ""
          tdDim = ""
          fd = rstCompFld.fields("compfieldid").value
          tdAttr = ""
          If Not IsNull(rstCompFld.fields("SubTableFieldSource")) Then
            tdAttr = Trim(rstCompFld.fields("SubTableFieldSource"))
          End If
          ul = -1
          arrGrp = Split(tdAttr, "~~")
          ulGrp = UBound(arrGrp)
          arrAll = Split("", "%%")
          If ulGrp >= 0 Then
            arrAll = Split(arrGrp(0), "%%")
          End If
          ulAll = UBound(arrAll)
          If ulAll = 0 Then
            arr = Split(arrAll(0), "**")
            ul = UBound(arr)
          ElseIf ulAll = 1 Then
            arr = Split(arrAll(1), "**")
            ul = UBound(arr)
            'TD
            arrTd = Split(arrAll(0), "**")
            ulTd = UBound(arrTd)
            sTd0 = ""
            sTd1 = ""
            sTd2 = ""
            sTd3 = ""
            If ulTd >= 0 Then
              For numTd = 0 To ulTd
                Select Case numTd
                  Case 0
                    sTd0 = arrTd(0)
                  Case 1
                    sTd1 = arrTd(1)
                  Case 2
                    sTd2 = arrTd(2)
                  Case 3
                    sTd3 = arrTd(3)
                End Select
              Next
              Select Case UCase(Trim(sTd0))
                Case "COLSPAN"
                  If IsNumeric(sTd1) Then
                    tdColSp = sTd1
                  End If
                Case "COLSPANALIGN"
                  If IsNumeric(sTd1) Then
                    tdColSp = sTd1
                  End If
                  If Len(sTd2) > 0 Then
                    tdAlign = "align=""" & sTd2 & """"
                  End If
                Case "ALIGN"
                  If Len(sTd1) > 0 Then
                    tdAlign = "align=""" & sTd1 & """"
                  End If
                Case "TDPROP"
                  'Colspan
                  If IsNumeric(sTd1) Then
                    tdColSp = sTd1
                  End If
                  'Align
                  If Len(sTd2) > 0 Then
                    tdAlign = " align=""" & sTd2 & """ "
                  End If
                  'Valign
                  If Len(sTd3) > 0 Then
                    tdVAln = " valign=""" & sTd3 & """ "
                  End If
                   'BgColor
                  If Len(sTd4) > 0 Then
                    tdBgCl = " bgcolor=""" & sTd4 & """ "
                  End If
                   'Style
                  If Len(sTd5) > 0 Then
                    tdStyl = " style=""" & sTd5 & """ "
                  End If
                   'Header DisplayName
                  If Len(sTd6) > 0 Then
                    tdDisp = sTd6
                  End If
                  'LabelName
                  If Len(sTd7) > 0 Then
                    tdLbl = sTd7
                  End If
              End Select
            End If
          End If
          src0 = ""
          src1 = ""
          src2 = ""
          src3 = ""
          src4 = ""
          src5 = ""
          If ul >= 0 Then
            For num = 0 To ul
              Select Case num
                Case 0
                  src0 = arr(0)
                Case 1
                  src1 = arr(1)
                Case 2
                  src2 = arr(2)
                Case 3
                  src3 = arr(3)
                Case 4
                  src4 = arr(4)
                Case 5
                  src5 = arr(5)
              End Select
            Next
          End If

          If Not IsNull(.fields(fd)) Then
            fdVl = Trim(.fields(fd))
          End If
          If UCase(Trim(src0)) = "USERWRITEPAD" Then
            tdDim = ""
            If IsNumeric(src2) Then
              tdDim = tdDim & " width=""" & CStr(CInt(src2) + 30) & """ "
            End If
            If IsNumeric(src3) Then
              tdDim = tdDim & " height=""" & CStr(CInt(src3) + 80) & """ "
            End If
            iFUrl = "wpgWritingPadViewer.asp?PositionforTableName=" & cmpSrcTb & "&" & cmpSrcFd & "=" & compDat & "&" & cmpFd & "=" & compCmp & "&PageMode=ProcessSelect&" & rqFd & "=" & compRq & "&StoreField=" & fd & "&PadWidth=" & src2 & "&PadHeight=" & src3
            fdVl2 = "<iframe name=""iFrm" & secCnt & """ " & tdDim & " frameborder=""0"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"" src=""" & iFUrl & """></iframe>"
          ElseIf UCase(Trim(src0)) = "USERDICOM" Then
            tdDim = ""
            If IsNumeric(src2) Then
              tdDim = tdDim & " width=""" & CStr(CInt(src2) + 30) & """ "
            End If
            If IsNumeric(src3) Then
              tdDim = tdDim & " height=""" & CStr(CInt(src3) + 80) & """ "
            End If
            iFUrl = "wpgPrtPrintLayoutAll.asp?PositionforTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=dwvDicomViewer&" & rqFd & "=" & compRq & "&" & cmpSrcFd & "=" & compDat & "&" & cmpFd & "=" & compCmp & "&PadWidth=" & src2 & "&PadHeight=" & src3
            fdVl2 = "<iframe name=""iFrm" & secCnt & """ height=""" & src3 & """ width=""" & src2 & """ frameborder=""0"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"" src=""" & iFUrl & """></iframe>"
          ElseIf Len(fdVl) > 0 Then
            'fdVl2 = GetCompFldVal2(compKey, compCmp, fd, fdVl, "10", "1")
            fdVl2 = GetCompFldVal3(compKey, compCmp, fd, fdVl, "11", "1", srcKy, cmpSrcFd, compDat)
          End If
          If IsNumeric(fdVl2) Then
            isNum = True
          End If
          If IsNumeric(tdColSp) Then
            For num = 1 To (CInt(tdColSp) - 1)
              If Not rstCompFld.EOF Then
                rstCompFld.MoveNext
              End If
            Next
            If rowClmMax = 0 Then
              response.write "<td class=""cmpTdSty"" valign=""top"" " & tdDim & " " & tdAlign & " colspan=""" & tdColSp & """>" & fdVl2 & "</td>"
            Else
              rowClmPos = GetRowColPos(compDat, compCmp, fd)
              If rowClmPos > 0 Then
                arrRowClm(1, rowClmPos) = fdVl2
              End If
            End If
          Else
            If rowClmMax = 0 Then
              response.write "<td class=""cmpTdSty"" valign=""top"" " & tdDim & " " & tdAlign & ">" & fdVl2 & "</td>"
            Else
              rowClmPos = GetRowColPos(compDat, compCmp, fd)
              If rowClmPos > 0 Then
                arrRowClm(1, rowClmPos) = fdVl2
              End If
            End If
          End If
          If Not rstCompFld.EOF Then
            rstCompFld.MoveNext
          End If
        Loop
      End If
      rstCompFld.Close
      If rowClmMax = 0 Then
        response.write "<td>"
        If Not isTrd Then
          'Clickable Url Link
          lnkCnt = lnkCnt + 1
          lnkID = "lnk" & CStr(lnkCnt)
          lnkText = "Trend"
          lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=DisplayCompTrend"
          lnkUrl = lnkUrl & "&PatientID=" & pat & "&CompTableKeyID=" & cmpFd & "&" & cmpSrcFd & "=" & compDat & "&" & cmpFd & "=" & compCmp
          navPop = "NAV"
          inout = "IN"
          fntSize = "10"
          fntColor = "#444488"
          bgColor = ""
          wdth = ""
          AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        ElseIf isNum Then
          'Clickable Url Link
          lnkCnt = lnkCnt + 1
          lnkID = "lnk" & CStr(lnkCnt)
          lnkText = "Graph"
          lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=GraphCompTrend"
          lnkUrl = lnkUrl & "&PatientID=" & pat & "&CompTableKeyID=" & cmpFd & "&" & cmpSrcFd & "=" & compDat & "&" & cmpFd & "=" & compCmp
          navPop = "NAV"
          inout = "IN"
          fntSize = "10"
          fntColor = "#444488"
          bgColor = ""
          wdth = ""
          AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        End If
        response.write "</td>"
        response.write "</tr>"
      End If
      .MoveNext
    Loop
    If rowClmMax > 0 Then
      cnt = cnt + 1
      'Hdr
      response.write "<tr>"
      response.write "<td class=""cmpTdSty"" valign=""top"" colspan=""7"">"
      response.write "<table class=""cmpTdSty"" cellspacing=""0"" cellpadding=""3"" width=""100%"" style=""font-size:11pt""><tr>"
      arrRowHdr = Split(rowClmHdr, "||")
      ulRowHdr = UBound(arrRowHdr)
      For numRowHdr = 0 To ulRowHdr
        response.write "<td class=""cmpTdSty"" valign=""top"" align=""center""><b>" & arrRowHdr(numRowHdr) & "</b></td>"
      Next
      response.write "</tr>"
      'Row
      response.write "<tr>"
      For rowClmNum = 1 To rowClmMax
        response.write "<td class=""cmpTdSty"" valign=""top"" align=""center"">" & arrRowClm(1, rowClmNum) & "</td>"
      Next
      response.write "</tr></table></td>"
      response.write "<td>"
      If Not isTrd Then
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = "Trend"
        lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=DisplayCompTrend"
        lnkUrl = lnkUrl & "&PatientID=" & pat & "&CompTableKeyID=" & cmpFd & "&" & cmpSrcFd & "=" & compDat
        navPop = "NAV"
        inout = "IN"
        fntSize = "10"
        fntColor = "#444488"
        bgColor = ""
        wdth = ""
        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
      ElseIf isNum Then
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = "Graph"
        lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=GraphCompTrend"
        lnkUrl = lnkUrl & "&PatientID=" & pat & "&CompTableKeyID=" & cmpFd & "&" & cmpSrcFd & "=" & compDat
        navPop = "NAV"
        inout = "IN"
        fntSize = "10"
        fntColor = "#444488"
        bgColor = ""
        wdth = ""
        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
      End If
      response.write "</td>"
      response.write "</tr>"
    End If
    response.write "</table>"
    'ClosePanelSectionEMR
    response.write "</td>"
    response.write "</tr>"
  End If
  .Close
End With
Set rst = Nothing
End Sub
Function HasCompData(rlTb, rqTb, rqFd, vst, cmpSrcFd, cmpSrcKy, cmpTabFd, cmpTabKy, cmpFd, cmpKy)
  Dim rst, sql, ot
  Set rst = CreateObject("ADODB.Recordset")
  ot = False
  sql = "select " & rlTb & "." & rqFd
  sql = sql & " from " & rlTb & "," & rqTb
  sql = sql & " where " & rlTb & "." & rqFd & "=" & rqTb & "." & rqFd
  If Len(Trim(cmpKy)) > 0 Then
    sql = sql & " and " & rlTb & "." & cmpFd & "='" & cmpKy & "'"
  ElseIf Len(Trim(cmpTabKy)) > 0 Then
    sql = sql & " and " & rlTb & "." & cmpTabFd & "='" & cmpTabKy & "'"
  ElseIf Len(Trim(cmpSrcKy)) > 0 Then
    sql = sql & " and " & rlTb & "." & cmpSrcFd & "='" & cmpSrcKy & "'"
  End If
  sql = sql & " and " & rqTb & ".VisitationID='" & vst & "'"
  With rst
    .maxrecords = 1
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      ot = True
    End If
    .Close
  End With
  HasCompData = ot
Set rst = Nothing
End Function
Sub OpenPanelSection()
  Dim sz, dm, pFx
  sz = "20"
  pFx = "S004"
  dm = "width=""" & sz & """ height=""" & sz & """"
  response.write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
    response.write "<tr>"
      response.write "<td " & dm & " ><img " & dm & " src=""images/style/" & pFx & "11.gif""></td>"
      response.write "<td style=""background:url('images/style/" & pFx & "12.gif') 0 0 repeat""></td>"
      response.write "<td " & dm & "><img " & dm & " src=""images/style/" & pFx & "13.gif""></td>"
    response.write "</tr>"
    response.write "<tr>"
      response.write "<td style=""background:url('images/style/" & pFx & "21.gif') 0 0 repeat""></td>"
      response.write "<td width=""100%"">"
End Sub
Sub ClosePanelSection()
  Dim sz, dm, pFx
  sz = "20"
  pFx = "S004"
  dm = "width=""" & sz & """ height=""" & sz & """"
      response.write "</td>"
      response.write "<td style=""background:url('images/style/" & pFx & "23.gif') 0 0 repeat""></td>"
    response.write "</tr>"
    response.write "<tr>"
      response.write "<td " & dm & "><img " & dm & " src=""images/style/" & pFx & "31.gif""></td>"
      response.write "<td style=""background:url('images/style/" & pFx & "32.gif') 0 0 repeat""></td>"
      response.write "<td " & dm & "><img " & dm & " src=""images/style/" & pFx & "33.gif""></td>"
    response.write "</tr>"
  response.write "</table>"
End Sub
Sub OpenPanelSectionEMR()
  Dim sz, dm, pFx
  sz = "15"
  pFx = "S003"
  dm = "width=""" & sz & """ height=""" & sz & """"
  response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"">"
    response.write "<tr>"
      response.write "<td " & dm & " ><img " & dm & " src=""images/style/" & pFx & "11.gif""></td>"
      response.write "<td style=""background:url('images/style/" & pFx & "12.gif') 0 0 repeat""></td>"
      response.write "<td " & dm & "><img " & dm & " src=""images/style/" & pFx & "13.gif""></td>"
    response.write "</tr>"
    response.write "<tr>"
      response.write "<td style=""background:url('images/style/" & pFx & "21.gif') 0 0 repeat""></td>"
      response.write "<td align=""left"">"
End Sub
Sub ClosePanelSectionEMR()
  Dim sz, dm, pFx
  sz = "15"
  pFx = "S003"
  dm = "width=""" & sz & """ height=""" & sz & """"
      response.write "</td>"
      response.write "<td style=""background:url('images/style/" & pFx & "23.gif') 0 0 repeat""></td>"
    response.write "</tr>"
    response.write "<tr>"
      response.write "<td " & dm & "><img " & dm & " src=""images/style/" & pFx & "31.gif""></td>"
      response.write "<td style=""background:url('images/style/" & pFx & "32.gif') 0 0 repeat""></td>"
      response.write "<td " & dm & "><img " & dm & " src=""images/style/" & pFx & "33.gif""></td>"
    response.write "</tr>"
  response.write "</table>"
End Sub
Function HasPrintOutAccess(jb, prt)
  Dim rstTblSql, sql, ot
  ot = False
  Set rstTblSql = CreateObject("ADODB.Recordset")
  With rstTblSql
    sql = "select JobScheduleID from printoutalloc "
    sql = sql & " where printlayoutid='" & prt & "' and jobscheduleid='" & jb & "'"
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      ot = True
    End If
    .Close
  End With
  HasPrintOutAccess = ot
  Set rstTblSql = Nothing
End Function
Function HasModuleMgrAccess(jb, tb)
  Dim rstTblSql, sql, ot
  ot = ""
  Set rstTblSql = CreateObject("ADODB.Recordset")
  With rstTblSql
    sql = "select ModuleManagerID from ModuleManageralloc "
    sql = sql & " where tableid='" & tb & "' and jobscheduleid='" & jb & "' order by ModuleManagerID"
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      ot = .fields("ModuleManagerID")
    End If
    .Close
  End With
  HasModuleMgrAccess = ot
  Set rstTblSql = Nothing
End Function
Function GetActiveAdmission(vst)
  Dim rst, sql, ot
  ot = ""
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    sql = "select AdmissionID from Admission "
    sql = sql & " where visitationid='" & vst & "' and AdmissionStatusid='A001' order by AdmissionDate desc"
    ' sql = sql & " where visitationid='" & vst & "' and AdmissionStatusid='A001' "
    ' sql = sql & " And TransProcessVal2ID IN ('AdmissionPro-T001','AdmissionPro-T002') order by AdmissionDate desc"
    .maxrecords = 1
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      ot = .fields("AdmissionID")
    End If
    .Close
  End With
  GetActiveAdmission = ot
  Set rst = Nothing
End Function
Sub InitPrintLayout()
  Dim htStr
  LoadCSS

  htStr = ""
  htStr = htStr & "<script id=""scptPrintLayoutExtraScript"" LANGUAGE=""javascript"">"
  htStr = htStr & vbCrLf
  htStr = htStr & "function PLExtraScriptOnLoad(){" & vbCrLf
  htStr = htStr & "var upd;" & vbCrLf
  htStr = htStr & "HideEle(""trPrintControl"");" & vbCrLf

  htStr = htStr & "form1.onsubmit = HandlePrtFormOnSubmit;" & vbCrLf
  htStr = htStr & "upd=Helpers.trim(GetPageVariable('ItemsUpdated'));" & vbCrLf
  htStr = htStr & "if(Helpers.ucase(upd)=='YES'){" & vbCrLf
  htStr = htStr & "window.close();" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  htStr = htStr & "function HandlePrtFormOnSubmit(){" & vbCrLf
  htStr = htStr & "form1.action = processurl2(form1.action);" & vbCrLf
  htStr = htStr & "return true;" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  htStr = htStr & "</script>"
  response.write htStr

  SetPageVariable "AutoHidePrintControl", "1"
End Sub

Function IsDiagnosed(vst)
  Dim rst, sql, ot
  Set rst = CreateObject("ADODB.Recordset")
  ot = False
  sql = "select * from Diagnosis Where VisitationID='" & vst & "' "
  With rst
      rst.open qryPro.FltQry(sql), conn, 3, 4
      If rst.RecordCount > 0 Then
          ot = True
      End If
      rst.Close
  End With
  Set rst = Nothing
  IsDiagnosed = ot
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
