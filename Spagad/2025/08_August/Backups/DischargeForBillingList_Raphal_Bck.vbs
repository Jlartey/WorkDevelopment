'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim str, lnkCnt, wrd, jb, lstWhCls, lblNm
lnkCnt = 0

SetDischargeListWhCls
DisplayDischargeList
Sub DisplayDischargeList()
  Dim rst, rst2, rst3, sql, pat, adm, currWd, prtUrl, modMgr
  Dim st1, st2, st3, arr, ul, num, bTy, wd, wdSc, wrdNm, bedNm, bd, vst, clr
  Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, mdO
  Dim trnsVal
  
  Set rst = CreateObject("ADODB.Recordset")
  Set rst2 = CreateObject("ADODB.Recordset")
  Set rst3 = CreateObject("ADODB.Recordset")
  
  currWd = ""
  
  prtUrl = "wpgVisitation.asp?PageMode=ProcessSelect"
  If HasPrintOutAccess(jSchd, "PatientMedicalRecord") Then
      prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation"
  ElseIf Len(HasModuleMgrAccess(jSchd, "Visitation")) > 0 Then
      prtUrl = "wpgSelectModuleManager.asp?PositionForTableName=Visitation&PositionForCtxTableName=Visitation"
  ElseIf HasPrintOutAccess(jSchd, "VisitationRCP") Then
      prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation"
  End If
  response.Write "<meta http-equiv=""refresh"" content=""120"">"
  response.Write "<table cellspacing=""0"" cellpadding=""0"" border=""0"" width=""100%"">"
  
  response.Write "<tr>"
  response.Write "<td>"
  
  sql = "select a.visitationID,a.WardID,a.BedID,a.PatientID,v.MedicalOutComeID,a.AdmissionDate,a.AdmissionID, a.TransProcessValID "
  sql = sql & " from Admission as a,Visitation as v "
  sql = sql & " where v.visitationID=a.visitationID "
  'sql = sql & " and (v.MedicalOutComeID='M006' or v.MedicalOutComeID='M007')"
  sql = sql & " and a.AdmissionStatusID='A007' " & lstWhCls & " order by a.WardID,a.AdmissionID"
  
  response.Write "<table cellspacing=""0"" cellpadding=""3"" border=""1"" style=""font-size:11pt;border-collapse:collapse;border-color:#444444"" width=""100%"">"
    'Header
    cnt = 0
    response.Write "<tr bgcolor=""#eeeeee"" style=""font-size:14pt"">"
    response.Write "<td colspan=""11""><b>All Patients " & lblNm & " Awaiting Billing As At: <font color=""#6666cc""><u>" & FormatDateDetail(Now()) & "</u></font></b></td>"
    response.Write "</tr>"
    
  rst2.open qryPro.FltQry(sql), conn, 3, 4
  'SetPageMessages sql
  If rst2.recordCount > 0 Then
    Do While Not rst2.EOF
      cnt = cnt + 1
      pat = rst2.fields("PatientID")
      adm = rst2.fields("AdmissionID")
      wd = rst2.fields("WardID")
      bd = rst2.fields("BedID")
      vst = rst2.fields("VisitationID")
      mdO = rst2.fields("MedicalOutComeID")
      trnsVal = UCase(rst2.fields("TransProcessValID"))
      
      clr = "#ffffff"
      
      If UCase(mdO) = "M006" Then 'Request for billing
        clr = "#ffff44"
      ElseIf UCase(mdO) = "M007" Then 'Complete billing
        clr = "#88ff88"
      End If
      
      If UCase(trnsVal) = UCase("AdmissionPro-T003") Then
        clr = "#ffff44"
      ElseIf UCase(trnsVal) = UCase("AdmissionPro-T004") Then
        clr = "#88ff88"
      End If
      
      
      If UCase(wd) <> UCase(currWd) Then
        wrdNm = GetComboName("Ward", wd)
        response.Write "<tr bgcolor=""#eeeeee"">"
        response.Write "<td colspan=""11""><b>" & wrdNm & "</b></td>"
        response.Write "</tr>"

        response.Write "<tr bgcolor=""#eeeeee"">"
        response.Write "<td align=""center""><b>No</b></td>"
        response.Write "<td align=""center""><b>Folder No</b></td>"
        response.Write "<td align=""center""><b>Patient Name</b></td>"
        response.Write "<td align=""center""><b>Visit. No</b></td>"
        response.Write "<td align=""center""><b>Admission. No</b></td>"
        response.Write "<td align=""center""><b>Bed</b></td>"
        response.Write "<td align=""center""><b>Admit Date</b></td>"
        response.Write "<td align=""center"" colspan=""4""><b>Control</b></td>"
        response.Write "</tr>"
        currWd = wd
      End If
      
      lnkUrl = "wpgAdmission.asp?PageMode=ProcessSelect&AdmissionID=" & adm
      navPop = "POP"
      inout = "IN"
      fntSize = ""
      fntColor = "#222244"
      bgColor = clr
      wdth = ""
      
      response.Write "<tr bgcolor=""" & clr & """>"
      response.Write "<td align=""left"">" & CStr(cnt) & "</td>"
      
      response.Write "<td align=""left"">"
      'Clickable Url Link
      lnkCnt = lnkCnt + 1
      lnkID = "lnk" & CStr(lnkCnt)
      lnkText = pat
      AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
      response.Write "</td>"
      
      response.Write "<td align=""left"">"
      'Clickable Url Link
      lnkCnt = lnkCnt + 1
      lnkID = "lnk" & CStr(lnkCnt)
      lnkText = Replace(GetComboName("Patient", pat), " ", "&nbsp;")
      AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
      response.Write "</td>"
      
      response.Write "<td align=""left"">"
      'Clickable Url Link
      lnkCnt = lnkCnt + 1
      lnkID = "lnk" & CStr(lnkCnt)
      lnkText = vst
      AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
      response.Write "</td>"
      
      response.Write "<td align=""left"">"
      'Clickable Url Link
      lnkCnt = lnkCnt + 1
      lnkID = "lnk" & CStr(lnkCnt)
      lnkText = adm
      AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
      response.Write "</td>"
      
      response.Write "<td align=""left"">"
      'Clickable Url Link
      lnkCnt = lnkCnt + 1
      lnkID = "lnk" & CStr(lnkCnt)
      lnkText = Replace(GetComboName("Bed", bd), " ", "&nbsp;")
      AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
      response.Write "</td>"
      
      response.Write "<td align=""left"">"
      'Clickable Url Link
      lnkCnt = lnkCnt + 1
      lnkID = "lnk" & CStr(lnkCnt)
      lnkText = FormatDateDetail(rst2.fields("AdmissionDate"))
      AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
      response.Write "</td>"
      
      response.Write "<td align=""left"">"
      'Clickable Url Link
      lnkCnt = lnkCnt + 1
      lnkID = "lnk" & CStr(lnkCnt)
      lnkText = "<b>Bill</b>"
      'lnkUrl = "wpgSelectPrintLayout.asp?PositionForTableName=Admission&AdmissionID=" & adm
      lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=Admission1&PositionForTableName=Admission&AdmissionID=" & adm
      navPop = "POP"
      inout = "IN"
      fntSize = ""
      fntColor = "#ff0000"
      bgColor = clr
      wdth = ""
      AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
      response.Write "</td>"

      response.Write "<td align=""left"">"
      'Clickable Url Link
      lnkCnt = lnkCnt + 1
      lnkID = "lnk" & CStr(lnkCnt)
      lnkText = "<b>Folder</b>"
      lnkUrl = prtUrl & "&VisitationID=" & vst
      navPop = "POP"
      inout = "IN"
      fntSize = ""
      fntColor = "#ff0000"
      bgColor = clr
      wdth = ""
      AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
      response.Write "</td>"
      
      response.Write "<td align=""left"">"
      'Clickable Url Link
      lnkCnt = lnkCnt + 1
      lnkID = "lnk" & CStr(lnkCnt)
      lnkText = "<b>Process Bill</b>"
      lnkUrl = "wpgNavigateFrame.asp?FrameType=WorkFlow&PositionForTableName=Visitation&VisitationID=" & vst
      navPop = "POP"
      inout = "IN"
      fntSize = ""
      fntColor = "#ff0000"
      bgColor = clr
      wdth = ""
      AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
      response.Write "</td>"
      
      If IsBilling(jSchd) Then 'Not ward
'        response.write "<td align=""left"">"
'        'Clickable Url Link
'        lnkCnt = lnkCnt + 1
'        lnkID = "lnk" & CStr(lnkCnt)
'        lnkText = "<b>Complete Bill</b>"
'        lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=DischargeForBilling&PositionForTableName=WorkingDay&WorkingDayID=DAY20180515&VisitationID=" & vst
'        navPop = "POP"
'        inout = "IN"
'        fntSize = ""
'        fntColor = "#ff0000"
'        bgColor = clr
'        wdth = ""
'        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
'        response.write "</td>"
        
        If trnsVal = UCase("AdmissionPro-T003") Then 'discharge pending payment
            response.Write "<td align=""left"">"
            
            'Clickable Url Link
            lnkCnt = lnkCnt + 1
            lnkID = "lnk" & CStr(lnkCnt)
            lnkText = "<b>Complete Bill for Printing</b>"
            lnkUrl = "wpgAdmissionPro.asp?PageMode=AddNew&PullupData=AdmissionID||" & adm & "&TransProcessVal2ID=AdmissionPro-T004"
            navPop = "POP"
            inout = "IN"
            fntSize = ""
            fntColor = "#ff0000"
            bgColor = clr
            wdth = ""
            AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            'response.write "</td>"
            
            'Response.Write "<br>"
            
            'response.write "<td align=""left"">"
            'Clickable Url Link
            lnkCnt = lnkCnt + 1
            lnkID = "lnk" & CStr(lnkCnt)
            lnkText = "<b>Return to ward</b>"
            lnkUrl = "wpgAdmissionPro.asp?PageMode=AddNew&PullupData=AdmissionID||" & adm & "&TransProcessVal2ID=AdmissionPro-T006"
            navPop = "POP"
            inout = "IN"
            fntSize = ""
            fntColor = "#ff0000"
            bgColor = clr
            wdth = ""
            AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            
            response.Write "</td>"
        ElseIf trnsVal = UCase("AdmissionPro-T004") Then  'completed for payment/printing
            If gBillUtils.GetOutstandingBill(vst) <= 0 Or GetComboNameFld("Visitation", vst, "ReceiptTypeID") = "R002" Then
                response.Write "<td align=""left"">"
                'Clickable Url Link
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>Discharge Patient</b>"
                lnkUrl = "wpgAdmissionPro.asp?PageMode=AddNew&PullupData=AdmissionID||" & adm & "&TransProcessVal2ID=AdmissionPro-T005"
                navPop = "POP"
                inout = "IN"
                fntSize = ""
                fntColor = "#ff0000"
                bgColor = clr
                wdth = ""
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.Write "</td>"
            End If
            
        End If
        
      End If
      
      response.Write "</tr>"
      rst2.MoveNext
    Loop
    
  End If
  rst2.Close
  response.Write "</table>"
  response.Write "</td>"
  response.Write "</tr>"
  response.Write "</table>"
End Sub

Sub AddUrlLink(lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth)
  Dim plusMinus, imgName, lnkOpClNavPop, align
   plusMinus = ""
   imgName = ""
   align = ""
   lnkOpClNavPop = inout & "||" & navPop & "||800||600||CLOSE"
  AddPrtNavLink lnkID, plusMinus, imgName, lnkText, lnkUrl, lnkOpClNavPop, fntSize, fntColor, bgColor, align, wdth
End Sub
Function ExtractWeekDate(wks)
  Dim yr, pos, dys, dt, dt2, cnt, wks2, endWk, ot, wks3
  ot = ""
  If (Len(Trim(wks)) = 9) And (UCase(Left(Trim(wks), 3)) = "WKS") And IsNumeric(Right(Trim(wks), 6)) Then
    endWk = GetSystemVar("EndOfWeek")
    yr = Mid(Trim(wks), 4, 4)
    pos = Mid(Trim(wks), 8, 2)
    dys = CInt(pos) * 7
    dt = DateAdd("d", dys, CDate("1 Jan " & yr))
    cnt = 0
    wks2 = FormatWorkingWeek(dt, endWk)
    Do While (UCase(Trim(wks)) <> UCase(Trim(wks2))) And cnt < 8
      cnt = cnt + 1
      If wks > wks2 Then
        dt = DateAdd("d", dys + cnt, CDate("1 Jan " & yr))
      Else
        dt = DateAdd("d", dys - cnt, CDate("1 Jan " & yr))
      End If
      wks2 = FormatWorkingWeek(dt, endWk)
    Loop
    ot = dt
    'Move Date to beginning of week
    cnt = 0
    dt2 = DateAdd("d", 0 - cnt, dt)
    wks3 = FormatWorkingWeek(dt2, endWk)
    Do While (UCase(Trim(wks3)) = UCase(Trim(wks)))
      cnt = cnt + 1
      dt2 = DateAdd("d", 0 - cnt, dt)
      wks3 = FormatWorkingWeek(dt2, endWk)
      ot = dt2
    Loop
    ot = DateAdd("d", 1, ot)
  End If
  ExtractWeekDate = ot
End Function
Sub SetDischargeListWhCls()
  Dim jb, wdNm, outC, ot
  jb = jSchd
  wdNm = Trim(GetComboName("Ward", jb))
  If Len(wdNm) > 0 Then 'Ward Profile
    lstWhCls = " and a.WardID='" & jb & "'"
    lstWhCls = " and a.BranchID='" & brnch & "'"
    lblNm = " from [" & GetComboName("JobSchedule", jb) & "] "
  ElseIf IsBilling(jb) Then
    lstWhCls = " and a.BranchID='" & brnch & "'"
    lblNm = " from [" & GetComboName("Branch", brnch) & "] "
  End If
End Sub
Function IsBilling(jb)
  Dim ot, lst, arr, ul, num
  ot = False
  lst = "S17||S18||S87||S97||S17A||M13"
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
Function HasPrintOutAccess(jb, prt)
  Dim rstTblSql, sql, ot
  ot = False
  Set rstTblSql = CreateObject("ADODB.Recordset")
  With rstTblSql
    sql = "select JobScheduleID from printoutalloc "
    sql = sql & " where printlayoutid='" & prt & "' and jobscheduleid='" & jb & "'"
    .open qryPro.FltQry(sql), conn, 3, 4
    If .recordCount > 0 Then
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
    If .recordCount > 0 Then
      .MoveFirst
      ot = .fields("ModuleManagerID")
    End If
    .Close
  End With
  HasModuleMgrAccess = ot
  Set rstTblSql = Nothing
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
