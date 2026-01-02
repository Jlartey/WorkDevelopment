'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
response.write "<meta http-equiv=""refresh"" content=""120"">"
response.write Glob_GetBootstrap5()

Dim pat, patNm, dur, bDt, gen, genNm, sltHt1, sltHt2, dyHt, modMgr, cDt, wkDy, wkDyNm, cWkDy
Dim cnt, vDt1, vDt2, rst, pCnt, num, sql2, htStr, dyNm, prtUrl, patVdt, dCnt
Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, tb, tbKy, tbNm, mdDet, wdNm
Dim recKy, hasPrt, vst, spTyp, spTypNm, vDt, lnkCnt, nDt, nDt2, prevDys, ordByTyp, sql0, startDy, endDy, dispTyp
Dim sDt, eDt, cMth, mth, clDpt, patDet, patAg, patAgDet, patBdt, patDys, spTypDet, sp, spNm, md, mdNm, mdOutWhCls
Dim labDet, radDet, pharmDet, theaDet, currSp, otSummary, lstWhCls2

        Set rst = CreateObject("ADODB.Recordset")
        lnkCnt = 0
        prevDys = 365
        ordByTyp = Trim(Request("OrderByType"))
        dispTyp = Trim(Request("DisplayType"))
        currSp = Trim(Request("Specialist"))
        currMs = Trim(Request("MedicalService"))
        dur = Trim(Request.queryString("NoOfDays"))
        If Not IsNumeric(dur) Then
          dur = "0"
        End If
        If IsEmpty(currMs) Then
          currMs = "M001"
        End If

        LoadCSS
        InitPageScript
        SetListWhCls2

        prtUrl = "wpgVisitation.asp?PageMode=ProcessSelect"
        If HasPrintOutAccess(jSchd, "PatientMedicalRecord") Then
          prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation"
        ElseIf Len(HasModuleMgrAccess(jSchd, tb)) > 0 Then
          prtUrl = "wpgSelectModuleManager.asp?PositionForTableName=" & tb & "&PositionForCtxTableName=Visitation"
        ElseIf HasPrintOutAccess(jSchd, "VisitationRCP") Then
          prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation"
          sltTyp = "YES"
        End If

        cnt = 0
        cnt = cnt + 1
        nDt = Now()
        nDt2 = FormatDate(nDt)
        vDt = DateAdd("d", -1 * (CInt(dur)), nDt)

        vDt1 = FormatDate(vDt) & " 00:00:00" 'CDate("1 Mar 2016") '
        vDt2 = FormatDate(vDt) & " 23:59:59"
  
    pCnt = 0

    cMth = ""
    mth = ""
    cWkDy = ""
    pCnt = 0
    dCnt = 0

    
    response.write "<table class=""table table-striped cmpTdSty"" cellpadding=""2"" border=""1"" cellspacing=""0"" width=""100%"" style=""font-size:10pt"">"

    response.write "<tr><td align=""left"" width=""100%"" valign=""top"" colspan=""10"">"
      response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
       response.write "<tr><td colspan=""2"" align=""center"">"

        response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""font-size:12pt"" width=""100%"">"
            dTyp = GetDispType2(jSchd)
            If dTyp = UCase("LAB") Then
                response.write "<tr><td class=""cpHdrTd2"" style=""color:#048d04"">&emsp;<u>LABORATORY&nbsp;REQUESTS&nbsp;BY&nbsp;DOCTOR</u>&emsp;</td>"
            ElseIf dTyp = UCase("IMAGING") Then
                response.write "<tr><td class=""cpHdrTd2"" style=""color:#048d04"">&emsp;<u>CLINICAL&nbsp;IMAGING&nbsp;REQUESTS&nbsp;BY&nbsp;DOCTOR</u>&emsp;</td>"
            End If
         response.write "<td>&emsp;&emsp;<b>Day&nbsp;:</b></td>"
         
         response.write "<td>"
         SetInvestigationDays prevDys, nDt, nDt2, dur
         response.write "</td>"
         
         response.write "<td>&nbsp;&nbsp;</td>"
         response.write "<td><b>&nbsp;&nbsp;Type&nbsp;:</b></td>"

         response.write "<td>"
           SetMedicalService brnch, vDt1, vDt2, currMs
         response.write "</td>"

         response.write "<td>&nbsp;&nbsp;</td>"
         response.write "<td>&nbsp;&nbsp;</td>"

         response.write "</tr>"
         response.write "</table></td></tr>"

        response.write "<tr><td colspan=""2"" align=""left"">"
        response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""font-size:12pt"" width=""100%"">"
         response.write "<tr>"
        response.write "<td style=""color:#048d04"" class=""cpHdrTd2"">&nbsp;&nbsp;<u>As&nbsp;At&nbsp;:&nbsp;&nbsp;" & FormatDateDetail(Now()) & "</u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"

        lnkCnt = lnkCnt + 1
        lnkID = "trslt||lnk" & CStr(lnkCnt)
        response.write "<td onclick=""RefreshPage()"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
        response.write "<b>Refresh&nbsp;</b></td>"

        lnkCnt = lnkCnt + 1
        lnkID = "trslt||lnk" & CStr(lnkCnt)
        response.write "<td onclick=""cmdPrtBackOnClick()"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
        response.write "<b>&nbsp;<<&nbsp;</b></td>"

        lnkCnt = lnkCnt + 1
        lnkID = "trslt||lnk" & CStr(lnkCnt)
        response.write "<td onclick=""cmdPrintOnClick2()"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
        response.write "<b>&nbsp;Print&nbsp;</b></td>"

        lnkCnt = lnkCnt + 1
        lnkID = "trslt||lnk" & CStr(lnkCnt)
        response.write "<td onclick=""cmdPrtForwardOnClick()"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
        response.write "<b>&nbsp;>>&nbsp;</b></td>"

        lnkCnt = lnkCnt + 1
        lnkID = "trslt||lnk" & CStr(lnkCnt)
        response.write "<td>" '' " onclick=""OpenDialog()"" style=""color:#ff0000"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
        ' response.write "<button type='button' class='btn btn-warning' data-toggle='modal' data-target='#alerts-modal'>Notifications</button>"
        ' ' Response.write "<b>&emsp; My Alerts &emsp;</b></td>"
        response.write "</td>"

         response.write "<td><b>&nbsp;&nbsp;Arrange&nbsp;By&nbsp;:</b></td>"
                chkNm = " "
                chkTm = " checked "
                chkTm2 = "  "
                If UCase(ordByTyp) = UCase("Name") Then
                        chkNm = " checked "
                        chkTm = " "
                        chkTm2 = " "
                ElseIf UCase(ordByTyp) = UCase("Time2") Then
                        chkNm = " "
                        chkTm = " "
                        chkTm2 = " checked "
                End If
         ' response.write "<td><nobr>&emsp;<input type=""radio"" id=""inpOrderByType""  name=""inpOrderByType""  value=""Name"" onclick=""inpOrderByTypeOnClick()"" " & chkNm & ">Name</nobr></td>"
         ' response.write "<td><nobr>&emsp;<input type=""radio"" id=""inpOrderByType""  name=""inpOrderByType""  value=""Time"" onclick=""inpOrderByTypeOnClick()"" " & chkTm & ">Time</nobr></td>"
         response.write "<td><nobr>&emsp;<input type=""radio"" id=""inpOrderByType""  name=""inpOrderByType""  value=""Name"" onclick=""NoOfDaysOnchange()"" " & chkNm & ">Name</nobr></td>"
         response.write "<td><nobr>&emsp;<input type=""radio"" id=""inpOrderByType""  name=""inpOrderByType""  value=""Time"" onclick=""NoOfDaysOnchange()"" " & chkTm & ">Time (New First)</nobr></td>"
         response.write "<td><nobr>&emsp;<input type=""radio"" id=""inpOrderByType""  name=""inpOrderByType""  value=""Time2"" onclick=""NoOfDaysOnchange()"" " & chkTm2 & ">Time (Old First)</nobr></td>"
        response.write "</tr>"
        response.write "</table>"

       response.write "</td></tr>"
      response.write "</table>"
    response.write "</td></tr>"

  
        sql1 = "select distinct LabByDoctor.VisitationID,LabByDoctor.PatientID,LabByDoctor.SpecialistTypeID,Patient.PatientName,Visitation.PatientAge,Patient.BirthDate "
        sql1 = sql1 & " ,Visitation.MedicalServiceID,Visitation.MedicalOutComeID "
        sql1 = sql1 & " ,Gender.GenderName,SpecialistType.SpecialistTypeName from LabByDoctor,Patient,Visitation,Gender,SpecialistType "
        sql1 = sql1 & " where LabByDoctor.PatientID=Patient.PatientID  and LabByDoctor.BranchID='" & brnch & "' And Patient.PatientID=Visitation.PatientID "
        sql1 = sql1 & " And Gender.GenderID=Visitation.GenderID And SpecialistType.SpecialistTypeID=Visitation.SpecialistTypeID "
        sql1 = sql1 & lstWhCls2 & " And LabByDoctor.VisitationID=Visitation.VisitationID And LabByDoctor.PatientID=Visitation.PatientID And Visitation.InitialVisitationID<>'SUB' "
        sql1 = sql1 & " and LabByDoctor.PrescriptionDate between '" & vDt1 & "' and '" & vDt2 & "' "
        If Trim(currMs) <> "" Then
                sql1 = sql1 & " And Visitation.MedicalServiceID='" & currMs & "' "
        End If
        sql1 = sql1 & " order by Patient.PatientName"

        sql0 = "select distinct LabByDoctor.VisitationID,LabByDoctor.PatientID,LabByDoctor.SpecialistTypeID,Patient.PatientName,Visitation.PatientAge,Patient.BirthDate,LabByDoctor.PrescriptionDate "
        sql0 = sql0 & " ,Visitation.MedicalServiceID,Visitation.MedicalOutComeID "
        sql0 = sql0 & " ,Gender.GenderName,SpecialistType.SpecialistTypeName from LabByDoctor,Patient,Visitation,Gender,SpecialistType "
        sql0 = sql0 & " where LabByDoctor.PatientID=Patient.PatientID and LabByDoctor.BranchID='" & brnch & "' And Patient.PatientID=Visitation.PatientID "
        sql0 = sql0 & " And Gender.GenderID=Visitation.GenderID And SpecialistType.SpecialistTypeID=Visitation.SpecialistTypeID "
        sql0 = sql0 & lstWhCls2 & " And LabByDoctor.VisitationID=Visitation.VisitationID And LabByDoctor.PatientID=Visitation.PatientID And Visitation.InitialVisitationID<>'SUB' "
        sql0 = sql0 & " and LabByDoctor.PrescriptionDate between '" & vDt1 & "' and '" & vDt2 & "' "
        If Trim(currMs) <> "" Then
                sql0 = sql0 & " And Visitation.MedicalServiceID='" & currMs & "' "
        End If
        
        sql2 = sql0 & " order by LabByDoctor.PrescriptionDate desc"
        sql3 = sql0 & " order by LabByDoctor.PrescriptionDate"

        sql = sql2
        If UCase(ordByTyp) = UCase("Name") Then
                sql = sql1
        ElseIf UCase(ordByTyp) = UCase("Time2") Then
                sql = sql3
        End If
        vstOld = "-"

    With rst
                '.maxrecords = 50
                .open qryPro.FltQry(sql), conn, 3, 4
                If .RecordCount > 0 Then
                .MoveFirst
                wkDyNm = GetComboName("WorkingDay", FormatWorkingDay(vDt1))
            response.write "<tr style=""font-weight:bold;font-size:12pt"" bgcolor=""#eeeeee""><td colspan=""10"" align=""left"" valign=""top"">"
            response.write "<b>" & wkDyNm & "</b> ->&emsp; " & rst.RecordCount & " Patients with doctor's requests "
            response.write "</td></tr>"

            response.write "<tr style=""font-weight:bold;font-size:12pt"" bgcolor=""#eeeeee"">"
            response.write "<td valign=""top"" align=""center"">No.</td>"
            response.write "<td valign=""top"">Patient&nbsp;Details</td>"
            response.write "<td valign=""top"">Consult&nbsp;Details</td>"
            ' response.write "<td valign=""top"">Status</td>"
           'response.write "<td valign=""top"">Request By Doctor<span style='margin-left:23%'>Cost(Ghs)</span></td>"
           response.write "<td valign=""top"">Request By Doctor</td>"
            'response.write "<td valign=""top""><span style='margin-left:5%'>Requests Done</span></td>"
            response.write "<td valign=""top"">Requests Done</td>"
            response.write "<td valign=""top"">Summary</td>"
            ' response.write "<td valign=""top"">Theatre</td>"
            response.write "<td valign=""top"">Control</td>"
            response.write "</tr>"
                Do While Not .EOF
                        vDt = ""
                        patDet = ""
                        patBdt = .fields("BirthDate")
                        ' patAgDet = .fields("VisitInfo6")
                        patAgDet = Glob_FormatDateInterval(patBdt, Now())
                        spTypDet = ""
                        mdDet = GetComboName("MedicalService", .fields("MedicalServiceID"))
                        vst = .fields("VisitationID")
                        pat = .fields("PatientID")
                        patNm = .fields("PatientName")
                        patAg = .fields("PatientAge")
                        spTypNm = .fields("SpecialistTypeName")
                        genNm = .fields("GenderName")
                        md = .fields("MedicalOutComeID")

                        ''Exempt Repeating Records
                        If UCase(vst) <> UCase(vstOld) Then
                                pCnt = pCnt + 1

                                'ConsultDetail
                            spTypDet = spTypNm & "<br>" & vst & "<br>" & mdDet
                                spTypDet = Replace(spTypDet, " ", "&nbsp;")

                                        'PatientDetail
                                        ' patDet = "<b>" & patNm & "</b><br>" & patAgDet & "&nbsp;&nbsp;&nbsp;" & genNm & "&nbsp;&nbsp;&nbsp;No:&nbsp;" & pat & "&nbsp;&nbsp;&nbsp;No:&nbsp;" & ""
                                        patDet = "<b>" & patNm & "</b><br>No:&nbsp;" & pat & "<br>" & patAgDet & "&emsp;" & genNm & "&emsp;"
                                        patDet = patDet & "<br>" & UCase(GetComboName("Sponsor", GetComboNameFld("Visitation", vst, "SponsorID"))) & ""
                                        patDet = Replace(patDet, " ", "&nbsp;")
                                        
                                        response.write "<tr>"
                                        response.write "<td valign=""top"" align=""center"">" & CStr(pCnt) & "</td>"

                                        ' response.write "<td valign=""top"">" & patDet & "</td>"
                                        response.write "<td valign=""top"">"
                                        'Clickable Url Link
                                        lnkCnt = lnkCnt + 1
                                        lnkID = "lnk" & CStr(lnkCnt)
                                        lnkText = patDet
                                        lnkUrl = prtUrl & "&VisitationID=" & vst
                                        navPop = "OPEN"
                                        inout = "IN"
                                        fntSize = "10"
                                        fntColor = "#424242"
                                        bgColor = clr
                                        wdth = ""
                                        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                                        response.write "</td>"

                                        ' response.write "<td valign=""top"">" & spTypDet & "</td>"
                                        otSummary = ""
                                        DispPatientStatusInfo vst, md, spTypDet
                                        DisplayDoctorRequest vst
                                        DisplayTestRequests vst
                                        response.write "<td valign=""top"">" & otSummary & "</td>"
                                        response.write "<td valign=""top"">"
                                        'Clickable Url Link
                                        lnkCnt = lnkCnt + 1
                                        lnkID = "lnk" & CStr(lnkCnt)
                                        lnkText = Replace("Open<br>Folder", " ", "&nbsp;")
                                        lnkUrl = prtUrl & "&VisitationID=" & vst
                                        navPop = "OPEN"
                                        inout = "IN"
                                        fntSize = ""
                                        fntColor = "#8888ff"
                                        bgColor = clr
                                        wdth = ""
                                        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                                        response.write "</td>"
                                        response.write "</tr>"
                        End If
                        vstOld = vst
                                .MoveNext
                Loop
                End If
                .Close
    End With

    response.write "</table>"
        Set rst = Nothing

        response.flush
        SetWardAlerts
        ChangeFacilityHeader
        
Sub SetWardAlerts()
    Set rst = CreateObject("ADODB.Recordset")
    dtNow = Now()
    minsAgo = 45
    vDt1 = FormatDateDetail(DateAdd("n", (-1 * minsAgo), dtNow))
    vDt2 = FormatDateDetail(dtNow)
    sql = "select distinct ld.VisitationID,ld.PatientID,ld.SpecialistTypeID,p.PatientName, ld.PrescriptionDate "
    sql = sql & " from LabByDoctor ld,Patient p, Visitation v "
    sql = sql & " where ld.PatientID=p.PatientID And v.VisitationID=ld.VisitationID and p.PatientID=v.PatientID "
    If dTyp = UCase("LAB") Then
        sql = sql & "  and ld.TestGroupID='B13' "
    ElseIf dTyp = UCase("IMAGING") Then
        sql = sql & "  and ld.TestGroupID='B19' "
    End If
        sql = sql & " And v.InitialVisitationID<>'SUB' and v.MedicalServiceID='M003' " ''Inpatient
        sql = sql & " and ld.BranchID='" & brnch & "' and ld.PrescriptionDate between '" & vDt1 & "' and '" & vDt2 & "' "
        sql = sql & " order by ld.VisitationID, ld.PrescriptionDate asc "
        
        With rst
            rst.open qryPro.FltQry(sql), conn, 3, 4
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                vstOld = "-"
                        response.write Glob_GetBootstrapToastAlertHeader("")
                Do While Not rst.EOF
                        vst = rst.fields("VisitationID")
                        If UCase(vst) <> UCase(vstOld) Then
                                        Set tOption = server.CreateObject("Scripting.Dictionary")
                                        alertText = .fields("PatientName")
                            dt = rst.fields("PrescriptionDate")
                                        ' tOption.Add "close", False
                                        tOption.add "close", True
                                        tOption.add "icon", True
                                        tOption.add "delay", 60

                                        tOption.add "title", "New Requests From Ward"
                                        tOption.add "subtitle", Glob_GetHowLong(dt, dtNow)
                                        tOption.add "button1", "See Details"
                                        tOption.add "button1Url", prtUrl & "&VisitationID=" & .fields("VisitationID")
                                        ' tOption.Add "button2", "See User Details"
                                        ' tOption.Add "button2Url", lnkUrl
                                        lnkCnt = lnkCnt + 1
                            response.write Glob_GetBootstrapToastAlert("danger", alertText, tOption, lnkCnt)
                        End If
                        vstOld = vst
                    response.flush
                    Set tOption = Nothing
                    rst.MoveNext
                Loop
                        response.write Glob_GetBootstrapToastAlertFooter()
            End If
            rst.Close
        End With
        
        Set rst = Nothing
End Sub
        

Sub ChangeFacilityHeader()
  Dim js, facNm, appNm
  facNm = GetComboName("Branch", brnch) '' GetComboNameFld("SystemUser", uName, "BranchID"))
  appNm = "Medifem Multi-Specialist Hospital And Fertility Management System" ''Get from table later
  js = "<script>"
  js = js & " var elP = this.window.parent.document;"
  js = js & " if (elP) { "
  js = js & "  var el = elP.getElementById(""trmnuEle||41||14""); "
  js = js & "  if (el) { "
  '' js = js & " alert('We are here'); "
  js = js & "   if (el.cells && el.cells.length > 0) { "
  'js = js & "    el.cells[0].innerHTML = '<b>" & appNm & "</b><br>" & facNm & "'; "
  js = js & "    el.cells[0].innerHTML = '<b>" & appNm & "</b>; "
  js = js & "   } "
  js = js & "  } "
  js = js & " } "
  js = js & " ; "
  js = js & " ; "
  js = js & "</script>"

  response.write js
  
End Sub


Sub AddUrlLink(lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth)
  Dim plusMinus, imgName, lnkOpClNavPop, align
   plusMinus = ""
   imgName = ""
   align = ""
   lnkOpClNavPop = inout & "||" & navPop & "||800||600||CLOSE"
  AddPrtNavLink lnkID, plusMinus, imgName, lnkText, lnkUrl, lnkOpClNavPop, fntSize, fntColor, bgColor, align, wdth
End Sub

'ExtractDates
Sub ExtractDates(inFlt, outDt1, outDt2)
  Dim arr, ul, num, dat1, dat2
  dat1 = ""
  dat2 = ""
  arr = Split(inFlt, "||")
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
        dat1 = FormatDate(CDate(dat1)) & " 00:00:00"
      End If
    Else 'No Dat1
      If IsDate(dat2) Then
        dat1 = FormatDate(CDate(dat2)) & " 0:00:00"
        dat2 = FormatDate(CDate(dat2)) & " 23:59:59"
      Else 'No Dat2
      End If
    End If
  End If
  outDt1 = dat1
  outDt2 = dat2
End Sub

Function ExtractWorkingDate(wkDay)
    Dim str
    ExtractWorkingDate = Null
    str = Trim(wkDay)
    If Len(str) = 11 Then
      If UCase(Left(str, 3)) = "DAY" Then
        ExtractWorkingDate = CDate(Mid(str, 10, 2) & " " & MonthName(CInt(Mid(str, 8, 2)), 1) & " " & Mid(str, 4, 4))
      End If
    End If
End Function

Sub LoadCSS()
  Dim str
  str = ""
  str = str & "<style type='text/css' id=""styPrt"">"
  str = str & ".cpHdrTd{font-size:14pt;font-weight:bold}"
  str = str & ".cpHdrTr{background-color:#eeeeee}"
  str = str & ".cpHdrTd2{font-size:12pt;font-weight:bold}"
  str = str & ".cpHdrTr2{background-color:#eeeeee}" 'fafafa
  str = str & "</style>"
  response.write str
  
  response.write "<style>"
  response.write ".cmpTdSty {"
  response.write "border:1px solid #d0d0d0;"
  response.write "border-collapse: collapse;"
  response.write "}"
  response.write "</style>"
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

Sub DispPatientStatusInfo(vst, md, str)
  Dim rs, ot, sql, wd, aDt, dDt, aSt, aStNm, wdNm, clr
    Set rs = CreateObject("ADODB.Recordset")
    ot = ""
    sql = "select a.WardID,a.AdmissionDate,a.AdmissionStatusID,ads.AdmissionStatusName,a.DischargeDate,w.WardName "
    sql = sql & " ,v.MedicalOutComeID, v.MedicalServiceID, ms.MedicalServiceName, md.MedicalOutComeName "
    sql = sql & " from Admission as a,Ward as w,AdmissionStatus as ads, Visitation v, MedicalOutCome md, MedicalService ms"
    sql = sql & " where a.AdmissionStatusid=ads.AdmissionStatusid and a.wardid=w.wardid and a.visitationid='" & vst & "' "
    sql = sql & " And v.MedicalServiceID=ms.MedicalServiceID And v.MedicalOutComeID=md.MedicalOutComeID "
    sql = sql & " And v.VisitationID=a.VisitationID order by AdmissionDate desc"
    With rs
      .maxrecords = 1
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        wdNm = .fields("WardName")
        aStNm = .fields("AdmissionStatusName")
        aDt = .fields("AdmissionDate")
        dDt = .fields("DischargeDate")
        aSt = .fields("AdmissionStatusID")
        If IsDate(aDt) Then
          If IsDate(dDt) Then
            ot = wdNm & "&nbsp;&nbsp;&nbsp;" & DateDiff("d", aDt, dDt) & "Day(s)"
          Else
            ot = wdNm & "&nbsp;&nbsp;&nbsp;" & DateDiff("d", aDt, Now()) & "Day(s)"
          End If
          ot = ot & "<br>" & FormatDate(aDt) & "&nbsp;&nbsp;&nbsp;<b>Ward</b>:&nbsp;" & aStNm
        Else
          ot = wdNm & "<br>" & aStNm
        End If
      End If
      .Close
    End With
    'Status
    ' ot = str & "<br><b>Visit</b>:<br>" & GetComboName("MedicalOutCome", md)
    ot = str & "<br>"
    If Len(wdNm) > 0 Then
      ot = ot & "<b>" & wdNm & "</b><br>"
    End If
    ot = Replace(ot & GetComboName("MedicalOutCome", md), " ", "&nbsp;")
    
    clr = ""
    If (UCase(md) = "M000") Then 'Unseen
      clr = "bgcolor=""#ffcccc"""
    ElseIf (UCase(md) = "M005") Then 'Seen/Undischarged
      clr = "bgcolor=""#ffffdd"""
    Else
      clr = "bgcolor=""#ddffdd"""
    End If
    response.write "<td valign=""top"" " & clr & ">"
    'Clickable Url Link
    lnkCnt = lnkCnt + 1
    lnkID = "lnk" & CStr(lnkCnt)
    lnkText = ot
    lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&VisitationID=" & vst
    lnkUrl = lnkUrl & "&SectionType=DISCH"
    navPop = "OPEN"
    inout = "IN"
    fntSize = "10"
    fntColor = ""
    bgColor = clr
    wdth = ""
    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
    response.write "</td>"
    Set rs = Nothing
End Sub

Sub DisplayDoctorRequest(vst)
        Dim sql, rst, cnt, clr, dTyp
    Set rst = CreateObject("ADODB.Recordset")
    ot = ""
    dTyp = GetDispType2(jSchd)
    'Request
    rq = 0
    sql = "select d.*,d.unitcost as unitcost, l.LabTestName from LabByDoctor d, LabTest l Where d.LabTestID=l.LabTestID "
    If dTyp = UCase("LAB") Then
        sql = sql & "  and d.TestGroupID='B13' "
    ElseIf dTyp = UCase("IMAGING") Then
        sql = sql & "  and d.TestGroupID='B19' "
    End If
    sql = sql & " And (d.VisitationID='" & vst & "') order by d.PrescriptionDate desc "
    response.write "<td>"
    With rst
        dtNow = Now()
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
                rq = rst.RecordCount
            rst.MoveFirst
            cnt = 0
            Do While Not rst.EOF
                cnt = cnt + 1
                If cnt > 1 Then
                        response.write "<br>"
                End If
                dt = rst.fields("PrescriptionDate")
                ' min = (Right("00" & Minute(dt), 2))
                ' hr = Hour(dt)
                ' hr2 = "AM"
                ' If hr > 12 Then
                '       hr = hr - 12
                '       hr2 = " PM"
                ' End If
                ' hrs = (Right("00" & hr, 2))
                ' response.write "<nobr>" & cnt & ". @" & Replace(hrs & ":" & min & " " & hr2 & " -> " & rst.fields("LabTestName")," ", "&nbsp;") & "</nobr>"
                
                'response.write "<nobr>" & cnt & "." & Replace(" <em>[" & Glob_GetHowLong(dt, dtNow) & "]</em> " & rst.fields("LabTestName"), " ", "&nbsp;") & "</nobr>"
                 lbTest = rst.fields("LabTestID")
                instType = rst.fields("InsuranceTypeID")
                tststatus = getTestStatus(lbTest, instType)
              
                If (tststatus) Then
                response.write "<nobr style=''><span>" & cnt & "." & Replace(" <em>[" & Glob_GetHowLong(dt, dtNow) & "]</em> " & rst.fields("LabTestName"), " ", "&nbsp;") & "&nbsp </span><span style='background-color:#a4f9a4a8;padding:0 4px;border-radius:30%;font-size:12px'>&#8373;" & rst.fields("unitcost") & "</span></nobr>"
                Else
                 response.write "<nobr style='width:80%; justify-content: space-between;'><span>" & cnt & "." & Replace(" <em>[" & Glob_GetHowLong(dt, dtNow) & "]</em> " & rst.fields("LabTestName"), " ", "&nbsp;") & "&nbsp </span><span style='white-space:nowrap; background-color:#f977776b;padding:0 4px;border-radius:30%;font-size:12px'>&#8373;" & rst.fields("unitcost") & "</span></nobr>"
                End If
                rst.MoveNext
            Loop
        End If
        rst.Close
    End With
    otSummary = "<b>Doctor's Req.: " & rq & "</b>"
    response.write "</td>"
    Set rst = Nothing
End Sub

Function getTestStatus(lbTest, insType)
 Dim rs, ot, sql, sp
  Set rs = CreateObject("ADODB.Recordset")
  sql = "select permuteStatusID as permID from LabTestCostMatrix where LabTestID = '" & lbTest & "' AND InsuranceTypeID = '" & insType & "'"
   With rs
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
      sts = rs.fields("permID")
      If (sts = "P001") Then
      isTrue = True
      Else
      isTrue = False
      End If
      
      End If
      .Close
    End With
    getTestStatus = isTrue
End Function
    
    
Sub DisplayTestRequests(vst)
        Dim sql, rst, cnt, clr, dTyp
    Set rst = CreateObject("ADODB.Recordset")
    dTyp = GetDispType2(jSchd)
    'Done & Comp
    ot = 0
    rq = 0
    rq2 = 0
    rq3 = 0
    'Investigation
    sql = "select i.LabTestID, l.LabTestName, i.RequestDate, i.RequestDate1, i.RequestStatusID, rs.RequestStatusName, lt.LabTechName, 'Investigation' as tbl "
    sql = sql & " ,LabRequestID from Investigation i, LabTest l, RequestStatus rs, LabTech lt Where i.RequestStatusID=rs.RequestStatusID "
    sql = sql & " And (i.VisitationID='" & vst & "' Or i.VisitationID='" & vst & "-C') And l.LabTestID=i.LabTestID and lt.LabTechID=i.LabTechID "
    If dTyp = UCase("LAB") Then
        sql = sql & "  and i.TestGroupID='B13' "
    ElseIf dTyp = UCase("IMAGING") Then
        sql = sql & "  and i.TestGroupID='B19' "
    End If
    sql = sql & " UNION "
    sql = sql & " select i.LabTestID, l.LabTestName, i.RequestDate, i.RequestDate1, i.RequestStatusID, rs.RequestStatusName, lt.LabTechName, 'Investigation' as tbl "
    sql = sql & " ,LabRequestID from Investigation2 i, LabTest l, RequestStatus rs, LabTech lt Where i.RequestStatusID=rs.RequestStatusID "
    sql = sql & " And (i.VisitationID='" & vst & "' Or i.VisitationID='" & vst & "-C') And l.LabTestID=i.LabTestID and lt.LabTechID=i.LabTechID "
    If dTyp = UCase("LAB") Then
        sql = sql & "  and i.TestGroupID='B13' "
    ElseIf dTyp = UCase("IMAGING") Then
        sql = sql & "  and i.TestGroupID='B19' "
    End If
    If dTyp = UCase("LAB") Then
        sql = sql & "  and i.TestGroupID='B13' "
    ElseIf dTyp = UCase("IMAGING") Then
        sql = sql & "  and i.TestGroupID='B19' "
    End If
    sql = sql & " order by i.RequestStatusID asc,i.RequestDate asc "
        response.write "<td><table>"
    With rst
                .open qryPro.FltQry(sql), conn, 3, 4
                If .RecordCount > 0 Then
                        cnt = 0
                        ' response.write "<tr><td>No.</td><td>Test Desc.</td></tr>"
                        Do While Not .EOF
                                cnt = cnt + 1
                                navPop = "POP"
                                inout = "IN"
                                fntSize = "10"
                                fntColor = ""
                                bgColor = clr
                                wdth = ""
                                rsStat = Trim(.fields("RequestStatusID"))
                                reqId = Trim(.fields("LabRequestID"))
                                rsStatNm = Trim(.fields("RequestStatusName"))
                                pTb = Trim(.fields("tbl"))
                                clr = ""
                                lbTch = ""
                                rqDt = Now()
                                lb = .fields("LabTestID")
                                lbNm = .fields("LabTestName")
                                If UCase(rsStat) = UCase("RRD001") Then ''not ready
                                        rq = rq + 1
                                clr = "bgcolor=""#ffcccc"""
                                        rqDt = .fields("RequestDate")
                                ElseIf UCase(rsStat) = UCase("RRD003") Then ''partially ready
                                        rq3 = rq3 + 1
                                clr = "bgcolor=""#ffff88"""
                                        rqDt = .fields("RequestDate1")
                                        lbTch = .fields("LabTechName")
                                ElseIf UCase(rsStat) = UCase("RRD002") Then ''ready
                                        rq2 = rq2 + 1
                                clr = "bgcolor=""#88ff9f"""
                                        rqDt = .fields("RequestDate1")
                                        lbTch = .fields("LabTechName")
                                End If
                Min = (Right("00" & Minute(rqDt), 2))
                hr = Hour(rqDt)
                hr2 = "AM"
                If hr > 12 Then
                        hr = hr - 12
                        hr2 = " PM"
                End If
                hrs = (Right("00" & hr, 2))
                                lnkUrl = "wpg" & pTb & ".asp?PageMode=ProcessSelect&LabRequestID=" & reqId & "&LabTestID=" & lb
                                response.write "<tr>"
                                
                                ' response.write "<td valign=""top"" " & clr & ">"
                                ' 'Clickable Url Link
                                ' lnkCnt = lnkCnt + 1
                                ' lnkID = "lnk" & CStr(lnkCnt)
                                ' lnkText = Replace(cnt, " ", "&nbsp;")
                                ' AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                                ' response.write "</td>"

                                ' response.write "<td valign=""top"" " & clr & ">"
                                ' 'Clickable Url Link
                                ' lnkCnt = lnkCnt + 1
                                ' lnkID = "lnk" & CStr(lnkCnt)
                                ' lnkText = Replace(lb, " ", "&nbsp;")
                                ' AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                                ' response.write "</td>"

                                response.write "<td valign=""top"" " & clr & ">"
                                'Clickable Url Link
                                lnkCnt = lnkCnt + 1
                                lnkID = "lnk" & CStr(lnkCnt)
                                lnkText = Replace(lbNm, " ", "&nbsp;")
                                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                                response.write "</td>"

                                response.write "<td valign=""top"" " & clr & ">"
                                'Clickable Url Link
                                lnkCnt = lnkCnt + 1
                                lnkID = "lnk" & CStr(lnkCnt)
                                lnkText = Replace("<nobr>&emsp;@ " & hrs & ":" & Min & " " & LCase(hr2) & "</nobr>", " ", "&nbsp;")
                                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                                response.write "</td>"

                                ' response.write "<td valign=""top"" " & clr & ">"
                                ' 'Clickable Url Link
                                ' lnkCnt = lnkCnt + 1
                                ' lnkID = "lnk" & CStr(lnkCnt)
                                ' lnkText = Replace(rsStatNm, " ", "&nbsp;")
                                ' AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                                ' response.write "</td>"

                                ' response.write "<td valign=""top"" " & clr & ">"
                                ' 'Clickable Url Link
                                ' lnkCnt = lnkCnt + 1
                                ' lnkID = "lnk" & CStr(lnkCnt)
                                ' lnkText = Replace(lbTch, " ", "&nbsp;")
                                ' AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                                ' response.write "</td>"
                                response.write "</tr>"
                                .MoveNext
                        Loop
                End If
                .Close
    End With
        response.write "</table></td>"
    If dTyp = UCase("LAB") Then
            otSummary = Replace(otSummary & "<br>Results Ready: " & rq2 & "<br>Not Valildated: " & rq3 & "<br>No Results: " & rq, " ", "&nbsp;")
    ElseIf dTyp = UCase("IMAGING") Then
            otSummary = Replace(otSummary & "<br>Reports Ready: " & rq2 & "<br>Not Valildated: " & rq3 & "<br>No Report: " & rq, " ", "&nbsp;")
    End If
End Sub

Sub SetListWhCls2()
  Dim jb
  jb = Trim(jSchd)
  dispTyp = GetDispType2(jb)

  If dispTyp = "LAB" Then
    lstWhCls2 = " and (LabByDoctor.TestCategoryID='B13' Or LabByDoctor.TestGroupID='B13')"
  ElseIf dispTyp = "IMAGING" Then
    lstWhCls2 = " and (LabByDoctor.TestCategoryID='B19' Or LabByDoctor.TestGroupID='B19')"
  End If
  lstWhCls2 = lstWhCls2 & " And LabByDoctor.WorkingDayID >= 'DAY20220501' "
End Sub


Function GetDispType2(jb)
  Dim ot
  ot = ""
  If UCase(Left(jb, 3)) = "M07" Then
    ot = "IMAGING"
  ElseIf UCase(Left(jb, 3)) = "M05" Then
    ot = "LAB"
  End If
  GetDispType2 = ot
End Function

Sub DispRadDetail(vst)
  Dim sql, ot, rq, don, comp, tot, rst, cnt, clr
    Set rs = CreateObject("ADODB.Recordset")
    ot = ""
    'Request
    rq = 0
    sql = "select count(PatientID) as cnt from LabByDoctor "
    sql = sql & " where VisitationID='" & vst & "' and TestGroupID='B19'"
    rq = GetVisitInfoCount(sql)
    'Done & Comp
    don = 0
    comp = 0
    'Investigation
    sql = "select count(PatientID) as cnt,RequestStatusID from Investigation "
    sql = sql & " where VisitationID='" & vst & "' and TestGroupID='B19' group by RequestStatusID"
    With rs
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        Do While Not .EOF
          If Not IsNull(.fields("cnt")) Then
            If IsNumeric(.fields("cnt")) Then
              rst = Trim(.fields("RequestStatusID"))
              cnt = .fields("cnt")
              don = don + cnt
              If UCase(rst) = "RRD002" Then 'Completed
                comp = comp + cnt
              End If
            End If
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
    'Investigation2
    sql = "select count(PatientID) as cnt,RequestStatusID from Investigation2 "
    sql = sql & " where VisitationID='" & vst & "' and TestGroupID='B19' group by RequestStatusID"
    With rs
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        Do While Not .EOF
          If Not IsNull(.fields("cnt")) Then
            If IsNumeric(.fields("cnt")) Then
              rst = Trim(.fields("RequestStatusID"))
              cnt = .fields("cnt")
              don = don + cnt
              If UCase(rst) = "RRD002" Then 'Completed
                comp = comp + cnt
              End If
            End If
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
    tot = rq + don + comp
    clr = ""
    If tot > 0 Then
      If (don >= rq) And (comp = don) Then
        clr = "bgcolor=""#ddffdd"""
      ElseIf don > 0 Then
        clr = "bgcolor=""#ffffdd"""
      Else
        clr = "bgcolor=""#ffcccc"""
      End If
      response.write "<td valign=""top"" " & clr & ">"
      'Clickable Url Link
      lnkCnt = lnkCnt + 1
      lnkID = "lnk" & CStr(lnkCnt)
      lnkText = Replace("Requests : " & CStr(rq) & "<br>      Done : " & CStr(don) & "<br>    Ready : " & CStr(comp), " ", "&nbsp;")
      lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&VisitationID=" & vst
      lnkUrl = lnkUrl & "&SectionType=RAD"
      navPop = "POP"
      inout = "IN"
      fntSize = "10"
      fntColor = ""
      bgColor = clr
      wdth = ""
      AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
    Else
      response.write "<td valign=""top"">"
    End If
    response.write "</td>"
End Sub

Sub SetInvestigationDays(prevDys, nDt, nDt2, dur)
    dyHt = "<select size=""1"" name=""NoOfDays"" id=""NoOfDays"" onchange=""NoOfDaysOnchange()"">"
    dyHt = dyHt & "<option value=""""></option>"
    
    sDt = FormatDate(DateAdd("d", -1 * (prevDys), nDt)) & " 00:00:00"
    eDt = FormatDate(nDt) & " 23:59:59"
    cMth = ""
    mth = ""
    ' sql0 = "select distinct WorkingDayID from LabByDoctor where LabByDoctor.BranchID='" & brnch & "' and PrescriptionDate between '" & sDt & "' and '" & eDt & "' order by WorkingDayID desc"
    sql0 = "select distinct WorkingDayID from LabByDoctor where LabByDoctor.BranchID='" & brnch & "' "
    sql0 = sql0 & lstWhCls2 & " and PrescriptionDate between '" & sDt & "' and '" & eDt & "' order by WorkingDayID desc"
    With rst
      .open qryPro.FltQry(sql0), conn, 3, 4
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          wkDy = Trim(.fields("WorkingDayID"))
          If Len(wkDy) = 11 Then
            If UCase(Left(wkDy, 3)) = "DAY" Then
              cDt = FormatDate(ExtractWorkingDate(wkDy))
              mth = Trim(FormatWorkingMonth(cDt))
              num = DateDiff("d", cDt, nDt2)
              dyNm = Right(CStr(num + 1000), 3) & " Days ago"
              If num = 0 Then
                dyNm = "Today      "
              ElseIf num = 1 Then
                dyNm = "Yesterday "
              End If
              dyNm = dyNm & " ->" & GetComboName("WorkingDay", wkDy)
              If UCase(cMth) <> UCase(mth) Then
                 dyHt = dyHt & "<optGroup label=""" & GetComboName("Workingmonth", mth) & """>"
                 cMth = mth
              End If
              If UCase(CStr(num)) = UCase(dur) Then
                dyHt = dyHt & "<option value=""" & CStr(num) & """ selected>" & dyNm & "</option>"
              Else
                 dyHt = dyHt & "<option value=""" & CStr(num) & """>" & dyNm & "</option>"
              End If
            End If
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
    dyHt = dyHt & "</select>"
    response.write dyHt
End Sub

Sub SetMedicalService(br, vDt1, vDt2, currMs)
  Dim rs, ot, sql, sp
  Set rs = CreateObject("ADODB.Recordset")
  sql = "select distinct MedicalServiceID from Visitation where BranchID='" & br & "' and VisitDate between '" & vDt1 & "' and '" & vDt2 & "'"
  sql = sql & " And MedicalServiceID IN ('M001','M003') order by MedicalServiceID"

  ot = "<select id=""MedicalService"" name=""MedicalService"" onchange=""MedicalServiceOnchange()"">"
  ot = ot & "<option></option>"
  With rs
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        sp = .fields("MedicalServiceID")
        If UCase(sp) = UCase(currMs) Then
          ot = ot & "<option value=""" & sp & """ selected>" & GetComboName("MedicalService", sp) & "</option>"
        Else
          ot = ot & "<option value=""" & sp & """>" & GetComboName("MedicalService", sp) & "</option>"
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  ot = ot & "</select>"
  response.write ot
  Set rs = Nothing
End Sub


Sub SetDoctorDays(prevDys, nDt, nDt2, uNm, dur)
  Dim sDt, eDt, ot, dyHt, sql0, wkDy, cDt, mth, num, dyNm, cMth, rst
  Set rst = CreateObject("ADODB.Recordset")
  cMth = ""
  mth = ""
  dyHt = "<select name=""NoOfDays"" id=""NoOfDays"" onchange=""NoOfDaysOnchange()"">"
  dyHt = dyHt & "<option value=""""></option>"
  sDt = FormatDate(DateAdd("d", -1 * (prevDys), nDt)) & " 00:00:00"
  eDt = FormatDate(nDt) & " 23:59:59"
  ' sql0 = "select distinct WorkingDayID from Visitation where SpecialistID='" & uNm & "' and VisitDate between '" & sDt & "' and '" & eDt & "' order by WorkingDayID desc"
  sql0 = "select distinct WorkingDayID from Visitation where VisitDate between '" & sDt & "' and '" & eDt & "' order by WorkingDayID desc"
  With rst
    .open qryPro.FltQry(sql0), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        wkDy = Trim(.fields("WorkingDayID"))
        If Len(wkDy) = 11 Then
          If UCase(Left(wkDy, 3)) = "DAY" Then
            cDt = FormatDate(ExtractWorkingDate(wkDy))
            mth = Trim(FormatWorkingMonth(cDt))
            num = DateDiff("d", cDt, nDt2)
            dyNm = Right(CStr(num + 1000), 3) & " Days ago"
            If num = 0 Then
              dyNm = "Today      "
            ElseIf num = 1 Then
              dyNm = "Yesterday "
            End If
            dyNm = dyNm & " ->" & GetComboName("WorkingDay", wkDy)
            If UCase(cMth) <> UCase(mth) Then
               'dyHt = dyHt & "<optGroup label=""" & GetComboName("Workingmonth", mth) & """>"
               cMth = mth
            End If
            If UCase(CStr(num)) = UCase(dur) Then
              dyHt = dyHt & "<option value=""" & CStr(num) & """ selected>" & dyNm & "</option>"
            Else
               dyHt = dyHt & "<option value=""" & CStr(num) & """>" & dyNm & "</option>"
            End If
          End If
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  dyHt = dyHt & "</select>"
  response.write dyHt
  Set rst = Nothing
End Sub

Sub LoadCSS()
  Dim str
  str = ""
  str = str & "<style type='text/css' id=""styPrt"">"
  str = str & ".cpHdrTd{font-size:14pt;font-weight:bold}"
  str = str & ".cpHdrTr{background-color:#eeeeee}"
  str = str & ".cpHdrTd2{font-size:12pt;font-weight:bold}"
  str = str & ".cpHdrTr2{background-color:#eeeeee}" 'fafafa
  str = str & "</style>"
  response.write str

  response.write "<style>"
  response.write ".cmpTdSty {"
  response.write "border:1px solid #d0d0d0;"
  response.write "border-collapse: collapse;"
  response.write "}"
  response.write "</style>"
End Sub

Sub InitPageScript()
  Dim htStr
  'Client Script
  htStr = ""
  htStr = htStr & "<script id=""scptPrintLayoutExtraScript"" LANGUAGE=""javascript"">" & vbCrLf
  htStr = htStr & vbCrLf
  htStr = htStr & "function PLExtraScriptOnLoad(){" & vbCrLf
  htStr = htStr & "window.onresize=windowOnresize;" & vbCrLf
  htStr = htStr & "HideEle(""trPrintControl"");" & vbCrLf
  htStr = htStr & "windowOnresize();" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  htStr = htStr & "function windowOnresize(){" & vbCrLf
  htStr = htStr & " var ht,ele;" & vbCrLf
  htStr = htStr & " ht=window.innerHeight;" & vbCrLf
  htStr = htStr & " if (Helpers.isnumeric(ht)){" & vbCrLf
  htStr = htStr & "ele = document.getElementById('iFrm1');" & vbCrLf

  If UCase(fullScrn) = "NO" Then 'No Full Screen
  htStr = htStr & "if (ele) {" & vbCrLf
  htStr = htStr & " ele.height=Helpers.cstr(Helpers.cint(ht)-80);" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "ele = document.getElementById('iFrm2');" & vbCrLf
  htStr = htStr & "if (ele) {" & vbCrLf
  htStr = htStr & " ele.height=Helpers.cstr(Helpers.cint(ht)-90);" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  Else
  htStr = htStr & "if (ele) {" & vbCrLf
  htStr = htStr & " ele.height=Helpers.cstr(Helpers.cint(ht));" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "ele = document.getElementById('iFrm2');" & vbCrLf
  htStr = htStr & "if (ele) {" & vbCrLf
  htStr = htStr & " ele.height=Helpers.cstr(Helpers.cint(ht));" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  End If
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  'RefreshPage()
  htStr = htStr & "function RefreshPage(){" & vbCrLf
  htStr = htStr & "window.location.reload();" & vbCrLf
  htStr = htStr & "}" & vbCrLf


  'NoOfDays()
  htStr = htStr & "function NoOfDaysOnchange(){" & vbCrLf
  htStr = htStr & "var ur,dy,sp,ordByTyp,ms;" & vbCrLf
  htStr = htStr & "dy=GetEleVal('NoOfDays');" & vbCrLf
  htStr = htStr & "sp=GetEleVal('Specialist');" & vbCrLf
  htStr = htStr & "ms=GetEleVal('MedicalService');" & vbCrLf
  htStr = htStr & "ordByTyp=GetCheckedRadio('inpOrderByType');" & vbCrLf
  htStr = htStr & "ur='wpgPrtPrintLayoutAll.asp?PrintLayoutName=MonitorLaboratory&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "ur=ur + '&WorkingDayID=DAY20160401&NoOfDays=' + dy + '&OrderByType=' + ordByTyp + '&Specialist=' + sp + '&MedicalService=' + ms;" & vbCrLf
  htStr = htStr & "window.location.href=processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  'Specialist()
  htStr = htStr & "function SpecialistOnchange(){" & vbCrLf
  htStr = htStr & "var ur,dy,sp,ordByTyp,ms;" & vbCrLf
  htStr = htStr & "dy=GetEleVal('NoOfDays');" & vbCrLf
  htStr = htStr & "sp=GetEleVal('Specialist');" & vbCrLf
  htStr = htStr & "ms=GetEleVal('MedicalService');" & vbCrLf
  htStr = htStr & "ordByTyp=GetCheckedRadio('inpOrderByType');" & vbCrLf
  htStr = htStr & "ur='wpgPrtPrintLayoutAll.asp?PrintLayoutName=MonitorLaboratory&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "ur=ur + '&WorkingDayID=DAY20160401&NoOfDays=' + dy + '&OrderByType=' + ordByTyp + '&Specialist=' + sp + '&MedicalService=' + ms;" & vbCrLf
  htStr = htStr & "window.location.href=processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  'MedicalService()
  htStr = htStr & "function MedicalServiceOnchange(){" & vbCrLf
  htStr = htStr & "var ur,dy,sp,ordByTyp,ms;" & vbCrLf
  htStr = htStr & "dy=GetEleVal('NoOfDays');" & vbCrLf
  htStr = htStr & "sp=GetEleVal('Specialist');" & vbCrLf
  htStr = htStr & "ms=GetEleVal('MedicalService');" & vbCrLf
  htStr = htStr & "ordByTyp=GetCheckedRadio('inpOrderByType');" & vbCrLf
  htStr = htStr & "ur='wpgPrtPrintLayoutAll.asp?PrintLayoutName=MonitorLaboratory&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "ur=ur + '&WorkingDayID=DAY20160401&NoOfDays=' + dy + '&OrderByType=' + ordByTyp + '&Specialist=' + sp + '&MedicalService=' + ms;" & vbCrLf
  htStr = htStr & "window.location.href=processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  
  'GetCheckedRadio()
  htStr = htStr & "function GetCheckedRadio(inp){" & vbCrLf
  htStr = htStr & "var ele,lth,n,ot;" & vbCrLf
  htStr = htStr & "ot='';" & vbCrLf
  htStr = htStr & "ele=document.getElementsByName(inp);" & vbCrLf
  htStr = htStr & "lth=ele.length;" & vbCrLf
  htStr = htStr & "if (lth>0){" & vbCrLf
  htStr = htStr & "for(n=0;n<lth;n++){" & vbCrLf
  htStr = htStr & "if (ele[n].checked){" & vbCrLf
  htStr = htStr & "ot=ele[n].value;" & vbCrLf
  htStr = htStr & "break;" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "return ot;" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  'StartDayOnChange()
  'EndDayOnChange()
  'DisplayTypeOnChange()


  htStr = htStr & "function formatwinposprt(wd, ht) {" & vbCrLf
  htStr = htStr & "var lft, tp;" & vbCrLf
  htStr = htStr & "var ot;" & vbCrLf
  htStr = htStr & "lft = Helpers.cstr((screen.availWidth - Helpers.cint(wd)) / 2);" & vbCrLf
  htStr = htStr & "tp = Helpers.cstr((screen.availHeight - Helpers.cint(ht)) / 2);" & vbCrLf
  htStr = htStr & "if (Helpers.cint(lft)<0){" & vbCrLf
  htStr = htStr & "lft=""0""" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "if (Helpers.cint(tp)<0){" & vbCrLf
  htStr = htStr & "tp=""0""" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "ot = ""top="" + tp + "",left="" + lft + "",height="" + ht + "",width="" + wd + "",status=no,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes"";" & vbCrLf
  htStr = htStr & "return ot;" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  htStr = htStr & "</script>"
  response.write htStr

  ' js = "<link rel='stylesheet' type='text/css' href='CSS/bootstrap.min.css'> " & vbCrLf
  ' js = js & "<script src='Scripts/jquery-3.3.1.js' language='javascript'></script> " & vbCrLf
  ' js = js & "<script src='Scripts/bootstrap.min.js' language='javascript'></script> " & vbCrLf
  js = js & "<script>" & vbCrLf
  js = js & "  " & vbCrLf
  js = js & "  " & vbCrLf
  js = js & "</script>"
  response.write js
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
