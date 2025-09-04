'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim lnkCnt, wdth

Call ProcessPage

Sub ProcessPage()
    Dim recKy
    recKy = Trim(Request.querystring("VisitationID"))

    If Not HasBedAssigned(recKy) Then
        Call LoadCSS
        Call ShowPage
    Else
        url = GetPageURL()
        response.Clear
        response.redirect url
    End If

End Sub
Function GetPageURL()
    Dim prtUrl, tb, tbKy, tbNm, vst
    tb = "Visitation"
    tbKy = "VisitationID"
    prtUrl = "wpgVisitation.asp?PageMode=ProcessSelect"
    vst = Trim(Request("VisitationID"))
    'modMgr = HasModuleMgrAccess(jSchd, tb)
    If HasPrintOutAccess(jschd, "PatientMedicalRecord") Then
        prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation&VisitationID=" & vst & "&FullScreen=Yes"
    ElseIf Len(HasModuleMgrAccess(jschd, tb)) > 0 Then
        prtUrl = "wpgSelectModuleManager.asp?PositionForTableName=" & tb & "&PositionForCtxTableName=Visitation&VisitationID=" & vst & "&FullScreen=Yes"
    ElseIf HasPrintOutAccess(jschd, "VisitationRCP") Then
        prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation&VisitationID=" & vst & "&FullScreen=Yes"
    End If
    GetPageURL = prtUrl
End Function
Function HasBedAssigned(ky)
    Dim sql, rst, waitBedNo, bedNo, ot

    ot = True
    sql = "select top 1 * from Admission where Visitationid='" & ky & "'"
    sql = sql & " order by AdmissionDate desc "

    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        waitBedNo = GetWaitingBedNo(rst.fields("WardID"))
        ot = (UCase(waitBedNo) <> UCase(rst.fields("BedID"))) 'Not on Waiting List
    End If
    HasBedAssigned = ot
End Function
Sub ShowPage()
    Dim tb, tbKy, tbNm, sTb, sTbKy, sTbNm, pTbKy, recKy, pRecKy
    tb = "Visitation"
    tbKy = "VisitationID"
    tbNm = "Visit"
    sTb = ""
    sTbKy = ""
    sTbNm = ""
    pTbKy = "PatientID"
    recKy = Trim(Request.querystring(tbKy))
    pRecKy = GetComboNameFld("Visitation", recKy, "PatientID")

    SetPageVariable "AutoHidePrintControl", "Yes"
    response.write "<table>"
    response.write "<tbody>"
        response.write "<tr><td  colspan='100' style='font-size:20px;font-weight:bold;text-align:center;'>" & GetComboName("Patient", pRecKy) & "</td></tr>"
        response.write "<tr><td colspan='100' style='color:red;padding:10px;font-size:18px;font-weight:bold;'>This patient is on the waiting list. Please Assign a bed to this patient to continue.</td></tr>"
        Call AddSecAdmitDetail(tb, tbKy, tbNm, pTbKy, pRecKy, recKy)
        response.write "<tr><td colspan='100' style='text-align:center;padding:10px;'><span style='color:green;font-size:18px;font-weight:bold;text-decoration:underline;cursor:pointer;' onclick='window.location.reload();'>Click here to open folder when done.</span></td></tr>"
    response.write "</tbody>"
    response.write "</table>"
End Sub
Function GetWaitingBedNo(wrd)
    Dim rst, sql, ot
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select BedID from Bed where Wardid='" & wrd & "' and BedNoID='000'"
    ot = ""
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .movefirst
            ot = .fields("BedID")
        End If
        .Close
    End With
    GetWaitingBedNo = ot
    Set rst = Nothing
End Function
Sub AddSecAdmitDetail(tb, tbKy, tbNm, pTbKy, pRecKy, recKy)
    Dim ag, sTb, sTbKy, sTbNm, pat, hasAcc, wd, ot, imgSrc, cmpUrl
    Dim lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, prtTb, modMgr, prt
    wd = "100%"
    'Admission
    sTb = "Admission"
    sTbKy = "AdmissionID"
    sTbNm = "Admission"
    If HasAccessRight(uName, "frm" & sTb, "View") Then
        AddPrtTabItem prtGrpTab, sTb, sTbNm
        response.write "<tr id=""trTbSc" & sTb & """>"
        response.write "<td>"
        'OpenPanelSection
        response.write "<table width=""100%"" border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse:collapse; border-color:#eeeeee"">"
            'Header
            currSect = "AdmitDetail"
            response.write "<tr>"
            response.write "<td>"
                response.write "<table width=""" & wd & """ border=""0"" cellpadding=""0"" cellspacing=""0"">"
                response.write "<tr class=""cpHdrTr"">"
                    response.write "<td><table border=""0"" cellpadding=""2"" cellspacing=""0""><tr>"
                    response.write "<td class=""cpHdrTd"">" & sTbNm & "</td>"
                    If HasActiveAdmission(recKy) Then
                        response.write "<td>"
                        'Clickable Url Link
                        lnkCnt = lnkCnt + 1
                        lnkID = "lnk" & CStr(lnkCnt)
                        lnkText = "<b>&nbsp;&nbsp;Discharge Pending Billing</b>"
                        lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=DischargeForBilling&PositionForTableName=WorkingDay&WorkingDayID=DAY20180514"
                        lnkUrl = lnkUrl & "&VisitationID=" & recKy
                        navPop = "POP"
                        inOut = "IN"
                        fntSize = ""
                        fntColor = "#44aa44"
                        bgColor = ""
                        wdth = ""
                        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
                        response.write "</td>"
                    End If
                    response.write "<td>"
                    'Clickable Url Link
                    lnkCnt = lnkCnt + 1
                    lnkID = "lnk" & CStr(lnkCnt)
                    lnkText = "<b>&nbsp;&nbsp;Discharge&nbsp;Patient&nbsp;/&nbsp;Close&nbsp;Visit</b>"
                    lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=EMRSelector&EMRDataID=TH080&EMRCompTabID=&CompTableKeyID=EMRComponentID&VisitationID=" & recKy
                    navPop = "POP"
                    inOut = "IN"
                    fntSize = ""
                    fntColor = "#444488"
                    bgColor = ""
                    wdth = ""
                    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
                    response.write "</td>"

                    If HasAccessRight(uName, "frm" & sTb, "FormSearch") Then
    '                    Response.Write "<td>"
    '                    'Clickable Url Link
    '                    lnkCnt = lnkCnt + 1
    '                    lnkID = "lnk" & CStr(lnkCnt)
                        lnkText = "<b>&nbsp;&nbsp;Search</b>"
                        lnkUrl = "wpgSrh" & sTb & ".asp"
                        navPop = "POP"
                        inOut = "IN"
                        fntSize = ""
                        fntColor = "#444488"
                        bgColor = ""
                        wdth = ""
                        lnkCnt = lnkCnt + 1
                        lnkID = "trslt||lnk" & CStr(lnkCnt)
                        response.write "<td style=""color:" & fntColor & """ onclick=""AddInternalIFrame('" & currSect & "','" & lnkUrl & "','500','100%','')"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
                        response.write lnkText & "</td>"
                    End If
                    response.write "</tr></table></td>"
                response.write "</tr>"
                response.write "</table>"
            response.write "</td>"
            response.write "</tr>"
            InsertIFrameInTRTag currSect, ""
            'Detail
            AddAdmission sTb, sTbKy, recKy
        response.write "</table>" 'Referral
        'ClosePanelSection
        response.write "</td>"
        response.write "</tr>"
    End If
    response.Flush
End Sub
Sub AddUrlLink(lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth)
    Dim plusMinus, imgName, lnkOpClNavPop, align
    plusMinus = ""
    imgName = ""
    align = ""
    lnkOpClNavPop = inOut & "||" & navPop & "||800||600||CLOSE"
    AddPrtNavLink lnkID, plusMinus, imgName, lnkText, lnkUrl, lnkOpClNavPop, fntSize, fntColor, bgColor, align, wdth
End Sub
Function HasActiveAdmission(vst)
    Dim rst, sql, ot
    ot = False
    Set rst = CreateObject("ADODB.Recordset")
    With rst
    sql = "select AdmissionID from Admission "
    sql = sql & " where visitationid='" & vst & "' and AdmissionStatusid='A001'"
    .maxrecords = 1
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
        ot = True
    End If
    .Close
    End With
    HasActiveAdmission = ot
    Set rst = Nothing
End Function
Sub AddAdmission(sTb, sTbKy, ky)
    Dim rst, sql, ot, cnt, hdr, c1, c2, c3, c4, c5, c6, c7, c8, c9, recKy, clr, hasPrt
    Dim lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select * from Admission where Visitationid='" & ky & "'"
    sql = sql & " order by AdmissionDate"
    cnt = 0
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .movefirst
            hasPrt = HasPrintOutAccess(jschd, sTb & "RCP")
            response.write "<tr>"

            response.write "<td align=""center"">"
            response.write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""1"" style=""border-collapse:collapse"" >"

            response.write "<tr class=""cpHdrTr2"">"
            response.write "<td><u>No.</u></td>"
            response.write "<td><u>Admission No.</u></td>"
            response.write "<td><u>Ward</u></td>"
            response.write "<td><u>Bed</u></td>"
            response.write "<td><u>Adm. Doctor</u></td>"
            response.write "<td><u>Attend. Doctor</u></td>"
            response.write "<td><u>Status</u></td>"
            response.write "<td><u>Adm. Date</u></td>"
            response.write "<td><u>Disch. Date</u></td>"
            response.write "<td><u>Entered By</u></td>"
            response.write "<td colspan=""2""><u>Control</u></td>"
            response.write "</tr>"

            Do While Not .EOF
                cnt = cnt + 1
                recKy = .fields(sTbKy)
                clr = ""
                c1 = ""
                c2 = ""
                c3 = ""
                c4 = ""
                c5 = ""
                c6 = ""
                c7 = ""
                c8 = ""
                c9 = ""
                'Clickable Url Link
                lnkUrl = "wpg" & sTb & ".asp?PageMode=ProcessSelect&" & sTbKy & "=" & recKy
                navPop = "POP"
                If hasPrt Then
                    lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=" & sTb & "RCP&PositionForTableName=" & sTb & "&" & sTb & "ID=" & recKy
                    navPop = "NAV"
                End If
                inOut = "IN"
                fntSize = ""
                fntColor = clr
                bgColor = ""
                wdth = ""

                c1 = .fields("AdmissionID")
                If IsNull(c1) Then
                    c1 = "-"
                End If

                c2 = .fields("WardID")
                If IsNull(c2) Then
                    c2 = "-"
                End If

                c3 = .fields("BedID")
                If IsNull(c3) Then
                    c3 = "-"
                End If

                c4 = .fields("MedicalStaffID")
                If IsNull(c4) Then
                    c4 = "-"
                End If

                c5 = .fields("MedicalStaff2ID")
                If IsNull(c5) Then
                    c5 = "-"
                End If

                c6 = .fields("AdmissionStatusID")
                If IsNull(c6) Then
                    c6 = "-"
                End If

                c7 = .fields("AdmissionDate")
                If IsNull(c7) Then
                    c7 = "-"
                End If
                If IsDate(c7) Then
                    c7 = FormatDateDetail(CDate(c7))
                End If

                c8 = .fields("DischargeDate")
                If IsNull(c8) Then
                    c8 = "-"
                End If
                If IsDate(c8) Then
                    c8 = FormatDateDetail(CDate(c8))
                End If

                c9 = .fields("SystemUserID")
                If IsNull(c9) Then
                    c9 = "-"
                End If

                clr = ""
                If UCase(c6) = "A003" Then
                    clr = "#ffdddd"
                ElseIf UCase(c6) = "A001" Then
                    clr = "#ddffdd"
                ElseIf UCase(c6) = "A001" Then
                    clr = "#ffffdd"
                End If

                'Clickable Url Link
                lnkUrl = "wpg" & sTb & ".asp?PageMode=ProcessSelect&" & sTbKy & "=" & recKy
                navPop = "POP"
                inOut = "IN"
                fntSize = ""
                fntColor = ""
                bgColor = clr
                wdth = ""
                response.write "<tr bgcolor=""" & clr & """>"
                response.write "<td class=""cpHdrTd""><b>" & CStr(cnt) & "</b></td>"

                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>" & c1 & "</b>"
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, "12", fntColor, bgColor, wdth
                response.write "</td>"

                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = GetComboName("Ward", c2)
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
                response.write "</td>"

                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = GetComboName("Bed", c3)
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
                response.write "</td>"

                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = GetComboName("MedicalStaff", c4)
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
                response.write "</td>"

                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = c5
                lnkText = GetComboName("MedicalStaff2", c5)
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
                response.write "</td>"

                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = GetComboName("AdmissionStatus", c6)
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
                response.write "</td>"

                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = c7
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
                response.write "</td>"

                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = c8
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
                response.write "</td>"

                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = GetComboName("Staff", GetComboNameFld("SystemUser", c9, "StaffID"))
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
                response.write "</td>"

                'Clickable Url Link
                response.write "<td>"
                If (UCase(c6) <> "A003") Then
                    lnkCnt = lnkCnt + 1
                    lnkID = "lnk" & CStr(lnkCnt)
                    If UCase(c3) = UCase(GetWaitingBedNo(c2)) Then
                        lnkText = "Assign Bed"
                    Else
                        lnkText = "Change Bed"
                    End If
                    fntColor = "#ff0000"
                    lnkUrl = "wpgAdmissionPro.asp?PageMode=AddNew&PullupData=AdmissionID||" & c1 & "&TransProcessVal2ID=AdmissionPro-T001"
                    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
                End If
                response.write "</td>"

                'Clickable Url Link
                response.write "<td>"
                If UCase(c6) = "A001" Then
                    lnkCnt = lnkCnt + 1
                    lnkID = "lnk" & CStr(lnkCnt)
                    lnkText = "Transfer Out"
                    fntColor = "#ff0000"
                    lnkUrl = "wpgAdmissionPro.asp?PageMode=AddNew&PullupData=AdmissionID||" & c1 & "&TransProcessVal2ID=AdmissionPro-T002"
                    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inOut, fntSize, fntColor, bgColor, wdth
                End If
                response.write "</td>"

                response.write "</tr>"

                .MoveNext
            Loop
            response.write "</table>"
            response.write "</td>"
            response.write "</tr>"
        End If
        .Close
    End With
    Set rst = Nothing
    response.Flush
End Sub
'///////////TAB/////
'InsertIFrameTableTag
Sub InsertIFrameInTRTag(sect, colspan)
    Dim ky, wWth, ky2, sTy, lnkID
    If Len(Trim(sect)) > 0 Then
        ky = "seciFrm" & sect
        response.write "<tr id=""" & ky & """  style=""display:none"">"
        If IsNumeric(colspan) Then
            response.write "<td colspan=""" & colspan & """ width=""100%"">"
        Else
            response.write "<td width=""100%"">"
        End If
            response.write "<table width=""100%"" id=""tbl" & ky & """ border=""0"" cellspacing=""0"" cellpadding=""0"">"
            response.write "<tr style=""font-size:10pt"">"
                lnkID = "trslt||secHideiFrm" & sect
                response.write "<td  bgcolor=""#ffffff"" width=""5%"" align=""left"" onclick=""HideIframeSection('" & sect & "')"" style=""color:#444488"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">Hide&nbsp;Section"
                response.write "</td>"
                response.write "<td width=""95%"">"
                response.write "</td>"
            response.write "</tr>"
            response.write "<tr>"
                response.write "<td colspan=""2""  width=""100%""><iframe id=""iFrm" & sect & """ name=""iFrm" & sect & "PFN"" width=""100%"" height=""0"" frameborder=""0"" scrolling=""yes"" marginwidth=""0"" marginheight=""0""></iframe>"
                response.write "</td>"
            response.write "</tr>"
            response.write "</table>"
        response.write "</td>"
        response.write "</tr>"
        response.Flush
    End If
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
            .movefirst
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
            .movefirst
            ot = .fields("ModuleManagerID")
        End If
        .Close
    End With
    HasModuleMgrAccess = ot
    Set rstTblSql = Nothing
End Function
Sub LoadCSS()
  Dim str
  str = ""
  str = str & "<style type='text/css' id=""styPrt"">"
  str = str & ".cpHdrTd{font-size:14pt;font-weight:bold}"
  str = str & ".cpHdrTr{background-color:#fdf6f6}"
  str = str & ".cpHdrTd2{font-size:12pt;font-weight:bold}"
  str = str & ".cpHdrTr2{background-color:#ffeeee}" 'fafafa
  str = str & ".cpHdrTr3{background-color:#eeeeee}"
  str = str & ".cpHdrTr4{background-color:#ccccff}"
  str = str & ".cmpTdSty {border:1px solid #cccccc;border-collapse: collapse}"
  str = str & "</style>"
  response.write str
  response.Flush
End Sub
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
