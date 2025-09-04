'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
response.write Glob_GetBootstrap5()
response.write Glob_GetIconFontAwesome()

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
    ' If Len(dur)<>9 Then
    '   dur = FormatWorkingMonth(Now())
    ' End If
    If IsEmpty(currMs) Then
      currMs = "M001"
    End If

    LoadCSS
    InitPageScript
    SetListWhCls2

    prtUrl = "wpgItemRequest2.asp?PageMode=ProcessSelect"
    ' If HasPrintOutAccess(jSchd, "PatientMedicalRecord") Then
    '   prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation"
    ' ElseIf Len(HasModuleMgrAccess(jSchd, tb)) > 0 Then
    '   prtUrl = "wpgSelectModuleManager.asp?PositionForTableName=" & tb & "&PositionForCtxTableName=Visitation"
    ' ElseIf HasPrintOutAccess(jSchd, "VisitationRCP") Then
    '   prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation"
    '   sltTyp = "YES"
    ' End If

    cnt = 0
    cnt = cnt + 1
    nDt = Now()
    pCnt = 0

    cMth = ""
    mth = ""
    cWkDy = ""
    pCnt = 0
    dCnt = 0
    currSto = Glob_GetUserItemStore(jSchd)

    response.write "<table class=""table table-striped cmpTdSty"" cellpadding=""2"" border=""1"" cellspacing=""0"" width=""100%"" style=""font-size:10pt"">"

    response.write "<tr><td align=""left"" width=""100%"" valign=""top"" colspan=""10"">"
      response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
       response.write "<tr><td colspan=""2"" align=""center"">"

        response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""font-size:12pt"" width=""100%"">"
            response.write "<tr><td class=""cpHdrTd2"" style=""color:" & Glob_BrandingColor("sage") & """>&emsp;<u>GENERAL&nbsp;ITEMS/CONSUMABLES&nbsp;&nbsp;REQUISITIONS&nbsp;</u>&emsp;</td>"
            dTyp = GetDispType2(jSchd)
         response.write "<td>&emsp;&emsp;<b>Month:&nbsp;</b></td>"

         response.write "<td>"
         ' SetPrescriptionDays prevDys, nDt, nDt2, dur
         SetRequisitionMonth currSto, dur
         response.write "</td>"

        '  ' response.write "<td>&nbsp;&nbsp;</td>"
        '  ' response.write "<td><b>&nbsp;&nbsp;Type&nbsp;:</b></td>"
        '  ' response.write "<td>"
        '  '   SetMedicalService brnch, vDt1, vDt2, currMs
        '  ' response.write "</td>"

        '  response.write "<td>&nbsp;&nbsp;</td>"
        '  response.write "<td>&nbsp;&nbsp;</td>"

        ' '  response.write "</tr>"
        ' '  response.write "</table></td></tr>"

        ' ' response.write "<tr><td colspan=""2"" align=""left"">"
        ' ' response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""font-size:12pt"" width=""100%"">"
        ' '  response.write "<tr>"
        response.write "<td style=""" & Glob_BrandingColor("sage") & """ class=""cpHdrTd2"">&nbsp;&nbsp;<u>As&nbsp;At&nbsp;:&nbsp;&nbsp;" & FormatDateDetail(Now()) & "</u>&emsp;&emsp;&nbsp;</td>"

        lnkCnt = lnkCnt + 1
        lnkID = "trslt||lnk" & CStr(lnkCnt)
        response.write "<td onclick=""RefreshPage()"" class=""btn_"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
        response.write "<b>Refresh&nbsp;</b></td>"

        ' lnkCnt = lnkCnt + 1
        ' lnkID = "trslt||lnk" & CStr(lnkCnt)
        ' response.write "<td onclick=""cmdPrtBackOnClick()"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
        ' response.write "<b>&nbsp;<<&nbsp;</b></td>"

        ' lnkCnt = lnkCnt + 1
        ' lnkID = "trslt||lnk" & CStr(lnkCnt)
        ' response.write "<td onclick=""cmdPrintOnClick2()"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
        ' response.write "<b>&nbsp;Print&nbsp;</b></td>"

        ' lnkCnt = lnkCnt + 1
        ' lnkID = "trslt||lnk" & CStr(lnkCnt)
        ' response.write "<td onclick=""cmdPrtForwardOnClick()"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
        ' response.write "<b>&nbsp;>>&nbsp;</b></td>"

        response.write "<td valign=""top"">"
        sTb2 = "ItemRequest2"
        If HasAccessRight(uName, "frm" & sTb2, "New") Then
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = "<b>Make New Request</b>"
        lnkUrl = "wpgItemRequest2.asp?PageMode=AddNew&" '' & "&ItemRequest2ID=" & vst
        navPop = "OPEN"
        inout = "IN"
        fntSize = "10"
        fntColor = "#444488"
        bgColor = ""
        wdth = ""
        Glob_AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        End If
        response.write "</td>"

        ' ' lnkCnt = lnkCnt + 1
        ' ' lnkID = "trslt||lnk" & CStr(lnkCnt)
        ' ' response.write "<td>" '' " onclick=""OpenDialog()"" style=""color:#ff0000"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
        ' ' ' response.write "<button type='button' class='btn btn-warning' data-toggle='modal' data-target='#alerts-modal'>Notifications</button>"
        ' ' ' ' Response.write "<b>&emsp; My Alerts &emsp;</b></td>"
        ' ' response.write "</td>"


        '  response.write "<td><b>&nbsp;&nbsp;Arrange&nbsp;By&nbsp;:</b></td>"
        '     chkNm = " "
        '     chkTm = " checked "
        '     chkTm2 = "  "
        '     If UCase(ordByTyp) = UCase("Name") Then
        '         chkNm = " checked "
        '         chkTm = " "
        '         chkTm2 = " "
        '     ElseIf UCase(ordByTyp) = UCase("Time2") Then
        '         chkNm = " "
        '         chkTm = " "
        '         chkTm2 = " checked "
        '     End If
        '  ' response.write "<td><nobr>&emsp;<input type=""radio"" id=""inpOrderByType""  name=""inpOrderByType""  value=""Name"" onclick=""inpOrderByTypeOnClick()"" " & chkNm & ">Name</nobr></td>"
        '  ' response.write "<td><nobr>&emsp;<input type=""radio"" id=""inpOrderByType""  name=""inpOrderByType""  value=""Time"" onclick=""inpOrderByTypeOnClick()"" " & chkTm & ">Time</nobr></td>"
        '  response.write "<td><nobr>&emsp;<input type=""radio"" id=""inpOrderByType""  name=""inpOrderByType""  value=""Name"" onclick=""NoOfDaysOnchange()"" " & chkNm & ">Name</nobr></td>"
        '  response.write "<td><nobr>&emsp;<input type=""radio"" id=""inpOrderByType""  name=""inpOrderByType""  value=""Time"" onclick=""NoOfDaysOnchange()"" " & chkTm & ">Time (New First)</nobr></td>"
        '  response.write "<td><nobr>&emsp;<input type=""radio"" id=""inpOrderByType""  name=""inpOrderByType""  value=""Time2"" onclick=""NoOfDaysOnchange()"" " & chkTm2 & ">Time (Old First)</nobr></td>"
        response.write "</tr>"
        response.write "</table>"

       response.write "</td></tr>"
      response.write "</table>"
    response.write "</td></tr>"


    sql = "select ItemRequest2.*, Staff.StaffName "
    sql = sql & " From ItemRequest2, SystemUser, Staff Where ItemRequest2.SystemUserID=SystemUser.SystemUserID And SystemUser.StaffID=Staff.StaffID  "
    ' sql = sql & " And ItemRequest2.WorkingMonthID='" & dur & "' And ItemRequest2.BranchID='" & brnch & "' "
    sql = sql & " And ItemRequest2.WorkingMonthID='" & dur & "' And ItemRequest2.ItemCategoryID<>'ITC002' "
    If Glob_HasTransProcessAccess2("ItemRequest2Pro", uName) Then
        ' sql = sql & " And (ItemRequest2.ItemStoreID='" & currSto & "' Or ItemRequest2.ItemRequestStoreID='" & currSto & "')  "
    ElseIf Trim(currSto) = "" Then
        sql = sql & " And ItemRequest2.BranchID='" & brnch & "' "
    ElseIf Trim(currSto) <> "" Then
        sql = sql & " And (ItemRequest2.ItemStoreID='" & currSto & "' Or ItemRequest2.ItemRequestStoreID='" & currSto & "')  "
        sql = sql & " And ItemRequest2.BranchID='" & brnch & "' "
    End If
    sql = sql & " order by ItemRequest2.RequestDate desc "
    ' response.write sql
    With rst
        '.maxrecords = 50
        .Open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            wkDyNm = GetComboName("WorkingMonth", dur)
            response.write "<tr style=""font-weight:bold;font-size:12pt"" bgcolor=""#eeeeee""><td colspan=""100"" align=""left"" valign=""top"">"
            response.write "<b>" & wkDyNm & "</b>&emsp;->&emsp; " & rst.RecordCount & " Requests "
            response.write "&emsp;&emsp;"
                response.write "My Store: " & GetComboName("JobSchedule", currSto)
            response.write "</td></tr>"

            response.write "<tr style=""font-weight:bold;font-size:12pt"" bgcolor=""#eeeeee"">"
            response.write "<td valign=""top"" align=""center"">No.</td>"
            response.write "<td valign=""top"">Request&nbsp;Details</td>"
            response.write "<td valign=""top"">Request&nbsp;Items</td>"
            response.write "<td valign=""top"">Approval&nbsp;Details</td>"
            ' response.write "<td valign=""top"">Status</td>"
            response.write "<td valign=""top"">Issuance</td>"
            response.write "<td valign=""top"">Acceptance</td>"
            ' response.write "<td valign=""top"">Summary</td>"
            ' response.write "<td valign=""top"">Theatre</td>"
            ' response.write "<td valign=""top"">Control</td>"
            response.write "</tr>"
            Do While Not .EOF
                vDt = ""
                patDet = ""
                clr = "#0fff0045" ''Final, Green
                TransProcessStatID = rst.fields("TransProcessStatID")
                dtReq = .fields("RequestDate")
                tmAgo = Glob_GetHowLong(dtReq, Now())
                spTypDet = ""
                jbNm = GetComboName("JobSchedule", .fields("JobScheduleID"))
                reqDrg = .fields("ItemRequest2ID")
                drgSto = .fields("ItemStoreID")
                reqNm = .fields("ItemRequest2Name")
                spTypNm = .fields("StaffName")
                md = .fields("ItemRequestStoreID")
                pCnt = pCnt + 1
                If UCase(TransProcessStatID) = UCase("T001") Then ''Initial, red
                    clr = "#ff000045"
                ElseIf UCase(TransProcessStatID) = UCase("T002") Then ''Authorize, yellow
                    clr = "#ffff0045"
                End If
                If UCase(currSto) = UCase(md) Then
                        tmAgo = tmAgo & "<br><b>Incoming Request</b>"
                ElseIf UCase(currSto) = UCase(drgSto) Then
                        tmAgo = tmAgo & "<br><b>Outgoing Request</b>"
                Else ''If UCase(currSto)=UCase(drgSto) Then
                        ' tmAgo = tmAgo & "<br><b>Approve/Authorization</b>"
                End If


                'Requisition
                spTypDet = "<b>" & spTypNm & "<br>" & jbNm & "</b><br>"
                patDet = "<br>" & reqNm & "<br>No:&nbsp;<b>" & reqDrg & "</b><br>" & tmAgo
                ' patDet = Replace(patDet, " ", "&nbsp;")

                response.write "<tr>"
                response.write "<td valign=""top"" align=""center"" style=""background-color:" & clr & ";"">" & CStr(pCnt) & "</td>"

                ' response.write "<td valign=""top"">" & patDet & "</td>"
                response.write "<td valign=""top"">"
                response.write spTypDet & patDet & "<br>"
                'Clickable Url Link
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>Adjust Approval Quantity</b>"
                lnkUrl = prtUrl & "&ItemRequest2ID=" & reqDrg
                navPop = "OPEN"
                inout = "IN"
                fntSize = "10"
                fntColor = "#3f8a00"
                bgColor = ""
                wdth = ""
                Glob_AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"

                response.write "<td valign=""top"">"
                DisplayRequestedItems reqDrg
                response.write "</td>"

                response.write "<td valign=""top"">"
                DisplayApprovals reqDrg
                response.write "</td>"

                response.write "<td valign=""top"">"
                reqIss = DisplayIssuedItems(reqDrg)
                response.write "</td>"

                response.write "<td valign=""top"">"
                DisplayAcceptedItems reqIss
                response.write "</td>"

                ' response.write "<td valign=""top"">"
                ' 'Clickable Url Link
                ' lnkCnt = lnkCnt + 1
                ' lnkID = "lnk" & CStr(lnkCnt)
                ' lnkText = Replace("Open<br>Folder", " ", "&nbsp;")
                ' lnkUrl = prtUrl & "&VisitationID=" & reqDrg
                ' navPop = "OPEN"
                ' inout = "IN"
                ' fntSize = ""
                ' fntColor = "#8888ff"
                ' bgColor = clr
                ' wdth = ""
                ' AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                ' response.write "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
        Else
            response.write "<tr><td colspan=""100"">No Requisition for this Month and Facility</td></tr>"
        End If
        .Close
    End With

    response.write "</table>"
    Set rst = Nothing

    response.flush
    ' SetReceiptAlerts
    ' ChangeFacilityHeader

Sub DisplayRequestedItems(reqId)
    Dim rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select dr.*, d.ItemName, u.UnitOfMeasureName from ItemRequestItems dr,Items d, UnitOfMeasure u Where d.ItemID=dr.ItemID And u.UnitOfMeasureID=d.UnitOfMeasureID "
    sql = sql & " And dr.ItemRequest2ID='" & reqId & "' order by d.ItemName "
    response.write "<table class='table table-responsive table-striped table-hover'>"
    With rst
        rst.Open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            response.write "<thead><tr>"
            response.write "<th>#</th>"
            response.write "<th>Code</th>"
            response.write "<th>Item / Description</th>"
            response.write "<td align='Right'><b>Qty (Req.)</b></td>"
            response.write "<td align='Right'><b>Qty (Appr.)</b></td>"
            response.write "<th>UOM</th>"
            response.write "</tr></thead>"
            Do While Not rst.EOF
                response.write "<tr>"
                response.write "<td>" & rst.AbsolutePosition & "</td>"
                response.write "<td>" & rst.fields("ItemID") & "</td>"
                response.write "<td>" & rst.fields("ItemName") & "</td>"
                response.write "<td align='Right'>" & FormatNumber(rst.fields("RequestValue1"), 1) & "</td>"
                response.write "<td align='Right'>" & FormatNumber(rst.fields("RequestedQty"), 1) & "</td>"
                response.write "<td>" & rst.fields("UnitOfMeasureName") & "</td>"
                response.write "</tr>"
                response.flush
                rst.MoveNext
            Loop
        Else

        End If
        rst.Close
    End With
    response.write "</table>"
    Set rst = Nothing
End Sub

Sub DisplayApprovals(reqId)
    Dim rst, sql, apprLevel1, apprLevel2
    Set rst = CreateObject("ADODB.Recordset")
    apprLevel1 = False
    apprLevel2 = False
    sql = "select * from ItemRequest2Pro Where TransProcessTblID='ItemRequest2Pro' And ItemRequest2ID='" & reqId & "' "
    sql = sql & " order by TransProcessDate1 "
    response.write "<table class='table table-responsive table-striped table-hover'>"
    With rst
        rst.Open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            Do While Not rst.EOF
                response.write "<tr>"
                If UCase(rst.fields("TransProcessVal2ID")) = UCase("ItemRequest2Pro-T002") Then
                    apprLevel1 = True
                End If
                If UCase(rst.fields("TransProcessVal2ID")) = UCase("ItemRequest2Pro-T003") Then
                    apprLevel1 = True
                    apprLevel2 = True
                End If
                response.write "<td><p>#" & rst.AbsolutePosition & ". <b>" & GetComboName("TransProcessVal2", rst.fields("TransProcessVal2ID")) & "</b><br>"
                response.write "By " & Glob_FormatName2(rst.fields("SystemUserID")) & "<br><em>[" & GetComboName("JobSchedule", rst.fields("JobScheduleID")) & "]</em>"
                response.write " on " & FormatDate(rst.fields("TransProcessDate1")) & "<br>"
                response.write "" & rst.fields("TransProcessDetail")
                response.write "</p></td>"
                response.write "</tr>"
                response.flush
                rst.MoveNext
            Loop
        Else
            response.write "<tr><th colspan='100' style='color:red;font-style:italic;'>No Approval</th></tr>"
        End If

        sTb2 = "ItemRequest2Pro"
        If Not apprLevel1 Then
            response.write "<tr><td>"
            ' If HasAccessRight(uName, "frm" & sTb2, "New") Then
            If HasAccessRight(uName, "frm" & sTb2, "New") And (Glob_HasTransProcessAccess(sTb2, uName, "T001", "T002") Or Glob_HasTransProcessAccess(sTb2, uName, "T002", "T002")) Then
                'Clickable Url Link
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>1. Authorize Request</b>"
                lnkUrl = "wpg" & sTb2 & ".asp?PageMode=AddNew&TransProcessVal2ID=ItemRequest2Pro-T002&PullupData=ItemRequest2ID||" & reqId
                navPop = "POP"
                inout = "IN"
                fntSize = "10"
                fntColor = "#444488"
                bgColor = ""
                wdth = ""
                Glob_AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            End If
            response.write "</td></tr>"
        End If
        If Not apprLevel2 Then
            response.write "<tr><td>"
            ' If HasAccessRight(uName, "frm" & sTb2, "New") Then
            If HasAccessRight(uName, "frm" & sTb2, "New") And (Glob_HasTransProcessAccess(sTb2, uName, "T001", "T003") Or Glob_HasTransProcessAccess(sTb2, uName, "T002", "T003")) Then
                'Clickable Url Link
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>2. Approve Request</b>"
                lnkUrl = "wpg" & sTb2 & ".asp?PageMode=AddNew&TransProcessVal2ID=ItemRequest2Pro-T003&PullupData=ItemRequest2ID||" & reqId
                navPop = "POP"
                inout = "IN"
                fntSize = "10"
                fntColor = "#444488"
                bgColor = ""
                wdth = ""
                Glob_AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            End If
            response.write "</td></tr>"
        End If
        rst.Close
    End With
    response.write "</table>"
    Set rst = Nothing
End Sub

Function DisplayIssuedItems(reqId)
    Dim rst, sql, ot
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select di.*, d.ItemName from ItemIssueItems di, Items d where d.ItemID=di.ItemID and di.ItemRequest2ID='" & reqId & "' "
    sql = sql & " order by d.ItemName "
    ot = ""
    response.write "<table class='table table-responsive table-striped table-hover'>"
    With rst
        rst.Open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            response.write "<thead><tr><th colspan='100'>"
            response.write "Issued By: " & Glob_FormatName2(rst.fields("SystemUserID")) & " @ " & FormatDate(rst.fields("IssuedDate1"))
            response.write "</th></tr></thead>"

            response.write "<thead><tr>"
            response.write "<th>#</th>"
            response.write "<th>Code</th>"
            response.write "<th>Item / Description</th>"
            response.write "<td align='Right'><b>Qty</b></td>"
            response.write "</tr></thead>"
            Do While Not rst.EOF
                response.write "<tr>"
                response.write "<td>" & rst.AbsolutePosition & "</td>"
                response.write "<td>" & rst.fields("ItemID") & "</td>"
                response.write "<td>" & rst.fields("ItemName") & "</td>"
                response.write "<td align='Right'>" & FormatNumber(rst.fields("IssuedQty"), 1) & "</td>"
                response.write "</tr>"
                ot = rst.fields("ItemIssueID")
                response.flush
                rst.MoveNext
            Loop
        Else
            response.write "<tr><th colspan='100'>No Issued Items</th></tr>"
            response.write "<tr><td>"
            sTb2 = "ItemIssue"
            If HasAccessRight(uName, "frm" & sTb2, "New") Then
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>&nbsp;&nbsp;Issue Items</b>"
                lnkUrl = "wpg" & sTb2 & ".asp?PageMode=AddNew&PullupData=ItemRequest2ID||" & reqId
                navPop = "POP"
                inout = "IN"
                fntSize = "10"
                fntColor = "#444488"
                bgColor = ""
                wdth = ""
                Glob_AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            End If
            response.write "</td></tr>"
        End If
        rst.Close
    End With
    response.write "</table>"
    Set rst = Nothing
    DisplayIssuedItems = ot
End Function

Sub DisplayAcceptedItems(reqId)
    Dim rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select da.*, d.ItemName from ItemAcceptItems da, Items d where d.ItemID=da.ItemID and da.ItemIssueID='" & reqId & "' "
    sql = sql & " order by d.ItemName "

    response.write "<table class='table table-responsive table-striped table-hover'>"
    With rst
        rst.Open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            response.write "<thead><tr><th colspan='100'>"
            response.write "Accepted By: " & Glob_FormatName2(rst.fields("SystemUserID")) & " @ " & FormatDate(rst.fields("AcceptDate1"))
            response.write "</th></tr></thead>"

            response.write "<thead><tr>"
            response.write "<th>#</th>"
            response.write "<th>Code</th>"
            response.write "<th>Item / Description</th>"
            response.write "<td align='Right'><b>Qty</b></td>"
            response.write "</tr></thead>"
            Do While Not rst.EOF
                response.write "<tr>"
                response.write "<td>" & rst.AbsolutePosition & "</td>"
                response.write "<td>" & rst.fields("ItemID") & "</td>"
                response.write "<td>" & rst.fields("ItemName") & "</td>"
                response.write "<td align='Right'>" & FormatNumber(rst.fields("AcceptAmt1"), 1) & "</td>"
                response.write "</tr>"
                response.flush
                rst.MoveNext
            Loop
        Else
            response.write "<tr><th colspan='100'>No Accepted Items</th></tr>"
            response.write "<tr><td>"
            sTb2 = "ItemAccept"
            If HasAccessRight(uName, "frm" & sTb2, "New") Then
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>&nbsp;&nbsp;Accept Items to My Stock</b>"
                lnkUrl = "wpg" & sTb2 & ".asp?PageMode=AddNew&PullupData=ItemIssueID||" & reqId
                navPop = "POP"
                inout = "IN"
                fntSize = "10"
                fntColor = "#3f8a00"
                bgColor = ""
                wdth = ""
                Glob_AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            End If
            response.write "</td></tr>"
        End If
        rst.Close
    End With
    response.write "</table>"
    Set rst = Nothing
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
    Dim Str
    ExtractWorkingDate = Null
    Str = Trim(wkDay)
    If Len(Str) = 11 Then
      If UCase(Left(Str, 3)) = "DAY" Then
        ExtractWorkingDate = CDate(Mid(Str, 10, 2) & " " & MonthName(CInt(Mid(Str, 8, 2)), 1) & " " & Mid(Str, 4, 4))
      End If
    End If
End Function

Function HasPrintOutAccess(jb, prt)
  Dim rstTblSql, sql, ot
  ot = False
  Set rstTblSql = CreateObject("ADODB.Recordset")
  With rstTblSql
    sql = "select JobScheduleID from printoutalloc "
    sql = sql & " where printlayoutid='" & prt & "' and jobscheduleid='" & jb & "'"
    .Open qryPro.FltQry(sql), conn, 3, 4
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
    .Open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      ot = .fields("ModuleManagerID")
    End If
    .Close
  End With
  HasModuleMgrAccess = ot
  Set rstTblSql = Nothing
End Function

Sub SetListWhCls2()
  Dim jb
  jb = Trim(jSchd)
  dispTyp = GetDispType2(jb)

  ' If dispTyp = "LAB" Then
  '   lstWhCls2 = " and (LabByDoctor.TestCategoryID='B13' Or LabByDoctor.TestGroupID='B13')"
  ' ElseIf dispTyp = "IMAGING" Then
  '   lstWhCls2 = " and (LabByDoctor.TestCategoryID='B19' Or LabByDoctor.TestGroupID='B19')"
  ' End If
  ' lstWhCls2 = lstWhCls2 & " And LabByDoctor.WorkingDayID >= 'DAY20220501' "
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

Sub SetRequisitionMonth(store, currMth)
    Set rst = CreateObject("ADODB.Recordset")
    dyHt = "<select size=""1"" name=""NoOfDays"" id=""NoOfDays"" onchange=""NoOfDaysOnchange()"">"
    dyHt = dyHt & "<option value=""""></option>"

    cYr = ""
    yr = ""
    dtDply = CDate(Glob_DeploymentDate())
    lstWhCls2 = " And dr.WorkingMonthID>='" & FormatWorkingMonth(dtDply) & "' "
    sql0 = "select distinct dr.WorkingMonthID, wm.WorkingMonthName, wm.WorkingYearID from ItemRequest2 dr, WorkingMonth wm where dr.WorkingMonthID=wm.WorkingMonthID "
    sql0 = sql0 & " And dr.BranchID='" & brnch & "' " & lstWhCls2
    ' If Trim(store)<>"" Or Not Glob_HasTransProcessAccess2("ItemRequest2Pro", uName) Then
    If Not (Glob_HasStaffLevel(uName) Or Glob_HasTransProcessAccess2("ItemRequest2Pro", uName)) Then
        sql = sql & " And (ItemRequest2.ItemStoreID='" & store & "' Or ItemRequest2.ItemRequestStoreID='" & store & "')  "
    End If
    sql0 = sql0 & " order by dr.WorkingMonthID desc"
    '' response.write sql0

    With rst
      .Open qryPro.FltQry(sql0), conn, 3, 4
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          wkMth = Trim(.fields("WorkingMonthID"))
          dyNm = Trim(.fields("WorkingMonthName")) '' & " -> " & GetComboName("WorkingYear", yr)
          yr = Trim(.fields("WorkingYearID"))
          If UCase(cYr) <> UCase(yr) Then
             dyHt = dyHt & "<optGroup label=""" & GetComboName("WorkingYear", yr) & """>"
             cYr = yr
          End If
          If UCase(CStr(currMth)) = UCase(wkMth) Then
            dyHt = dyHt & "<option value=""" & CStr(wkMth) & """ selected>" & dyNm & "</option>"
          Else
             dyHt = dyHt & "<option value=""" & CStr(wkMth) & """>" & dyNm & "</option>"
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
    .Open qryPro.FltQry(sql), conn, 3, 4
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

Sub SetItemStore(br, jb, currSto)
  Dim rs, ot, sql, dSto
  Set rs = CreateObject("ADODB.Recordset")
  ' sql = "select  distinct ds.ItemStoreID, ds.ItemStoreName from ItemStore ds, ItemStore2 ds2 where ds.JobScheduleID=ds2.JobScheduleID "
  ' sql = sql & " And ds.BranchID='" & br & "' and ds.ItemStoreID IN ('M0601','M0602','M0603','M0604','M0605','M0612') "
  ' sql = sql & " order by ds.ItemStoreID "

  sql = "select distinct ds.ItemStoreID, ds.ItemStoreName from ItemStore ds Where ds.BranchID='" & br & "' "
  sql = sql & " And ds.ItemStoreID IN ('M0601','M0602','M0603','M0604','M0605','M0612') "
  ' sql = sql & " UNION "
  ' sql = sql & " select  distinct ds.ItemStoreID, ds.ItemStoreName from ItemStore ds, ItemStore2 ds2 where ds.JobScheduleID=ds2.JobScheduleID "
  ' sql = sql & " And ds.BranchID='" & br & "' and ds.ItemStoreID IN ('M0601','M0602','M0603','M0604','M0605','M0612') "
  sql = sql & " order by ds.ItemStoreID "

  ot = "<select id=""ItemStore"" name=""ItemStore"" onchange=""ItemStoreOnchange()"">"
  ot = ot & "<option></option>"
  With rs
    .Open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        dSto = .fields("ItemStoreID")
        If UCase(dSto) = UCase(currSto) Then
          ot = ot & "<option value=""" & dSto & """ selected>" & .fields("ItemStoreName") & "</option>"
        Else
          ot = ot & "<option value=""" & dSto & """>" & .fields("ItemStoreName") & "</option>"
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

Sub SetItemStoreIC(br, jb, currSto)
  Dim rs, ot, sql, dSto
  Set rs = CreateObject("ADODB.Recordset")
  sql = "select distinct js.JobScheduleID, js.JobScheduleName from JobSchedule js "
  sql = sql & " Where js.JobScheduleID IN (SELECT JobScheduleID From ItemStore2 Where JobScheduleID IN   "
  sql = sql & "  ('" & jb & "','M06IC','M0601IC','M0602IC','M0603IC','M0604IC','M0605IC','M0612IC') and BranchID='" & br & "'  "
  sql = sql & " )  "
  sql = sql & " order BY js.Jobscheduleid "

  ot = "<select id=""ItemStore"" name=""ItemStore"" onchange=""ItemStoreOnchange()"">"
  ot = ot & "<option></option>"
  With rs
    .Open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        dSto = .fields("JobScheduleID")
        ' If UCase(dSto) = UCase(currSto) Then
        If UCase(dSto) = UCase(jb) Then
          ot = ot & "<option value=""" & dSto & """ selected>" & .fields("JobScheduleName") & "</option>"
        Else
          ot = ot & "<option value=""" & dSto & """>" & .fields("JobScheduleName") & "</option>"
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

Function GetItemStore(jb)
    Dim ot
    Set rst = CreateObject("ADODB.Recordset")
    ot = GetComboNameFld("ItemStore", jb, "JobScheduleID")
    If Len(Trim(ot)) > 0 Then
        ot = Trim(ot)
    Else
        ot = ""
        sql = "select top 1 * from ItemStore2 Where JobScheduleID='" & jb & "' "
        With rst
            rst.Open qryPro.FltQry(sql), conn, 3, 4
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                ot = rst.fields("ItemStoreID")
            End If
            rst.Close
        End With
    End If
    Set rst = Nothing
    GetItemStore = ot
End Function

Sub LoadCSS()
  Dim Str
  Str = ""
  Str = Str & "<style type='text/css' id=""styPrt"">"
  Str = Str & ".cpHdrTd{font-size:14pt;font-weight:bold}"
  Str = Str & ".cpHdrTr{background-color:#eeeeee}"
  Str = Str & ".cpHdrTd2{font-size:12pt;font-weight:bold}"
  Str = Str & ".cpHdrTr2{background-color:#eeeeee}" 'fafafa
  Str = Str & ".table{font-size:14px;}"
  Str = Str & "</style>"
  response.write Str

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
  htStr = htStr & "ur='wpgPrtPrintLayoutAll.asp?PrintLayoutName=MonitorItemRequisition&PositionForTableName=WorkingDay';" & vbCrLf
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
  htStr = htStr & "ur='wpgPrtPrintLayoutAll.asp?PrintLayoutName=MonitorItemRequisition&PositionForTableName=WorkingDay';" & vbCrLf
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
  htStr = htStr & "ur='wpgPrtPrintLayoutAll.asp?PrintLayoutName=MonitorItemRequisition&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "ur=ur + '&WorkingDayID=DAY20160401&NoOfDays=' + dy + '&OrderByType=' + ordByTyp + '&Specialist=' + sp + '&MedicalService=' + ms;" & vbCrLf
  htStr = htStr & "window.location.href=processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  ' ItemStoreOnchange()
  htStr = htStr & "function ItemStoreOnchange(){}" & vbCrLf

  ' ' ChangeJobSchedule()
  ' htStr = htStr & "function ChangeJobSchedule(){" & vbCrLf
  ' htStr = htStr & "var ur,ds;" & vbCrLf
  ' htStr = htStr & "ds=GetEleVal('ItemStore');" & vbCrLf
  ' htStr = htStr & "ur='wpgSystemUser.asp?PageMode=ProcessSelect&ActionType=ChangeItemStore&SystemUserID=" & uName & "';" & vbCrLf
  ' htStr = htStr & "ur=ur + '&ItemStoreID=' + ds + '&JobScheduleID=" & jSchd & "';" & vbCrLf
  ' htStr = htStr & "window.location.href=processurl(ur);" & vbCrLf
  ' htStr = htStr & "}" & vbCrLf

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
