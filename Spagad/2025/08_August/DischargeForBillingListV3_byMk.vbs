'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
    'M006 - T009
    'M007 - T010
    'M008 - T004
    'M009 - T005
    'M010 - T006
    'M011 - T007
    Dim str, lnkCnt, wrd, jb, lstWhCls, lstWhCls2, dispTyp, wdNm, iFUrl, selectedYear
    
    selectedYear = Trim(Request.QueryString("YearFilter"))
'    If IsNull(selectedYear) Or selectedYear = "" Then
    If Len(selectedYear) = 0 Then
        selectedYear = "YRS" & Year(Now)
    End If
    
    lnkCnt = 0
    jb = jSchd

    LoadCSS

    SetDischargeListWhCls
    dispTyp = GetDispType2(jb)
    DisplayDischargeList
    Sub DisplayDischargeList()
    Dim rst, rst2, rst3, sql, pat, adm, currWd, prtUrl, modMgr
    Dim st1, st2, st3, arr, ul, num, bTy, wd, wdSc, wrdNm, bedNm, bd, vst, clr, insSc, dichPat, rType
    Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, mdO, clss, otTot
    Set rst = CreateObject("ADODB.Recordset")
    Set rst2 = CreateObject("ADODB.Recordset")
    Set rst3 = CreateObject("ADODB.Recordset")
    Set rstDropdown = CreateObject("ADODB.Recordset")
    currWd = ""
    
    ' Year dropdown options
'    sql = "SELECT WorkingYearID, WorkingYearName FROM WorkingYear ORDER BY WorkingYearName DESC"
    sql = "SELECT distinct adm.WorkingYearID, yrs.WorkingYearName"
    sql = sql & " FROM Admission adm"
    sql = sql & " join WorkingYear yrs ON yrs.WorkingYearID = adm.WorkingYearID"
    sql = sql & " ORDER BY WorkingYearName DESC"
    rstDropdown.open sql, conn, 3, 4
    dropdownOptions = ""
            dropdownOptions = "<option value=''>" & GetComboName("workingyear", selectedYear) & "</option>"
    With rstDropdown
        If .recordCount > 0 Then
'            .MoveFirst
            .MoveFirst
            Do Until .EOF
                dropdownOptions = dropdownOptions & "<option value='" & .fields("WorkingYearID") & "'>" & .fields("WorkingYearName") & "</option>"
'                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With
    rstDropdown.Close
    Set rstDropdown = Nothing
    
    prtUrl = "wpgVisitation.asp?PageMode=ProcessSelect"
    If HasPrintOutAccess(jSchd, "PatientMedicalRecord") Then
        prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation"
    ElseIf Len(HasModuleMgrAccess(jSchd, "Visitation")) > 0 Then
        prtUrl = "wpgSelectModuleManager.asp?PositionForTableName=Visitation&PositionForCtxTableName=Visitation"
    ElseIf HasPrintOutAccess(jSchd, "VisitationRCP") Then
        prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation"
    End If
    
    response.write "<meta http-equiv=""refresh"" content=""180"">"
    response.write "<div class='header-container'>"
    response.write "<div class='filters'>"
    response.write "<div class='filter-row'>"
'    response.write "    <label for='yearFilter' class='font-style'>Filter: </label>"
    response.write "    <select class='myselect' id='yearFilter' name='yearFilter' onchange='updateUrl()'>"
    response.write dropdownOptions
    response.write "    </select>"
    response.write "</div>"
'    response.write "<div class='filter-row'>"
'    response.write "    <button class='mybutton' type='button' onclick='updateUrl()'>Show Data</button>"
'    response.write "</div>"
    response.write "</div>"
    response.write "</div>"
    
    response.write "<script>"
    response.write "    function updateUrl() {"
    response.write "        const year = document.getElementById('yearFilter').value;"
    response.write "        const baseUrl = 'http://172.19.0.36/hms/wpgPrtPrintLayoutAll.asp';"
    response.write "        const params = new URLSearchParams({"
    response.write "            PrintLayoutName: 'DischargeForBillingList',"
    response.write "            PositionForTableName: 'WorkingDay',"
    response.write "            WorkingDayID: 'DAY20180515',"
    response.write "            YearFilter: year"
    response.write "        });"
    response.write "        const newUrl = baseUrl + '?' + params.toString();"
    response.write "        window.location.href = newUrl;"
    response.write "    }"
    response.write "</script>"
    
    
    response.write "<meta http-equiv=""refresh"" content=""180"">"
    response.write "<table cellspacing=""0"" cellpadding=""0"" border=""0"" width=""100%"">"
    
    response.write "<tr>"
    response.write "<td>"
    
    sql = "select a.visitationID,a.WardID,a.BedID,a.PatientID,a.MainInfo3,a.TransProcessStatID,a.AdmissionDate,a.DischargeDate,a.AdmissionID,a.InsuranceSchemeID, v.ReceiptTypeID, a.SponsorID, ISNULL(SUM(PatientBill.BillAmt3),0) BillAmt3 "
    sql = sql & " from Admission as a "
    sql = sql & " inner join Visitation as v on a.VisitationID = v.VisitationID "
    sql = sql & " left join PatientBill on PatientBill.VisitationID = v.VisitationID "
    sql = sql & " where 1=1 " & lstWhCls2
    sql = sql & " and a.WorkingYearID = '" & selectedYear & "'"
    sql = sql & " and a.AdmissionStatusID='A001' " & lstWhCls
    sql = sql & " group by a.visitationID,a.WardID,a.BedID,a.PatientID,a.MainInfo3,a.TransProcessStatID,a.AdmissionDate,a.DischargeDate,a.AdmissionID,a.InsuranceSchemeID, v.ReceiptTypeID, a.SponsorID "
    sql = sql & " order by a.WardID,a.AdmissionID "
    
    '  sql = sql & " " & lstWhCls & " order by a.WardID,a.AdmissionID"
    
    
    response.write "<table cellspacing=""0"" cellpadding=""2"" border=""1"" style=""font-size:10pt;border-collapse:collapse;border-color:#444444"" width=""100%"">"
        'Header
        cnt = 0
    '    If (UCase(jSchd) = "BILLINGHEAD") Or (UCase(jSchd) = "CLAIMMANAGER") Or (UCase(jSchd) = "M13") Then
    '      response.write "<tr bgcolor=""#eeeeee"" style=""font-size:14pt""><td colspan=""14"">"
    '      response.write "<table cellspacing=""0"" cellpadding=""2"" border=""1"" style=""font-size:10pt;border-collapse:collapse;border-color:#444444"" width=""100%"">"
    '      response.write "<tr bgcolor=""#eeeeee"" style=""font-size:14pt"">"
    '
    '      iFUrl = "wpgBrowseviewLayout.asp?BrowseviewName=TreatChargesBySystemUser_TreatmentBilling&HomeDispType=11&OpenProp=Pop"
    '      response.write "<td colspan=""5"" width=""40%"" valign=""top"">"
    '      response.write "<iframe id=""iFrmBilling"" width=""100%"" onload=""iFrmOnLoad('iFrmBilling')"" height=""350"" align=""left"" frameborder=""0"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"" src=""" & iFUrl & """></iframe>"
    '      response.write "</td>"
    '
    '      iFUrl = "wpgBrowseviewLayout.asp?BrowseviewName=TreatChargesByWorkingDay_SystemUser_TreatmentBilling&HomeDispType=11&OpenProp=Pop"
    '      response.write "<td colspan=""5"" width=""30%"" valign=""top"">"
    '      response.write "<iframe id=""iFrmBilling2"" width=""100%"" onload=""iFrmOnLoad('iFrmBilling2')"" height=""350"" align=""left"" frameborder=""0"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"" src=""" & iFUrl & """></iframe>"
    '      response.write "</td>"
    '
    '      iFUrl = "wpgBrowseviewLayout.asp?BrowseviewName=AdmissionByTransProcessVal_WorkingDay_WardBilling&HomeDispType=11&OpenProp=Pop"
    '      response.write "<td colspan=""4"" width=""30%"" valign=""top"">"
    '      response.write "<iframe id=""iFrmBilling3"" width=""100%"" onload=""iFrmOnLoad('iFrmBilling3')"" height=""350"" align=""left"" frameborder=""0"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"" src=""" & iFUrl & """></iframe>"
    '      response.write "</td>"
    '
    '      response.write "</tr></table></td></tr>"
    '    End If

    '    response.write "<tr bgcolor=""#eeeeee"" style=""font-size:14pt""><td colspan=""14"">"
    '    response.write "<table cellspacing=""0"" cellpadding=""2"" border=""1"" style=""font-size:10pt;border-collapse:collapse;border-color:#444444"" width=""100%"">"
    '    response.write "<tr bgcolor=""#eeeeee"" style=""font-size:14pt"">"
    '
    '    'Consumable Cost Sheet
    '    iFUrl = "wpgBrowseviewLayout.asp?BrowseviewName=TreatChargesByJobSchedule_TreatmentWard&HomeDispType=11&OpenProp=Pop"
    '    response.write "<td colspan=""7"" width=""50%"" valign=""top"">"
    '    response.write "<iframe id=""iFrmWardConsum"" width=""100%"" onload=""iFrmOnLoad('iFrmWardConsum')"" height=""350"" align=""left"" frameborder=""0"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"" src=""" & iFUrl & """></iframe>"
    '    response.write "</td>"
    '
    '    'Medical Cost Sheet
    '    iFUrl = "wpgBrowseviewLayout.asp?BrowseviewName=DrugSaleItemsByJobSchedule_DrugWard&HomeDispType=11&OpenProp=Pop"
    '    response.write "<td colspan=""7"" width=""50%"" valign=""top"">"
    '    response.write "<iframe id=""iFrmWardDrug"" width=""100%"" onload=""iFrmOnLoad('iFrmWardDrug')"" height=""350"" align=""left"" frameborder=""0"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"" src=""" & iFUrl & """></iframe>"
    '    response.write "</td>"
    '    response.write "</tr></table></td></tr>"
        
        
        response.write "<tr bgcolor=""#eeeeee"" style=""font-size:14pt"">"
        If Len(selectedYear) = 0 Then
            response.write "<td colspan=""15""><b>All Patients Awaiting Discharge As At: <font color=""#6666cc""><u>" & FormatDateDetail(Now()) & "</u></font></b></td>"
        Else
            response.write "<td colspan=""15""><b>All Patients Awaiting Discharge As At: <font color=""#6666cc""><u>" & GetComboName("WorkingYear", selectedYear) & "</u></font></b></td>"
        End If
        response.write "</tr>"
        
    rst2.open sql, conn, 3, 4
    'SetPageMessages sql
    If rst2.recordCount > 0 Then
        Do While Not rst2.EOF
        cnt = cnt + 1
        pat = rst2.fields("PatientID")
        adm = rst2.fields("AdmissionID")
        wd = rst2.fields("WardID")
        bd = rst2.fields("BedID")
        vst = rst2.fields("VisitationID")
        mdO = rst2.fields("TransProcessStatID")
        insSc = rst2.fields("InsuranceSchemeID")
        rType = rst2.fields("ReceiptTypeID")
        spnID = rst2.fields("SponsorID")
        outBill = rst2.fields("BillAmt3")
        otTot = ""
        dichPat = False
        If Not IsNull(rst2.fields("MainInfo3")) Then
            otTot = rst2.fields("MainInfo3")
        End If
        clr = "#ffffff"
        clss = ""
        If UCase(mdO) = "T009" Then 'Request for billing
            clr = "#ffff44"
        ElseIf UCase(mdO) = "T010" Then 'Completed for Printing
            clr = "#ccffcc"
            If IsNumeric(otTot) Then
                If CDbl(otTot) <= 0 Or (rType <> "R001" And spnID <> "EME") Then
                    clss = " class=""animTR"" "
                    dichPat = True
                End If
            End If
            If outBill <= 0 Then
                clss = " class=""animTR"" "
                dichPat = True
            End If
        ElseIf UCase(mdO) = "T004" Then 'Referred t Social Welfare
            clr = "#aaaaff"
        ElseIf UCase(mdO) = "T013" Then 'Completed for Review
            clr = "#00ffff"
        End If
        If UCase(wd) <> UCase(currWd) Then
            wrdNm = GetComboName("Ward", wd)
            response.write "<tr bgcolor=""#eeeeee"">"
            response.write "<td colspan=""15"" style='font-size: 14pt;; font-weight: bold;'>" & wrdNm & "</td>"
            response.write "</tr>"

            response.write "<tr bgcolor=""#eeeeee"">"
            response.write "<td align=""center""><b>No</b></td>"
            response.write "<td align=""center""><b>Folder No</b></td>"
            response.write "<td align=""center""><b>Patient Name</b></td>"
            response.write "<td align=""center""><b>Visit. No</b></td>"
            response.write "<td align=""center""><b>Type</b></td>"
            response.write "<td align=""center""><b>Admission. No</b></td>"
            response.write "<td align=""center""><b>Bed</b></td>"
            response.write "<td align=""center""><b>Admit Date</b></td>"
            response.write "<td align=""center""><b>Disch Date</b></td>"
            response.write "<td align=""center"" colspan=""6""><b>Control</b></td>"
            response.write "</tr>"
            currWd = wd
        End If
        
        lnkUrl = "wpgAdmission.asp?PageMode=ProcessSelect&AdmissionID=" & adm
        navPop = "POP"
        inout = "IN"
        fntSize = ""
        fntColor = "#222244"
        bgColor = clr
        wdth = ""
        
        response.write "<tr bgcolor=""" & clr & """> "
        response.write "<td align=""left"">" & CStr(cnt) & "</td>"
        
        response.write "<td align=""left"">"
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = pat
        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        response.write "</td>"
        
        response.write "<td align=""left""" & clss & ">"
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = Replace(GetComboName("Patient", pat), " ", "&nbsp;")
        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        response.write "</td>"
        
        response.write "<td align=""left"">"
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = vst
        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        response.write "</td>"
        
        response.write "<td align=""left"">"
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = Replace(GetComboName("InsuranceScheme", insSc), " ", "&nbsp;")
        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        response.write "</td>"
        
        response.write "<td align=""left"">"
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = adm
        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        response.write "</td>"
        
        response.write "<td align=""left"">"
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = Replace(GetComboName("Bed", bd), " ", "&nbsp;")
        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        response.write "</td>"
        
        response.write "<td align=""left"">"
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = FormatDate(rst2.fields("AdmissionDate"))
        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        response.write "</td>"
        
        response.write "<td align=""left"">"
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = "-"
        If Not IsNull(rst2.fields("DischargeDate")) Then
            If IsDate(rst2.fields("DischargeDate")) Then
            lnkText = FormatDate(rst2.fields("DischargeDate"))
            End If
        End If
        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        response.write "</td>"
        
        response.write "<td align=""left"">"
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = "<b>Diag</b>"
        lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=DischargeForBilling&PositionForTableName=WorkingDay&WorkingDayID=DAY20180515&DisplayType=Diagnosis&VisitationID=" & vst
        navPop = "POP"
        inout = "IN"
        fntSize = "9"
        fntColor = "#4444cc"
        bgColor = clr
        wdth = ""
        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        response.write "</td>"
        
        If (UCase(dispTyp) = "WARDNURSE") And ((UCase(mdO) = "T009") Or (UCase(mdO) = "T013")) Then 'Ward Nurse and Discharge for Billing 15 Feb 2019,22 May 2019'
            response.write "<td align=""left""" & clss & ">"
            response.write "</td>"
        Else
            If CDate(GetComboNameFld("Visitation", vst, "VisitDate")) < CDate("01 Mar 2021") Then
                response.write "<td align=""left""" & clss & ">"
                'Clickable Url Link
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>Bill</b>"
                'lnkUrl = "wpgSelectPrintLayout.asp?PositionForTableName=Admission&AdmissionID=" & adm
                lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=Admission1&PositionForTableName=Admission&AdmissionID=" & adm
                navPop = "POP"
                inout = "IN"
                fntSize = "9"
                fntColor = "#ff0000"
                bgColor = clr
                wdth = ""
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
            Else
                'progressive patient bill
                response.write "<td align=""left""" & clss & ">"
                'Clickable Url Link
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>Prog. Bill</b>"
                'lnkUrl = "wpgSelectPrintLayout.asp?PositionForTableName=Admission&AdmissionID=" & adm
                lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationBill7&PositionForTableName=Visitation&VisitationID=" & vst
                navPop = "POP"
                inout = "IN"
                fntSize = "9"
                fntColor = "#ff0000"
                bgColor = clr
                wdth = ""
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
            End If
        End If
        response.write "<td align=""left"">"
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = "<b>Folder</b>"
        lnkUrl = prtUrl & "&VisitationID=" & vst
        navPop = "POP"
        inout = "IN"
        fntSize = "9"
        fntColor = "#ff0000"
        bgColor = clr
        wdth = ""
        AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        response.write "</td>"
        
        If UCase(dispTyp) = "BILLING" Then 'Not ward
            response.write "<td align=""left"">"
            'Clickable Url Link
            lnkCnt = lnkCnt + 1
            lnkID = "lnk" & CStr(lnkCnt)
            lnkText = "<b>Process Bill</b>"
            lnkUrl = "wpgNavigateFrame.asp?FrameType=WorkFlow&PositionForTableName=Visitation&VisitationID=" & vst
            navPop = "POP"
            inout = "IN"
            fntSize = "9"
            fntColor = "#ff0000"
            bgColor = clr
            wdth = ""
            AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            response.write "</td>"
            
            
            If (UCase(mdO) = "T009") Then
                'Clickable Url Link
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>Return To Ward</b>"
                lnkUrl = "wpgAdmissionPro.asp?PageMode=AddNew&PullupData=AdmissionID||" & adm & "&TransProcessVal2ID=AdmissionPro-T014"
                navPop = "POP"
                inout = "IN"
                fntSize = "9"
                fntColor = "#ff0000"
                bgColor = clr
                wdth = ""
                response.write "<td align=""left"">"
                    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
            End If
            
            'Clickable Url Link
            lnkCnt = lnkCnt + 1
            lnkID = "lnk" & CStr(lnkCnt)
            lnkText = "-"
            lnkUrl = "-"
            If (UCase(mdO) = "T009") Then 'Complete for Review
                lnkText = "<b>Complete For Review</b>"
                lnkUrl = "wpgAdmissionPro.asp?PageMode=AddNew&PullupData=AdmissionID||" & adm & "&TransProcessVal2ID=AdmissionPro-T013"
                'lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=DischargeForBilling&PositionForTableName=WorkingDay&WorkingDayID=DAY20180515&VisitationID=" & vst
            ElseIf (UCase(mdO) = "T013") Then 'Complete for Printing
                If HasTransProcessAccess("AdmissionPro", uname, "T013", "T010") Then
                lnkText = "<b>Complete For Printing</b>"
                lnkUrl = "wpgAdmissionPro.asp?PageMode=AddNew&PullupData=AdmissionID||" & adm & "&TransProcessVal2ID=AdmissionPro-T010"
                End If
            End If
            navPop = "POP"
            inout = "IN"
            fntSize = "9"
            fntColor = "#ff0000"
            bgColor = clr
            wdth = ""
            If lnkUrl <> "-" Then
                response.write "<td align=""left"">"
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
            End If
            
            If (UCase(mdO) = "T004") Or (UCase(mdO) = "T005") Or (UCase(mdO) = "T006") Then 'Bill at Social Welfare
                response.write "<td align=""left"">"
                'Clickable Url Link
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>View Soc.Welfare</b>"
                lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=AdmissionRCP&PositionForTableName=Admission&AdmissionID=" & adm
                navPop = "POP"
                inout = "IN"
                fntSize = "9"
                fntColor = "#ff0000"
                bgColor = clr
                wdth = ""
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
            End If
        ElseIf UCase(dispTyp) = "WARDNURSE" Then 'Not ward
            If UCase(mdO) = "T010" Then 'Bill is Completed
                If dichPat Then
                    response.write "<td align=""left""" & clss & ">"
                    'Clickable Url Link
                    lnkCnt = lnkCnt + 1
                    lnkID = "lnk" & CStr(lnkCnt)
                    lnkText = "<b>Discharge Pat.</b>"
                    lnkUrl = "wpgAdmissionPro.asp?PageMode=AddNew&PullupData=AdmissionID||" & adm & "&TransProcessVal2ID=AdmissionPro-T003"
                    navPop = "POP"
                    inout = "IN"
                    fntSize = "9"
                    fntColor = "#6666cc"
                    bgColor = clr
                    wdth = ""
                    AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                    response.write "</td>"
                Else
                    response.write "<td align=""left"">"
                    'Clickable Url Link
                    lnkCnt = lnkCnt + 1
                    lnkID = "lnk" & CStr(lnkCnt)
                    lnkText = "<b>Refer Soc.Welfare</b>"
                    lnkUrl = "wpgAdmissionPro.asp?PageMode=AddNew&PullupData=AdmissionID||" & adm & "&TransProcessVal2ID=AdmissionPro-T004"
                    navPop = "POP"
                    inout = "IN"
                    fntSize = "9"
                    fntColor = "#ff0000"
                    bgColor = clr
                    wdth = ""
                    'AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth '@ken hide refer social welfare
                    response.write "</td>"
                End If
            ElseIf (UCase(mdO) = "T004") Or (UCase(mdO) = "T005") Or (UCase(mdO) = "T006") Then 'Bill at Social Welfare
                response.write "<td align=""left"">"
                'Clickable Url Link
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>View Soc.Welfare</b>"
                lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=AdmissionRCP&PositionForTableName=Admission&AdmissionID=" & adm
                navPop = "POP"
                inout = "IN"
                fntSize = "9"
                fntColor = "#ff0000"
                bgColor = clr
                wdth = ""
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
            End If
        ElseIf (UCase(dispTyp) = "SOCIALWELFARE") Or (UCase(dispTyp) = "MEDICALDIRECTOR") Or (UCase(dispTyp) = "ACCOUNTANT") Then 'Not ward
            If (UCase(mdO) = "T004") Or (UCase(mdO) = "T005") Or (UCase(mdO) = "T006") Then 'Bill at Social Welfare
            response.write "<td align=""left"">"
            'Clickable Url Link
            lnkCnt = lnkCnt + 1
            lnkID = "lnk" & CStr(lnkCnt)
            lnkText = "<b>View Soc.Welfare</b>"
            lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=AdmissionRCP&PositionForTableName=Admission&AdmissionID=" & adm
            navPop = "POP"
            inout = "IN"
            fntSize = "9"
            fntColor = "#ff0000"
            bgColor = clr
            wdth = ""
            AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            response.write "</td>"
            End If
        End If
        
        response.write "</tr>"
        response.flush
        rst2.MoveNext
        Loop
    End If
    rst2.Close
    response.write "</table>"
    response.write "</td>"
    response.write "</tr>"
    response.write "</table>"
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
    Dim jb, outC, ot, tmp
    jb = Trim(jSchd)
    lstWhCls2 = " and (a.TransProcessStatID='T009' or a.TransProcessStatID='T010' or a.TransProcessStatID='T013' or a.TransProcessStatID='T004' or a.TransProcessStatID='T005' or a.TransProcessStatID='T006')"
    wdNm = Trim(GetComboName("Ward", jb))
    If Len(wdNm) > 0 Or isWard(jb) Then 'Ward Profile
    
        'lstWhCls = " and a.WardID='" & jb & "'"
        lstWhCls = " and 1 = 1 "
        
        tmp = UCase(Replace(jSchd, "W0", "W", 1, -1, 1))
        If Len(GetComboName("Ward", tmp)) > 0 Then
            If tmp = "W09" Then 'HDU
                tmp = "W10"
                lstWhCls = " and a.WardID='" & tmp & "' "
            ElseIf tmp = "W10" Then 'ICU
                tmp = "W09"
                lstWhCls = " and a.WardID='" & tmp & "' "
            ElseIf tmp = "W07" Or tmp = "W05" Then 'gynae / maternity
                lstWhCls = " and a.WardID in ('W07', 'W05') "
            ElseIf tmp = "W08" Or tmp = "W09" Then 'theatre /hdu/recovery
                lstWhCls = " and a.BranchID='" & brnch & "'"
            ElseIf tmp = "W06" Then 'male ward/annex
                lstWhCls = " and a.WardID in ('W06', 'W11') "
            ElseIf tmp = "W04" Then 'female ward/annex
                lstWhCls = " and a.WardID in ('W04', 'W12') "
            Else
                lstWhCls = " and a.WardID='" & tmp & "' "
            End If
        Else
            lstWhCls = " and 1=0 "
        End If
        
    ElseIf UCase(jb) = "M27076" Then 'Medical Director
        lstWhCls2 = " and (a.TransProcessStatID='T004' or a.TransProcessStatID='T005' or a.TransProcessStatID='T006')"
    ElseIf UCase(jb) = "M27071" Then 'Social Walfare
        lstWhCls2 = " and (a.TransProcessStatID='T004' or a.TransProcessStatID='T005' or a.TransProcessStatID='T006')"
    ElseIf UCase(jb) = "M27020" Then 'Accountant
        lstWhCls2 = " and (a.TransProcessStatID='T004' or a.TransProcessStatID='T005' or a.TransProcessStatID='T006')"
    End If
End Sub

Function isWard(jb)
    Dim ot

    ot = False

    If (UCase(Left(jb, 2))) = "W0" Or (UCase(jb) = "W1") Or (UCase(jb) = "W2") Or (UCase(jb) = "W3") Then
        ot = True
    End If
    
    isWard = ot
    End Function

    Function HasPrintOutAccess(jb, prt)
    Dim rstTblSql, sql, ot
    ot = False
    Set rstTblSql = CreateObject("ADODB.Recordset")
    With rstTblSql
        sql = "select JobScheduleID from printoutalloc "
        sql = sql & " where printlayoutid='" & prt & "' and jobscheduleid='" & jb & "'"
        .open sql, conn, 3, 4
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
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
        .MoveFirst
        ot = .fields("ModuleManagerID")
        End If
        .Close
    End With
    HasModuleMgrAccess = ot
    Set rstTblSql = Nothing
    End Function
    Function GetDispType2(jb)
    Dim ot
    ot = ""
    If Len(Trim(lstWhCls)) > 0 Then
        ot = "WARDNURSE"
    ElseIf UCase(Left(jb, 3)) = "M02" Then
        ot = "WARDNURSE"
    ElseIf UCase(jb) = "BILLINGHEAD" Then
        ot = "BILLING"
    ElseIf UCase(jb) = "CLAIMMANAGER" Then
        ot = "BILLING"
    ElseIf UCase(jb) = "M13" Then
        ot = "BILLING"
    ElseIf UCase(jb) = "M27071" Then
        ot = "SOCIALWELFARE"
    ElseIf UCase(jb) = "M27076" Then
        ot = "MEDICALDIRECTOR"
    ElseIf UCase(jb) = "M27020" Then
        ot = "ACCOUNTANT"
    End If
    GetDispType2 = ot
    End Function
    
    Sub LoadCSS()
    Dim str
    str = ""
    str = str & "<style type='text/css' id=""styPrt"">"
    str = str & ".animTR{animation-name:trBlink; animation-duration: 8s; animation-iteration-count:infinite; animation-timing-function:cubic-bezier(0.25,0.1,0.25,1);}"
    str = str & "@keyframes trBlink {"
    str = str & "from {background-color: #00ff00;}"
    str = str & "to {background-color: #ccffcc;}"
    str = str & "}"
    
    str = str & ".header-container {"
    str = str & "    display: flex;"
'    str = str & "    justify-content: center;"
'    str = str & "    width: 100%;"
'    str = str & "    margin: 20px 0;"
    str = str & "}"
    str = str & ".filters {"
    str = str & "    display: flex;"
    str = str & "    flex-direction: row;"
'    str = str & "    width: 300px;"
    'str = str & "    height: 120px;"
'    str = str & "    padding: 20px;"
    str = str & "    border: 1px solid #d1d5db;"
    str = str & "    background-color: #f9fafb;"
    str = str & "    border-radius: 8px;"
    str = str & "    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);"
    str = str & "}"
    str = str & ".filter-row {"
    str = str & "    display: flex;"
'    str = str & "    margin-bottom: 15px;"
'    str = str & "    width: 100%;"
    str = str & "}"
    str = str & ".font-style {"
    str = str & "    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;"
    str = str & "    font-size: 16px;"
    str = str & "    font-weight: 600;"
    str = str & "    color: #1f2937;"
    str = str & "    text-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);"
    str = str & "    margin-bottom: 8px;"
    str = str & "    display: block;"
    str = str & "}"
    str = str & ".myselect {"
    str = str & "    width: 100%;"
    str = str & "    padding: 10px 12px;"
    str = str & "    border-radius: 6px;"
    str = str & "    border: 1px solid #d1d5db;"
    str = str & "    font-size: 14px;"
    str = str & "    font-family: 'Inter', sans-serif;"
    str = str & "    background-color: #ffffff;"
    str = str & "    background-position: right 10px center;"
    str = str & "    background-repeat: no-repeat;"
    str = str & "    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);"
    str = str & "    transition: border-color 0.2s ease-in-out, box-shadow 0.2s ease-in-out;"
    str = str & "    appearance: none;"
    str = str & "    cursor: pointer;"
    str = str & "}"
    str = str & ".myselect:hover {"
    str = str & "    border-color: #2563eb;"
    str = str & "}"
    str = str & ".myselect:focus {"
    str = str & "    outline: none;"
    str = str & "    border-color: #2563eb;"
    str = str & "    box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);"
    str = str & "}"
    str = str & ".mybutton {"
    str = str & "    margin-top: 30px;"
    str = str & "    padding: 10px 20px;"
    str = str & "    border: none;"
    str = str & "    background: linear-gradient(90deg, #2563eb, #1d4ed8);"
    str = str & "    color: #ffffff;"
    str = str & "    font-size: 14px;"
    str = str & "    font-weight: 600;"
    str = str & "    font-family: 'Inter', sans-serif;"
    str = str & "    border-radius: 6px;"
    str = str & "    cursor: pointer;"
    str = str & "    transition: transform 0.2s ease-in-out, background 0.2s ease-in-out;"
    str = str & "    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);"
    str = str & "}"
    str = str & ".mybutton:hover {"
    str = str & "    background: linear-gradient(90deg, #1d4ed8, #1e40af);"
    str = str & "    transform: scale(1.02);"
    str = str & "}"
    str = str & ".mybutton:focus {"
    str = str & "    outline: none;"
    str = str & "    box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.3);"
    str = str & "}"

    str = str & "</style>"
    
    response.write str
    End Sub
    
    Function CleanReceiptNos(recNo)
    Dim arr, ul, num, lst, ot, arr2, ul2, num2, lst2, rec, cnt, recOk
    lst = Trim(recNo)
    lst2 = ""
    cnt = 0
    If Len(lst) > 0 Then
        arr = Split(lst, ",")
        ul = UBound(arr)
        For num = 0 To ul
        rec = Trim(arr(num))
        If Len(rec) > 0 Then
            If ReceiptCharValid(rec) Then
            arr2 = Split(lst2, ",")
            ul2 = UBound(arr2)
            recOk = True
            For num2 = 0 To ul2
                If UCase(Trim(arr2(num2))) = UCase(Trim(rec)) Then
                recOk = False
                Exit For
                End If
            Next
            If recOk Then
                cnt = cnt + 1
                If cnt > 1 Then
                lst2 = lst2 & ","
                End If
                lst2 = lst2 & rec
            End If
            End If
        End If
        Next
    End If
    CleanReceiptNos = lst2
    End Function

    Function ReceiptCharValid(rec)
    Dim ot, lst, lth, num, ch, pos
    ot = True
    lst = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890/\_- "
    lth = Len(rec)
    For num = 1 To lth
        ch = Mid(rec, num, 1)
        pos = InStr(1, UCase(lst), UCase(ch))
        If pos < 1 Then
        ot = False
        Exit For
        End If
    Next
    ReceiptCharValid = ot
    End Function
    Function HasTransProcessAccess(tbl, uNm, stStg, edStg)
    Dim rst, sql, rt, jb, ot
    ot = False
    Set rst = CreateObject("ADODB.Recordset")
    rt = tbl & "-" & stStg & "-" & edStg
    jb = GetComboNameFld("SystemUser", uNm, "JobScheduleID")
    sql = "select tableid from TransProcessorAcc2 where InitialScheduleID='" & jb & "' and TransProcessRightID='" & rt & "'"
    If recExist(sql) Then
        ot = True
    Else
        sql = "select tableid from TransProcessorAcc where InitialSystemUserID='" & uNm & "' and TransProcessRightID='" & rt & "'"
        If recExist(sql) Then
        ot = True
        End If
    End If
    Set rst = Nothing
    HasTransProcessAccess = ot
    End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
