'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
SetPageVariable "AutoHidePrintControl", "1"
Dim lnkCnt, iCnt
lnkCnt = 0
response.write Glob_GetBootstrap5()
response.write Glob_DisplayHeader("REPORTS CONTROL PANEL")
' printReport jSchd
iCnt = 0
DisplayPrintLayout jSchd
DisplayReportBrowseView jSchd


Sub DisplayPrintLayout(jb)
    Dim rstPrt, sqlPrt
    Set rstPrt = CreateObject("ADODB.Recordset")
    response.write "<h3>MY REPORTS DASHBOARD</h3>"

    response.write "<table class=""table table-striped"">"
    sqlPrt = " "
    sqlPrt = sqlPrt & " SELECT prt.PrintLayoutID, alloc.JobScheduleID, prt.TableID, alloc.ItemPos, alloc.PrintDetails, prt.CodeDescription "
    sqlPrt = sqlPrt & " , prt.PrintLayoutName, prt.PrintAfterSave, prt.UserAccessibleID, prt.PrintInputFilter, prt.Description, prt.KeyPrefix "
    sqlPrt = sqlPrt & "  FROM printoutalloc alloc, printlayout prt WHERE prt.PrintLayoutID=alloc.PrintLayoutID  "
    ' sqlPrt = sqlPrt & " and prt.TableID=alloc.TableID And prt.UserAccessibleID='USA001' "
    sqlPrt = sqlPrt & " And prt.UserAccessibleID='USA001' "
    ' sqlPrt = sqlPrt & " AND alloc.JobScheduleID='" & jb & "' " '' " And prt.KeyPrefix IN ('Report','RecordKey') "
    sqlPrt = sqlPrt & " AND alloc.JobScheduleID='" & jb & "' And prt.KeyPrefix IN ('Report','RecordKey') "
    ' sqlPrt = sqlPrt & " ORDER BY alloc.TableID, alloc.ItemPos desc; " ''@bless - 23 March 2023 //Group by PerformVar13 IDs in PrintOutAlloc.ItemPos
    sqlPrt = sqlPrt & " ORDER BY alloc.ItemPos asc "

    With rstPrt
        rstPrt.open qryPro.FltQry(sqlPrt), conn, 3, 4

        If rstPrt.recordCount > 0 Then
            rstPrt.MoveFirst
            response.write "<tr>"
            ' response.write "<th>#</th>"
            response.write "<th><i class=""fa fa-mail-forward""></i></th>"
            response.write "<th>Name</th>"
            ' response.write "<th>Staff</th>"
            response.write "<th>Title </th>"
            response.write "<th>Description</th>"
            ' response.write "<th>Description </th>"
            ' response.write "<th>Time</th>"
            response.write "</tr>"
            c04Old = "-"

            Do While Not rstPrt.EOF
                'response.write rstPrt.fields("TableID")
                c01 = rstPrt.fields("PrintLayoutID")
                c02 = (rstPrt.fields("PrintLayoutName"))
                c03 = rstPrt.fields("TableID")
                c04 = rstPrt.fields("ItemPos")
                c05 = rstPrt.fields("Description")
                If Not IsNull(c05) Then c05 = Replace(c05, vbNewLine, "<br>")
                c06 = rstPrt.fields("KeyPrefix")
                c07 = rstPrt.fields("PrintDetails")
                c08 = rstPrt.fields("PrintInputFilter")
                c09 = rstPrt.fields("PrintAfterSave")
                c10 = rstPrt.fields("CodeDescription")
                If Not IsNull(c10) Then c10 = Replace(c10, vbNewLine, "<br>")
                c11 = "-"
                c12 = "-"

                ''@bless - 23 March 2023
                If UCase(c04) <> UCase(c04Old) Or UCase(c04Old) = "-" Then
                    response.write "<tr>"
                    response.write "<th>" & GetComboNameFld("PerformVar13", c04, "KeyPrefix") & "</th>"
                    response.write "<th colspan=""4"">" & GetComboName("PerformVar13", c04) & "</th>"
                    ' response.write "<th>Description </th>"
                    ' response.write "<th>Time</th>"
                    response.write "</tr>"
                    c04Old = c04
                End If

                navPop = "OPEN" 'IN'
                inout = "IN"
                fntSize = "10"
                fntColor = "#4444cc"
                bgColor = clr
                wdth = ""
                wpg = "wpgPrtPrintInputFilter.asp?"
                wpg = "wpgNavigateFrame.asp?FrameType=PrintFilter&FilterButtonLabel=Process%20Report"
                    wpg = wpg & "&ProClickType=InFrame&FilterHeight=220&FilterPerRow=2&"
                If IsNull(c08) Or IsEmpty(c08) Or Trim(c08) = "" Then
                    wpg = "wpgPrtPrintLayoutAll.asp?"
                End If
                lnkUrl = wpg & "PrintLayoutName=" & c01 & "&noglobal=1&PositionForTableName=" & c03 & "&" & c03 & "ID="
                If Not IsNull(c06) And UCase(c06) = UCase("RecordKey") Then
                    lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=ReportRCPFilter&TableID=" & c03 & "&PrintLayoutID=" & c01
                    lnkUrl = lnkUrl & "&PositionForTableName=WorkingDay&WorkingDayID=&FilterButtonLabel=Search%20Record"
                    lnkUrl = lnkUrl & "&ProClickType=InFrame&FilterHeight=110"
                End If
                response.write "<tr>"
                response.write "<td>" & rstPrt.AbsolutePosition & "</td>"
                iCnt = rstPrt.AbsolutePosition

                response.write "<td align=""left"">"
                'Clickable Url Link
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>" & c02 & "</b>"
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"

                response.write "<td>" & c05 & "</td>"
                response.write "<td>" & c10 & "</td>"
                response.write "</tr>"
                rstPrt.MoveNext
                response.Flush
            Loop
            iCnt = rstPrt.recordCount
        Else
            response.write "<tr><th colspan=""999"">" & "No Report Available Based on Criteria Provided" & "</th></tr>"
        End If
        rstPrt.Close
    End With

    response.write "</table>"
    Set rstPrt = Nothing
End Sub

Sub DisplayReportBrowseView(jb)
    Dim rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    ' response.write "<h3>ALL MY REPORTS LISTS 2</h3>"
    response.write "<table class=""table table-striped"">"

    sql = "  "
    sql = sql & " SELECT bv.BrowseViewID, bv.BrowseViewName, bv.TableID, bv.ReportGroupByID, bv.Description "
    sql = sql & "  , bv.KeyPrefix, bv.ReportInfo1, alloc.JobScheduleID, tb.TableName, tb.DisplayName, tb.DisplayName2 "
    sql = sql & "  , tb.SystemModuleID "
    sql = sql & " FROM BrowseView bv, BrowseViewAlloc alloc, Tables tb "
    sql = sql & " WHERE bv.BrowseViewID=alloc.BrowseViewID AND bv.UserAccessibleID='USA001' AND bv.TableID=tb.TableID "
    sql = sql & "  AND alloc.AddToMyReport='Yes' AND alloc.JobScheduleID='" & jb & "' " ''" --AND bv.TableID='Receipt' "
    sql = sql & "  And bv.KeyPrefix IN ('YES','Y')  "
    sql = sql & " ORDER BY bv.TableID "
    sql = sql & "  "
    sql = sql & "  "
    ' response.write sql

    iCnt = iCnt + 1
    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.recordCount > 0 Then
            rst.MoveFirst
            response.write "<tr>"
            ' response.write "<th>#</th>"
            response.write "<th><i class=""fa fa-mail-forward""></i></th>"
            response.write "<th>Name</th>"
            ' response.write "<th>Staff</th>"
            response.write "<th>Description </th>"
            response.write "<th>Detail</th>"
            ' response.write "<th>Description </th>"
            ' response.write "<th>Time</th>"
            response.write "</tr>"

            Do While Not rst.EOF
                c01 = Trim(rst.fields("BrowseViewID"))
                c02 = Trim(rst.fields("BrowseViewName"))
                c03 = Trim(rst.fields("TableID"))
                c04 = Trim(rst.fields("ReportGroupByID"))
                c05 = Trim(rst.fields("Description"))
                c06 = Trim(rst.fields("KeyPrefix"))
                c07 = Trim(rst.fields("ReportInfo1"))
                c08 = Trim(rst.fields("JobScheduleID"))
                c09 = Trim(rst.fields("TableName"))
                c10 = Trim(rst.fields("DisplayName"))
                c11 = Trim(rst.fields("DisplayName2"))
                c12 = Trim(rst.fields("SystemModuleID"))
                c13 = "-" '' rst.fields("")
                c14 = "-" '' rst.fields("")
                c15 = "-" '' rst.fields("")
                c16 = "-" '' rst.fields("")
                c17 = "-" '' rst.fields("")
                c18 = "-" '' rst.fields("")
                c19 = "-" '' rst.fields("")
                c20 = "-" '' rst.fields("")
                ' If IsNull(c06) Or IsEmpty(c06) Or UCase(c06)=UCase("No") Then '' Or Instr(c06, "No", 1)>=1 Then
                '   rst.MoveNext
                '   Loop
                ' End If

                navPop = "NAV" 'IN'
                inout = "IN"
                fntSize = "10"
                fntColor = "#4444cc"
                bgColor = clr
                wdth = ""
                wpg = "wpgBrowseViewLayout.asp?BrowseViewName="
                wpg = "wpgNavigateBrowse.asp?BrowseViewName="
                wpg = "wpgNavigateFrame.asp?FrameType=Browse&BROWSEVIEWNAME=&"
                wpg = "wpgBrowseViewLayout.asp?HomeDispType=9&OpenProp=NAV&BrowseViewName="
                lnkUrl = wpg & c01 & "&PositionForTableName=" & c03 & ""
                response.write "<tr>"
                response.write "<td>" & iCnt & "</td>"

                response.write "<td align=""left"">"
                'Clickable Url Link
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                ' lnkText = "<b>[" & c03 & "] " & c02 & "</b>"
                lnkText = "<b>" & c02 & "</b>"
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"

                response.write "<td>" & c05 & "</td>"
                ' response.write "<td>[" & c12 & " [" & c10 & "]]</td>"
                ' response.write "<td>" & c05 & "</td>"
                ' response.write "<td>" & c06 & "</td>"
                ' response.write "<td>" & c08 & "</td>"
                rst.MoveNext
                iCnt = iCnt + 1
                response.Flush
            Loop

        Else
            response.write "<tr><th colspan=""999"">" & "No Report Available for User" & "</th></tr>"
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


Sub printReport(jb1)
    Dim iCnt
    iCnt = 0
    ' reports = ""
    ' reports = reports & "wpgPrtPrintLayoutAll.asp?PrintLayoutName=rptWardRegister&PositionForTableName=WorkingDay** General Ward Register "
    ' reports = reports & "||wpgPrtPrintInputFilter.asp?PrintLayoutName=EMRConsultingRoomRegister3&PositionForTableName=WorkingDay&WorkingDayID=&EMRDataID=&EMRComponentID=** " ''
    ' reports = reports & "||wpgPrtPrintInputFilter.asp?PrintLayoutName=EMRConsultingRoomRegister3&PositionForTableName=WorkingDay&WorkingDayID=&EMRDataID=&EMRComponentID=** " ''
    ' reports = reports & "||wpgPrtPrintInputFilter.asp?PrintLayoutName=EMRConsultingRoomRegister3&PositionForTableName=WorkingDay&WorkingDayID=&EMRDataID=&EMRComponentID=** " ''
    ' reports = reports & "||wpgPrtPrintInputFilter.asp?PrintLayoutName=EMRConsultingRoomRegister3&PositionForTableName=WorkingDay&WorkingDayID=&EMRDataID=&EMRComponentID=** " ''
    ' reports = reports & "||wpgPrtPrintInputFilter.asp?PrintLayoutName=EMRConsultingRoomRegister3&PositionForTableName=WorkingDay&WorkingDayID=&EMRDataID=&EMRComponentID=** " ''
    ' reports = reports & "||wpgPrtPrintInputFilter.asp?PrintLayoutName=EMRConsultingRoomRegister3&PositionForTableName=WorkingDay&WorkingDayID=&EMRDataID=&EMRComponentID=** " ''
    ' reports = reports & "||wpgPrtPrintInputFilter.asp?PrintLayoutName=EMRConsultingRoomRegister3&PositionForTableName=WorkingDay&WorkingDayID=&EMRDataID=&EMRComponentID=** " ''
    ' reports = reports & "|| ** "
    ' reports = reports & "|| ** "
    ' reports = reports & "|| ** "
    ' reports = reports & "|| ** "

    rptList = "|| ** "
    ' rptList = rptList & "wpgPrtPrintLayoutAll.asp?PrintLayoutName=rptWardRegister&PositionForTableName=WorkingDay** General Ward Register "
    jb = GetDispType(jb1)
    Select Case UCase(jb)
        Case UCase("DOCTOR")
            rptList = rptList & GetListingDoctor()
            rptList = rptList & GetListingSpecialist(jb1)
        Case UCase("NURSE")
            rptList = rptList & GetListingNurse()
            rptList = rptList & GetListingWardDept(jb1)
        Case UCase("WARD")
            rptList = rptList & GetListingWard()
            rptList = rptList & GetListingWardDept(jb1)
        Case UCase("Account")
            rptList = rptList & GetListingAccount()
        Case UCase("Claim")
            rptList = rptList & GetListingClaim()
        Case UCase("Billing")
            rptList = rptList & GetListingBilling()
        Case UCase("PHARMACY")
            rptList = rptList & GetListingPharmacy()
        Case UCase("DISPENSING")
            rptList = rptList & GetListingPharmacy()
            rptList = rptList & GetListingDispensing()
        Case Else

    End Select

    Reports = Split(rptList, "||")
    response.write "<table class='mytable' width=""100%"" border=""1"" cellspacing=""0"" cellpadding=""3"" style=""font-family:Arial, 'Times New Roman';font-size:14px;"">"
    For Each report In Reports
        iCnt = iCnt + 1
        data = Split(Trim(report), "**")
        If UBound(data) >= 0 Then
            If Trim(data(0)) <> "" Then
                href = data(0)
                rptNm = "<b>Description Error</b>"
                If UBound(data) >= 1 Then
                    rptNm = data(1)
                End If
                ' tblName = data(0)
                ' If Trim(tblName) = "" Then
                '     tblName = "WorkingDay"
                ' End If
                response.write "<tr>"
                response.write "<td>" & iCnt & "</td>"
                response.write "<td><a href='" & href & "' target=""_self"">" & rptNm & "</a></td>"
                response.write "</tr>"
            End If
        End If
    Next
    response.write "</table>"
End Sub

Function GetDispType(jb)
    Dim ot
    ot = "OTHER"

    Select Case UCase(Left(jb, 3))
        Case UCase("M12")
            ot = UCase("STORE")
        Case UCase("M06")
            ot = UCase("PHARMACY")
        Case UCase("M03")
            ot = UCase("DOCTOR")
        Case UCase("M05")
            ot = UCase("LABORATORY")
        Case UCase("M02")
            ot = UCase("NURSE")
        Case UCase("W01")
            ot = UCase("WARD")
        Case UCase("W02")
            ot = UCase("WARD")
        Case UCase("W03")
            ot = UCase("WARD")
        Case UCase("W04")
            ot = UCase("WARD")
        Case Else
            Select Case UCase(jb)
                Case UCase("BillingHead")
                    ot = "Billing"
                Case UCase("claimManager")
                    ot = "Claim"
                Case UCase("DPT001")
                    ot = "Account"
                Case Else

            End Select
    End Select

    GetDispType = UCase(ot)
End Function

Function GetListingDoctor()
    Dim strLst
    strLst = "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    GetListingDoctor = strLst
End Function

Function GetListingSpecialist(jb1)
    Dim strLst
    strLst = "|| ** "
    Select Case UCase(jb1)
        Case UCase("M0301")
            strLst = strLst & "|| ** "
        Case UCase("M0302")
            strLst = strLst & "|| ** "
        Case UCase("M0303")
            strLst = strLst & "|| ** "
        Case UCase("M0304")
            strLst = strLst & "|| ** "
        Case UCase("M0305")
            strLst = strLst & "|| ** "
        Case UCase("M0306")
            strLst = strLst & "|| ** "
        Case UCase("M0307")
            strLst = strLst & "|| ** "
        Case UCase("M0308")
            strLst = strLst & "|| ** "
        Case UCase("M0309")
            strLst = strLst & "|| ** "
        Case UCase("M0310")
            strLst = strLst & "|| ** "
        Case UCase("M0311")
            strLst = strLst & "|| ** "
        Case UCase("M0312")
            strLst = strLst & "|| ** "
        Case UCase("M0313")
            strLst = strLst & "|| ** "
        Case UCase("M0314")
            strLst = strLst & "|| ** "
        Case UCase("M0315")
            strLst = strLst & "|| ** "
        Case UCase("M0316")
            strLst = strLst & "|| ** "
        Case UCase("M0317")
            strLst = strLst & "|| ** "
        Case UCase("M0318")
            strLst = strLst & "|| ** "
        Case UCase("M0319")
            strLst = strLst & "|| ** "
        Case UCase("M0320")
            strLst = strLst & "|| ** "
        Case UCase("M0321")
            strLst = strLst & "|| ** "
        Case UCase("M0322")
            strLst = strLst & "|| ** "
        Case UCase("M0323")
            strLst = strLst & "|| ** "
        Case UCase("M0324")
            strLst = strLst & "|| ** "
        Case UCase("M0325")
            strLst = strLst & "|| ** "
        Case UCase("M0326")
            strLst = strLst & "|| ** "
        Case UCase("M0327")
            strLst = strLst & "|| ** "
        Case UCase("M0328")
            strLst = strLst & "|| ** "
        Case UCase("M0329")
            strLst = strLst & "|| ** "
        Case UCase("M0330")
            strLst = strLst & "|| ** "
        Case Else

    End Select

    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    GetListingSpecialist = strLst
End Function

Function GetListingNurse()
    Dim strLst
    strLst = "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    GetListingNurse = strLst
End Function

Function GetListingWard()
    Dim strLst
    strLst = "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    GetListingWard = strLst
End Function

Function GetListingWardDept(jb1)
    Dim strLst
    strLst = "|| ** "
    Select Case UCase(jb1)
        Case UCase("W01")
            strLst = strLst & "|| ** "
        Case UCase("W02")
            strLst = strLst & "|| ** "
        Case UCase("W03")
            strLst = strLst & "|| ** "
        Case UCase("W04")
            strLst = strLst & "|| ** "
        Case UCase("W05")
            strLst = strLst & "|| ** "
        Case UCase("W06")
            strLst = strLst & "|| ** "
        Case UCase("W07")
            strLst = strLst & "|| ** "
        Case UCase("W08")
            strLst = strLst & "|| ** "
        Case UCase("W09")
            strLst = strLst & "|| ** "
        Case UCase("W10")
            strLst = strLst & "|| ** "
        Case UCase("W11")
            strLst = strLst & "|| ** "
        Case UCase("W12")
            strLst = strLst & "|| ** "
        Case UCase("W13")
            strLst = strLst & "|| ** "
        Case UCase("W14")
            strLst = strLst & "|| ** "
        Case UCase("W15")
            strLst = strLst & "|| ** "
        Case UCase("W16")
            strLst = strLst & "|| ** "
        Case UCase("W17")
            strLst = strLst & "|| ** "
        Case UCase("W18")
            strLst = strLst & "|| ** "
        Case UCase("W19")
            strLst = strLst & "|| ** "
        Case UCase("W20")
            strLst = strLst & "|| ** "
        Case UCase("W21")
            strLst = strLst & "|| ** "
        Case UCase("W22")
            strLst = strLst & "|| ** "
        Case UCase("W23")
            strLst = strLst & "|| ** "
        Case UCase("W24")
            strLst = strLst & "|| ** "
        Case UCase("W25")
            strLst = strLst & "|| ** "
        Case UCase("W26")
            strLst = strLst & "|| ** "
        Case UCase("W27")
            strLst = strLst & "|| ** "
        Case UCase("W28")
            strLst = strLst & "|| ** "
        Case UCase("W29")
            strLst = strLst & "|| ** "
        Case UCase("W30")
            strLst = strLst & "|| ** "
        Case UCase("W31")
            strLst = strLst & "|| ** "
        Case UCase("W32")
            strLst = strLst & "|| ** "
        Case UCase("W33")
            strLst = strLst & "|| ** "
        Case UCase("W34")
            strLst = strLst & "|| ** "
        Case UCase("W35")
            strLst = strLst & "|| ** "
        Case UCase("W36")
            strLst = strLst & "|| ** "
        Case UCase("W37")
            strLst = strLst & "|| ** "
        Case UCase("W38")
            strLst = strLst & "|| ** "
        Case UCase("W39")
            strLst = strLst & "|| ** "
        Case Else
    End Select

    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    GetListingWardDept = strLst
End Function

Function GetListingAccount()
    Dim strLst
    strLst = "|| ** "
    strLst = strLst & "||wpgPrtPrintInputFilter.asp?PrintLayoutName=rptReturnsPaediatrics&PositionForTableName=WorkingDay&WorkingDayID= **Ward Admission Returns (Inpatient) "
    strLst = strLst & "||wpgPrtPrintInputFilter.asp?PrintLayoutName=rptWardRegister&PositionForTableName=WorkingDay&WorkingDayID= **General Ward Register "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    GetListingAccount = strLst
End Function

Function GetListingClaim()
    Dim strLst
    strLst = "|| ** "
    strLst = strLst & "||wpgPrtPrintInputFilter.asp?PrintLayoutName=rptReturnsPaediatrics&PositionForTableName=WorkingDay&WorkingDayID= **Ward Admission Returns (Inpatient) "
    strLst = strLst & "||wpgPrtPrintInputFilter.asp?PrintLayoutName=rptWardRegister&PositionForTableName=WorkingDay&WorkingDayID= **General Ward Register "
    strLst = strLst & "|| ** "
    GetListingClaim = strLst
End Function

Function GetListingBilling()
    Dim strLst
    strLst = "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    GetListingBilling = strLst
End Function

Function GetListingPharmacy()
    Dim strLst
    strLst = "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    GetListingPharmacy = strLst
End Function

Function GetListingDispensing()
    Dim strLst
    strLst = "||files/bnf_81.pdf **BNF 81 (March 2021) "
    strLst = strLst & "||files/moh_gndp_stg_ed7.pdf **MOH GNDP Standard Treatment Guidelines (Edition 7) "
    strLst = strLst & "||files/martindale_ed38.pdf **Martindale The Complete Drug Reference (Edition 38) "
    strLst = strLst & "||files/eml_ghana_2017.pdf **MOH GNDP Essential Drug List (Edition 7)"
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    GetListingDispensing = strLst
End Function

Function GetListingLaboratory()
    Dim strLst
    strLst = "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    GetListingLaboratory = strLst
End Function

Function GetListingImaging()
    Dim strLst
    strLst = "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    strLst = strLst & "|| ** "
    GetListingImaging = strLst
End Function

' Function GetListing__ ()
'     Dim strLst
'     strLst = "|| ** "
'     strLst = strLst & "|| ** "
'     strLst = strLst & "|| ** "
'     strLst = strLst & "|| ** "
'     GetListing__ = strLst
' End Function


'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
