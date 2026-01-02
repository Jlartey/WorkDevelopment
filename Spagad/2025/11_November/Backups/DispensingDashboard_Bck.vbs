'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim filterDict

AddPageCSS
AddPageJS

SetPageVariable "AutoHidePrintControl", "Yes"
Set filterDict = BuildDict(Request.QueryString)
ShowPage filterDict


Function BuildDict(qryString)
    Dim ot, Key, cWDay, dt

    dt = Now
    cWDay = FormatWorkingDay(dt)
    Set ot = CreateObject("Scripting.Dictionary")
    ot.CompareMode = 1

    For Each Key In qryString
        ot.Add Key, Trim(qryString(Key))
    Next

    If Len(ot("StartDay")) = 0 Then
        ot("StartDay") = cWDay
    End If

    If Len(ot("EndDay")) = 0 Then
        ot("EndDay") = dt
        ot("EndDay") = cWDay
    End If

    If False And (UCase(ot("EndDay")) < UCase(ot("StartDay"))) Then
        ot("EndDay") = ot("StartDay")
    Else
         ot("EndDay") = ot("StartDay")
    End If
    
    If Len(ot("SystemUserID")) = 0 Then
        ot("SystemUserID") = uname
    End If
    
    Set BuildDict = ot
End Function
Sub ShowPage(filterDict)
    Dim rptGen, tmp

    'Set rptGen = New PRTGLO_RptGen
    Set rptGen = New PRTGLO_RptGen2
    
    tmp = GetMyDispenses(filterDict)
    rptGen.AddReport tmp(0), tmp(1)
   
    tmp = GetPendingDoctorPrescriptions(filterDict)
    rptGen.AddReport tmp(0), tmp(1)

    'tmp = GetDoctorPrescriptions(filterDict)
    'rptGen.AddReport tmp(0), tmp(1)

    tmp = GetUnservedDispenses(filterDict)
    rptGen.AddReport tmp(0), tmp(1)
    
    rptGen.DefaultReport = "My Dispenses"
    rptGen.StyleAsDashboard = True
    rptGen.ShowReport

End Sub
Function GetReportControls(rptGen, name)
    Dim queryStringKey, itemsArr, displayLabel, href, ot

    ot = "<div name=""control-panel"" style=""display:inline-block;text-align:left;"">"

    queryStringKey = "StartDay=" & FormatWorkingDay(Now)
    Set itemsArr = GetWorkingDays(-230)
    displayLabel = "Select Day"
    ot = ot & rptGen.CreateSelectNavigation(queryStringKey, itemsArr, displayLabel)

    queryStringKey = "SystemUserID=" & filterDict("SystemUserID")
    Set itemsArr = GetWorkingUsers(filterDict("StartDay"), filterDict("SystemUserID"))
    displayLabel = "Select User"
    ot = ot & rptGen.CreateSelectNavigation(queryStringKey, itemsArr, displayLabel)
        
    href = "wpgDrugSale.asp?PageMode=AddNew&PullUpData=VisitationID||E01"
    ot = ot & "<div class=""pad navigation-select""><span class=""link"" style=""color:#a44949"" onclick=""open_link(this)"" data-url=""" & href & """>Walk-in Dispense</span></div>"
    ot = ot & "</div>"

    GetReportControls = ot
End Function
Function GetWorkingUsers(dy, curUsr)
    Dim sql, rst, ot, stfName

    Set GetWorkingUsers = CreateObject("Scripting.Dictionary")
    GetWorkingUsers.CompareMode = 1

    Set rst = CreateObject("ADODB.RecordSet")

    sql = "select distinct DrugSale.SystemUserID, Staff.StaffName "
    sql = sql & " from DrugSale "
    sql = sql & " inner join SystemUser on SystemUser.SystemUserID=DrugSale.SystemUserID "
    sql = sql & " inner join Staff on Staff.StaffID=SystemUser.StaffID "
    sql = sql & " where DrugSale.WorkingDayID='" & dy & "' and DrugSale.BranchID='" & brnch & "'"
    sql = sql & " order by Staff.StaffName asc "

    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            GetWorkingUsers.Add rst.fields("SystemUserID").value, rst.fields("StaffName").value
            rst.MoveNext
        Loop
    End If
    If Not GetWorkingUsers.Exists(curUsr) Then
        stfName = GetComboName("Staff", (GetComboNameFld("SystemUser", curUsr, "StaffID")))
        GetWorkingUsers.Add curUsr, stfName
    End If
    If Not GetWorkingUsers.Exists(uname) Then
        stfName = GetComboName("Staff", (GetComboNameFld("SystemUser", uname, "StaffID")))
        GetWorkingUsers.Add uname, stfName
    End If
    rst.Close

End Function
Function GetWorkingDays(pastDays)
    Dim sql, rst, ot

    Set GetWorkingDays = CreateObject("Scripting.Dictionary")
    GetWorkingDays.CompareMode = 1

    Set rst = CreateObject("ADODB.RecordSet")

    sql = "select WorkingDayID, WorkingDayName "
    sql = sql & " from WorkingDay "
    sql = sql & " where datediff(day, cast(getdate() as date), cast(WorkDate as date)) >= " & pastDays
    sql = sql & "   and cast(WorkDate as date)<=cast(getdate() as date) "
    sql = sql & " order by WorkingDayID desc "

    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            GetWorkingDays.Add rst.fields("WorkingDayID").value, rst.fields("WorkingDayName").value
            rst.MoveNext
        Loop
    End If
    rst.Close

End Function
Function GetMyDispenses(filterDict)
    Dim sql, args

    sql = "select DrugSale.DrugSaleID, DrugSale.PatientID"
    sql = sql & " , (case when DrugSale.PatientID='P1' then DrugSale.DrugSaleName else Patient.PatientName end )as [PatientName]"
    sql = sql & " , (case when DrugSale.PatientID='P1' then '[P1 Walk-in] ' + DrugSale.DrugSaleName else '[' + Patient.PatientID + '] '+ Patient.PatientName end )as [Patient]"
    sql = sql & " , DrugSale.DispenseDate "
    sql = sql & " , InsuranceScheme.InsuranceSchemeName as [Scheme] "
    sql = sql & " , DrugSale.ReceiptTypeID "
    sql = sql & " , MedicalService.MedicalServiceName"
    sql = sql & " , ReceiptType.ReceiptTypeName "
    sql = sql & " , DrugSale.PostTransactionID "
    sql = sql & " , DrugSale.VisitationID "
    sql = sql & " , DrugSale.WorkingDayID, DrugSale.MainInfo2 as ReceiptNo"
    sql = sql & " , max(DrugSale.DispenseDate) over (partition by DrugSale.PatientID, iif(DrugSale.PatientID='P1', DrugSale.DrugSaleName, Patient.PatientName) ) as RecentTransDate"
    sql = sql & " from DrugSale"
    sql = sql & " left join Patient on Patient.PatientID=DrugSale.PatientID"
    sql = sql & " left join InsuranceScheme on InsuranceScheme.InsuranceSchemeID=DrugSale.InsuranceSchemeID"
    sql = sql & " left join MedicalService on MedicalService.MedicalServiceID=DrugSale.MedicalServiceID"
    sql = sql & " left join ReceiptType on ReceiptType.ReceiptTypeID=DrugSale.ReceiptTypeID"
    sql = sql & " where 1 = 1 "

    If Len(filterDict("StartDay")) > 0 Then
        sql = sql & "   and DrugSale.WorkingDayID>='" & filterDict("StartDay") & "' "
    End If
    If Len(filterDict("EndDay")) > 0 Then
        sql = sql & "   and DrugSale.WorkingDayID<='" & filterDict("EndDay") & "' "
    End If

    If Len(filterDict("SystemUserID")) > 0 Then
        sql = sql & "   and DrugSale.SystemUserID='" & filterDict("SystemUserID") & "' " 'user
    End If

    sql = sql & " order by RecentTransDate desc, PatientName, [Patient] asc, DrugSale.DispenseDate desc"

    args = "title=My Dispenses"
    args = args & ";SubGroupFields=Patient"
    args = args & ";SubTotalLabelName=*"
    args = args & ";IgnoreFromComputations=*"
    args = args & ";HiddenFields=DrugSaleID|Patient|WorkingDayID|ReceiptNo|ReceiptTypeID|PostTransactionID|VisitationID|PatientID|Scheme|PatientName|DispenseDate|ReceiptTypeName|RecentTransDate"
    args = args & "|MedicalServiceName"
    args = args & ";ExtraFields=Dispense No|Dispense Info|Dispenses|Returns|Payment Info|Refund|Control"
    'args = args & ";ExtraFields=Dispenses|Returns|Payment Info|Refund|Control"
    args = args & ";FieldFunctions=Dispense Info:GetMyDispenseInfo|Dispenses:GetMyDispenseItems|Returns:GetMyReturnsItems|Payment Info:GetMyDispensePay"
    args = args & "|Refund:GetMyDispensePayRefund|Control:GetMyDispenseCtrl|Dispense No:GetDispenseNo"
    args = args & ";ControlPanelFunction=GetReportControls"
    args = args & ";ShowSummary=No"
    GetMyDispenses = Array(sql, args)
End Function
Function GetDispenseNo(RECOBJ, fieldNAme)
    Dim ot, href
    
    If True Then
        href = "wpgDrugSale.asp?PageMode=ProcessSelect&DrugSaleID=" & RECOBJ("DrugSaleID")
        ot = ot & "<div class=""""><span class=""link"" style=""color:#a44949"" onclick=""open_link(this)"" data-url=""" & href & """>" & RECOBJ("DrugSaleID") & "</span></div>"
    End If
    GetDispenseNo = ot
End Function
Function GetMyDispenseInfo(RECOBJ, fieldNAme)
    Dim ot, href

    ot = ""

    'ot = ot & "<div class=""hdr"">"& RECOBJ("Patient") &"</div>"
    ot = ot & "<div>" & RECOBJ("MedicalServiceName") & "</div>"
    ot = ot & "<div>" & RECOBJ("ReceiptTypeName") & "<span class=""big pad"">&#9830;</span>" & RECOBJ("Scheme") & "</div>"
    ot = ot & "<div>" & FormatDateDetail(RECOBJ("DispenseDate")) & "</div>"

    If True Then
        href = "wpgDrugSale.asp?PageMode=AddNew&PullUpData=VisitationID||" & RECOBJ("VisitationID")
        If UCase(RECOBJ("PatientID")) = "P1" Then
            ot = ot & "<div class=""pad navigation-select""><span class=""link"" style=""color:#a44949"" onclick=""open_link(this)"" data-url=""" & href & """>Walk-in Dispense</span></div>"
        Else
            ot = ot & "<div class=""""><span class=""link"" style=""color:#a44949"" onclick=""open_link(this)"" data-url=""" & href & """>Dispense Drug</span></div>"
        End If
    End If

    ot = ot & "<div></div>"

    GetMyDispenseInfo = ot
End Function
Function GetMyDispenseItems(RECOBJ, fieldNAme)
    Dim sql, rst, ot, sql0, totAmt, maxRows, href

    maxRows = 5
    totAmt = 0
    Set rst = CreateObject("ADODB.RecordSet")
    sql0 = GetDrugSaleDetailsSQL(RECOBJ("DrugSaleID"), RECOBJ("WorkingDayID"), RECOBJ("WorkingDayID"))
    sql = " with DrugSaleDetails as (" & sql0 & ")"
    sql = sql & "select * "
    sql = sql & " from DrugSaleDetails"
    sql = sql & " order by DrugName asc"
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst

        ot = ot & "<div class=""hdr"">NO.</div>"
        ot = ot & "<div class=""hdr"">Item</div>"
        ot = ot & "<div class=""hdr num-cell"">Qty</div>"
        ot = ot & "<div class=""hdr num-cell"">Amt</div>"

        Do While Not rst.EOF And rst.AbsolutePosition <= maxRows
            ot = ot & "<div>" & rst.AbsolutePosition & "</div>"
            ot = ot & "<div>" & rst.fields("DrugName") & "</div>"
            ot = ot & "<div class=""num-cell"">" & rst.fields("Qty") & "</div>"
            ot = ot & "<div class=""num-cell"">" & FormatNumber(rst.fields("FinalAmt")) & "</div>"

            If IsNumeric(rst.fields("FinalAmt")) Then
                totAmt = totAmt + rst.fields("FinalAmt")
            End If
            rst.MoveNext
        Loop
        If rst.recordCount > maxRows Then
            href = "wpgDrugSale.asp?PageMode=ProcessSelect&DrugSaleID=" & RECOBJ("DrugSaleID")
            ot = ot & "<div style=""grid-column: span 4""> "
            ot = ot & "<span class=""link"" style=""color:#a44949"" onclick=""open_link(this)"" data-url=""" & href & """>+" & (rst.recordCount - maxRows) & " More...</span>"
            ot = ot & "</div>"
            Do While Not rst.EOF
                If IsNumeric(rst.fields("FinalAmt")) Then
                    totAmt = totAmt + rst.fields("FinalAmt")
                End If
                rst.MoveNext
            Loop
        End If

        ot = ot & "<div style=""grid-column: span 2;""></div><div class=""num-cell"" style=""grid-column: span 2;border-top:1px solid silver;"">GHC " & FormatNumber(totAmt) & "</div>"

        ot = "<div style=""display:grid;grid-template-columns: auto 1fr auto auto;"">" & ot & "</div>"
    End If
    rst.Close
    GetMyDispenseItems = ot
End Function
Function GetMyReturnsItems(RECOBJ, fieldNAme)
    Dim sql, rst, ot, sql0, detail, totAmt, maxRows, href

    maxRows = 5
    totAmt = 0
    Set rst = CreateObject("ADODB.RecordSet")
    sql0 = GetDrugReturnDetailsSQL(RECOBJ("DrugSaleID"))
    sql = " with DrugReturnDetails as (" & sql0 & ")"
    sql = sql & "select * "
    sql = sql & " from DrugReturnDetails"
    sql = sql & " order by DrugName asc"

    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst

        ot = ot & "<div class=""hdr"">NO.</div>"
        ot = ot & "<div class=""hdr"">Item</div>"
        ot = ot & "<div class=""hdr num-cell"">Qty</div>"
        ot = ot & "<div class=""hdr num-cell"">Amt</div>"

        Do While Not rst.EOF And rst.AbsolutePosition <= maxRows
            ot = ot & "<div>" & rst.AbsolutePosition & "</div>"
            ot = ot & "<div>" & rst.fields("DrugName") & "</div>"
            ot = ot & "<div class=""num-cell"">" & rst.fields("ReturnQty") & "</div>"
            ot = ot & "<div class=""num-cell"">" & FormatNumber(rst.fields("FinalAmt")) & "</div>"

            If IsNumeric(rst.fields("FinalAmt")) Then
                totAmt = totAmt + rst.fields("FinalAmt")
            End If
            rst.MoveNext
        Loop
        If rst.recordCount > maxRows Then
            href = "wpgDrugReturn.asp?PageMode=ProcessSelect&DrugReturnID=" & RECOBJ("DrugReturnID")
            ot = ot & "<div style=""grid-column: span 4"">+" & (rst.recordCount - maxRows) & " More...</div>"
            'ot = ot & "<span class=""link"" style=""color:#a44949"" onclick=""open_link(this)"" data-url=""" & href & """>+" & (rst.RecordCount - maxRows) & " More...</span>"
            'ot = ot & "</div>"
            Do While Not rst.EOF
                If IsNumeric(rst.fields("FinalAmt")) Then
                    totAmt = totAmt + rst.fields("FinalAmt")
                End If
                rst.MoveNext
            Loop
        End If

        ot = ot & "<div style=""grid-column: span 2;""></div><div class=""num-cell"" style=""grid-column: span 2;border-top:1px solid silver;"">GHC " & FormatNumber(totAmt) & "</div>"

        ot = "<div style=""display:grid;grid-template-columns: auto 1fr auto auto;"">" & ot & "</div>"
    End If
    rst.Close
    If True Then
        href = "wpgDrugReturn.asp?PageMode=AddNew&VisitationID=" & RECOBJ("VisitationID")
        href = href & "&PullUpData=DrugSaleID||" & RECOBJ("DrugSaleID")
        ot = ot & "<div class=""pad""><span class=""link"" style=""color:#a44949"" onclick=""open_link(this)"" data-url=""" & href & """>Return Drugs(s)</span></div>"
    End If
    GetMyReturnsItems = ot
End Function
Function GetMyDispenseCtrl(RECOBJ, fieldNAme)
    Dim ot, rst, sql, href, debt

    ot = ""
    If RECOBJ("ReceiptTypeID") = "R001" Then
        sql = "select sum(BillAmt3) as [unpaid] from PatientBill where KeyPrefix='" & RECOBJ("DrugSaleID") & "' or  KeyPrefix='" & RECOBJ("DrugSaleID") & "-DRG2'"
        Set rst = CreateObject("ADODB.RecordSet")
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.recordCount > 0 Then
            rst.MoveFirst
            If IsNumeric(rst.fields("unpaid")) Then
                debt = rst.fields("unpaid")
            End If
        End If
    End If

    If debt > 0.001 And Not OnAdmission(RECOBJ("VisitationID")) Then 'marginal errors, unpaid
        ot = ot & "<div class=""pad""><span class=""link"" style=""color:#a44949;"" onclick=""issue_invoice(this)"" "
        ot = ot & " data-kfld=""DrugSaleID"" data-tbl=""DrugSale"" "
        ot = ot & " data-billid=""" & RECOBJ("DrugSaleID") & """ data-vst=""" & RECOBJ("VisitationID") & """>Issue Invoice</span></div>"
        RECOBJ(":RowClass") = "unpaid"
    Else 'paid
        If UCase(RECOBJ("PostTransactionID")) = "P001" Then 'not posted
            ot = ot & "<div class=""pad""><span class=""link"" style=""color:#a44949;"" onclick=""post_sales(this)"" data-drug-sale-id=""" & RECOBJ("DrugSaleID") & """ >Mark as Served</span></div>"
            RECOBJ(":RowClass") = "not-posted"
        ElseIf UCase(RECOBJ("PostTransactionID")) = "P002" Then 'posted
            ot = ot & "<div class=""pad"" style=""color:#2d831e;font-weight:bold;"" >SERVED</div> "
            RECOBJ(":RowClass") = "posted"
        End If
        href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionForm1&PositionForTableName=DrugSale&DrugSaleID=" & RECOBJ("DrugSaleID")
        ot = ot & "<div class=""pad""><span class=""link"" style=""color:#0a95ff"" onclick=""open_link(this)"" data-url=""" & href & """>Pickup List</span></div>"
    End If

    If UCase(RECOBJ("PatientID")) <> "P1" Then
        href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation&FullScreen=Yes&VisitationID=" & RECOBJ("VisitationID")
        'href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation&VisitationID=" & RECOBJ("VisitationID")
        ot = ot & "<div class=""pad""><span class=""link"" style=""color:#0a95ff"" onclick=""open_link(this, true)"" data-url=""" & href & """>Open Folder</span></div>"
    End If

    GetMyDispenseCtrl = ot
End Function
Function GetMyDispensePayRefund(RECOBJ, fieldNAme)
    Dim sql, rst, recId, rec, tmp, ot, totAmt, maxRows

    maxRows = 5
    If Not IsNull(RECOBJ("ReceiptNo")) Then
        For Each rec In Split(RECOBJ("ReceiptNo"), ",")
            tmp = Trim(rec)
            If Len(tmp) > 0 Then
                If Len(recId) > 0 Then recId = recId & ", "
                recId = recId & " '" & tmp & "' "
            End If
        Next
        If Len(recId) > 0 Then
            sql = "select * from Receipt where ReceiptID in (" & recId & ") and PaidAmount>0"
            Set rst = CreateObject("ADODB.RecordSet")
            rst.open qryPro.FltQry(sql), conn, 3, 4
            If rst.recordCount > 0 Then
                ot = ot & "<div class=""hdr"">NO.</div>"
                ot = ot & "<div class=""hdr"">REC#</div>"
                ot = ot & "<div class=""hdr num-cell"">Amt</div>"

                Do While Not rst.EOF And rst.AbsolutePosition <= maxRows
                    ot = ot & "<div>" & rst.AbsolutePosition & "</div>"
                    ot = ot & "<div>" & rst.fields("ReceiptID") & "</div>"
                    ot = ot & "<div class=""num-cell"">" & FormatNumber(rst.fields("PaidAmount")) & "</div>"

                    If IsNumeric(rst.fields("PaidAmount")) Then
                        totAmt = totAmt + rst.fields("PaidAmount")
                    End If
                    rst.MoveNext
                Loop
                If rst.recordCount > maxRows Then
                    ot = ot & "<div style=""grid-column: span 3"">+" & (rst.recordCount - maxRows) & " More...</div>"
                    Do While Not rst.EOF
                        If IsNumeric(rst.fields("PaidAmount")) Then
                            totAmt = totAmt + rst.fields("PaidAmount")
                        End If
                        rst.MoveNext
                    Loop
                End If

                ot = ot & "</div><div style=""grid-column: span 3;border-top:1px solid silver;"" class=""num-cell"">GHC " & FormatNumber(totAmt) & "</div>"

                ot = "<div style=""display:grid;grid-template-columns: auto 1fr auto;"" >" & ot & "</div>"
            End If
        End If
    End If

    GetMyDispensePayRefund = ot
End Function
Function GetMyDispensePay(RECOBJ, fieldNAme)
    Dim sql, rst, recId, rec, tmp, ot, totAmt, maxRows

    maxRows = 5
    If Not IsNull(RECOBJ("ReceiptNo")) Then
        For Each rec In Split(RECOBJ("ReceiptNo"), ",")
            tmp = Trim(rec)
            If Len(tmp) > 0 Then
                If Len(recId) > 0 Then recId = recId & ", "
                recId = recId & " '" & tmp & "' "
            End If
        Next
        If Len(recId) > 0 Then
            sql = "select * from Receipt where ReceiptID in (" & recId & ") "
            Set rst = CreateObject("ADODB.RecordSet")
            rst.open qryPro.FltQry(sql), conn, 3, 4
            If rst.recordCount > 0 Then
                ot = ot & "<div class=""hdr"">NO.</div>"
                ot = ot & "<div class=""hdr"">REC#</div>"
                ot = ot & "<div class=""hdr num-cell"">Amt</div>"

                Do While Not rst.EOF And rst.AbsolutePosition <= maxRows
                    ot = ot & "<div>" & rst.AbsolutePosition & "</div>"
                    ot = ot & "<div>" & rst.fields("ReceiptID") & "</div>"
                    ot = ot & "<div class=""num-cell"">" & FormatNumber(rst.fields("ReceiptAmount1")) & "</div>"

                    If IsNumeric(rst.fields("ReceiptAmount1")) Then
                        totAmt = totAmt + rst.fields("ReceiptAmount1")
                    End If
                    rst.MoveNext
                Loop
                If rst.recordCount > maxRows Then
                    ot = ot & "<div style=""grid-column: span 3"">+" & (rst.recordCount - maxRows) & " More...</div>"
                    Do While Not rst.EOF
                        If IsNumeric(rst.fields("ReceiptAmount1")) Then
                            totAmt = totAmt + rst.fields("ReceiptAmount1")
                        End If
                        rst.MoveNext
                    Loop
                End If

                ot = ot & "</div><div style=""grid-column: span 3;border-top:1px solid silver;"" class=""num-cell"">GHC " & FormatNumber(totAmt) & "</div>"

                ot = "<div style=""display:grid;grid-template-columns: auto 1fr auto;"" >" & ot & "</div>"
            End If
        End If
    End If

    GetMyDispensePay = ot
End Function
Function GetMyDispensesDetail(filterDict)
    Dim sql, args

    sql = " with DrugSaleDetails as ("
    sql = sql & " " & GetDrugSaleDetailsSQL("", filterDict("StartDay"), filterDict("endDay"))
    sql = sql & " )"
    sql = sql & " select DrugSaleDetails.DrugSaleID, DrugSaleDetails.PatientID"
    sql = sql & " , (case when DrugSaleDetails.PatientID='P1' then DrugSaleDetails.DrugSaleName else Patient.PatientName end )as [PatientName]"
    sql = sql & " from DrugSaleDetails"
    sql = sql & " left join Patient on Patient.PatientID=DrugSaleDetails.PatientID"
    sql = sql & " where 1 = 1 "

    sql = sql & " order by PatientName asc"

    args = "title=My Dispenses Detail"
    args = args & ";SubGroupFields=PatientName"
    args = args & ";SubTotalLabelName=*"
    args = args & ";IgnoreFromComputations=*"
    args = args & ";ExtraFields=Dispense Info|Dispenses|Returns|Payment Info|Refund|Control"
    args = args & ";ShowSummary=No"

    GetMyDispensesDetail = Array(sql, args)
End Function
Function GetDrugReturnDetailsSQL(drugsaleid)
    Dim sql

    sql = " select RetDetail.DrugSaleID, RetDetail.DrugID, RetDetail.PatientID, RetDetail.DrugSaleName, RetDetail.DrugName"
    sql = sql & "   , sum(RetDetail.ReturnQty) as ReturnQty, avg(RetDetail.UnitCost) as UnitCost, sum(RetDetail.FinalAmt) as FinalAmt "
    sql = sql & " from ("
    sql = sql & "   select DrugReturnItems.DrugSaleID, DrugReturnItems.DrugID, DrugSale.PatientID, DrugSale.DrugSaleName, Drug.DrugName"
    sql = sql & "   , DrugReturnItems.ReturnQty, DrugReturnItems.UnitCost, DrugReturnItems.FinalAmt"
    sql = sql & "   from DrugReturnItems"
    sql = sql & "   inner join DrugSale on DrugSale.DrugSaleID=DrugReturnItems.DrugSaleID"
    sql = sql & "   inner join Drug on Drug.DrugID=DrugReturnItems.DrugID"
    sql = sql & "   where 1=1 "
    sql = sql & " and DrugSale.DrugSaleID='" & drugsaleid & "'  "
    sql = sql & " "
    sql = sql & "   union all "
    sql = sql & " "
    sql = sql & "   select DrugReturnItems2.DrugSaleID, DrugReturnItems2.DrugID, DrugSale.PatientID, DrugSale.DrugSaleName, Drug.DrugName"
    sql = sql & "   , DrugReturnItems2.ReturnQty, DrugReturnItems2.UnitCost, DrugReturnItems2.MainItemValue1 as FinalAmt"
    sql = sql & "   from DrugReturnItems2"
    sql = sql & "   inner join DrugSale on DrugSale.DrugSaleID=DrugReturnItems2.DrugSaleID"
    sql = sql & "   inner join Drug on Drug.DrugID=DrugReturnItems2.DrugID"
    sql = sql & "   where 1=1 "
    sql = sql & " and DrugSale.DrugSaleID='" & drugsaleid & "'  "
    sql = sql & " "
    sql = sql & " ) as RetDetail"
    sql = sql & " group by RetDetail.DrugSaleID, RetDetail.DrugID, RetDetail.PatientID, RetDetail.DrugSaleName, RetDetail.DrugID, RetDetail.DrugName"

    GetDrugReturnDetailsSQL = sql
End Function
Function GetDrugSaleDetailsSQL(drugsaleid, startDay, endDay)
    Dim sql

    sql = " select DispDetail.DrugSaleID, DispDetail.DrugID, DispDetail.PatientID, DispDetail.DrugSaleName, DispDetail.DrugName"
    sql = sql & "   , sum(DispDetail.Qty) as Qty, avg(DispDetail.UnitCost) as UnitCost, sum(DispDetail.FinalAmt) as FinalAmt "
    sql = sql & " from ("
    sql = sql & "   select DrugSaleItems.DrugSaleID, DrugSaleItems.DrugID, DrugSale.PatientID, DrugSale.DrugSaleName, Drug.DrugName"
    sql = sql & "   , DrugSaleItems.Qty, DrugSaleItems.UnitCost, DrugSaleItems.FinalAmt"
    sql = sql & "   from DrugSaleItems"
    sql = sql & "   inner join DrugSale on DrugSale.DrugSaleID=DrugSaleItems.DrugSaleID"
    sql = sql & "   inner join Drug on Drug.DrugID=DrugSaleItems.DrugID"
    sql = sql & "   where 1=1 "
    If Len(drugsaleid) > 0 Then sql = sql & " and DrugSale.DrugSaleID='" & drugsaleid & "'  "
    If Len(startDay) > 0 Then sql = sql & " and DrugSale.WorkingDayID>='" & startDay & "'  "
    If Len(endDay) > 0 Then sql = sql & " and DrugSale.WorkingDayID<='" & endDay & "'   "
    sql = sql & " "
    sql = sql & "   union all "
    sql = sql & " "
    sql = sql & "   select DrugSaleItems2.DrugSaleID, DrugSaleItems2.DrugID, DrugSale.PatientID, DrugSale.DrugSaleName, Drug.DrugName"
    sql = sql & "   , DrugSaleItems2.DispenseAmt1 as Qty, DrugSaleItems2.UnitCost, DrugSaleItems2.DispenseAmt2 as FinalAmt"
    sql = sql & "   from DrugSaleItems2"
    sql = sql & "   inner join DrugSale on DrugSale.DrugSaleID=DrugSaleItems2.DrugSaleID"
    sql = sql & "   inner join Drug on Drug.DrugID=DrugSaleItems2.DrugID"
    sql = sql & "   where 1=1 "
    If Len(drugsaleid) > 0 Then sql = sql & " and DrugSale.DrugSaleID='" & drugsaleid & "'  "
    If Len(startDay) > 0 Then sql = sql & " and DrugSale.WorkingDayID>='" & startDay & "'  "
    If Len(endDay) > 0 Then sql = sql & " and DrugSale.WorkingDayID<='" & endDay & "'   "
    sql = sql & " "
    sql = sql & " ) as DispDetail"
    sql = sql & " group by DispDetail.DrugSaleID, DispDetail.DrugID, DispDetail.PatientID, DispDetail.DrugSaleName, DispDetail.DrugID, DispDetail.DrugName"

    GetDrugSaleDetailsSQL = sql
End Function
Function GetDoctorPrescriptions(filterDict)
    Dim sql, args

    sql = "select distinct Prescription.PatientID, Prescription.KeyPrefix as PrescriptionNo, Prescription.SpecialistID"
    sql = sql & " , '[' + Prescription.PatientID + '] ' + Patient.PatientName as [PatientName]"
    sql = sql & " , max(Prescription.PrescriptionDate) over(partition by Prescription.PatientID) as RecentTransDate"
    sql = sql & " , Visitation.VisitationID, ReceiptType.ReceiptTypeName "
    sql = sql & " , Sponsor.SponsorName as Scheme, SpecialistGroup.SpecialistGroupName"
    sql = sql & " , SpecialistType.SpecialistTypeName"
    sql = sql & " from Prescription "
    sql = sql & " left join DrugSale on DrugSale.DrugSaleID="
    sql = sql & " inner join Visitation on Visitation.VisitationID=Prescription.VisitationID"
    sql = sql & " inner join Patient on Patient.PatientID=Visitation.PatientID"
    sql = sql & " inner join ReceiptType on ReceiptType.ReceiptTypeID=Visitation.ReceiptTypeID"
    sql = sql & " inner join Sponsor on Sponsor.SponsorID=Visitation.SponsorID"
    sql = sql & " inner join SpecialistType on SpecialistType.SpecialistTypeID=Visitation.SpecialistTypeID"
    sql = sql & " inner join SpecialistGroup on SpecialistGroup.SpecialistGroupID=Visitation.SpecialistGroupID"
    sql = sql & " where 1=1 and (Prescription.PrescriptionStatusID='P001' and Prescription.TransProcessValID<>'PrescriptionPro-T003') "
    If Len(filterDict("StartDay")) > 0 Then
        sql = sql & " and Prescription.WorkingDayID >='" & filterDict("StartDay") & "' "
    End If
    If Len(filterDict("EndDay")) > 0 Then
        sql = sql & " and Prescription.WorkingDayID <='" & filterDict("EndDay") & "' "
    End If
    sql = sql & "   and Prescription.BranchID='" & brnch & "'"
    sql = sql & " order by RecentTransDate desc" ', Patient.PatientName asc, Prescription.PatientID asc"

    args = "title=Doctor Prescriptions"
    args = args & ";SubGroupFields=PatientName"
    args = args & ";SubTotalLabelName=*"
    args = args & ";IgnoreFromComputations=*"
    args = args & ";HiddenFields=PatientID|PatientName|SpecialistID|RecentTransDate|PrescriptionNo|VisitationID|Scheme|ReceiptTypeName|SpecialistTypeName|SpecialistGroupName"
    args = args & ";ExtraFields=Visit Detail|Prescription Detail|Dispense Detail|Control"
    args = args & ";FieldFunctions=Prescription Detail:GetPrescriptionDetail|Control:GetPrescriptionControl|Visit Detail:GetPrescriptionVisitDetail"
    args = args & ";ControlPanelFunction=GetReportControls"
    args = args & ";ShowSummary=No"

    GetDoctorPrescriptions = Array(sql, args)
End Function
Function GetPendingDoctorPrescriptions(filterDict)
    Dim sql, args

    sql = "select distinct Prescription.PatientID, Prescription.KeyPrefix as PrescriptionNo, Prescription.SpecialistID"
    sql = sql & " , '[' + Prescription.PatientID + '] ' + Patient.PatientName as [PatientName]"
    sql = sql & " , max(Prescription.PrescriptionDate) over(partition by Prescription.PatientID) as RecentTransDate"
    sql = sql & " , Visitation.VisitationID, ReceiptType.ReceiptTypeName "
    sql = sql & " , Sponsor.SponsorName as Scheme, SpecialistGroup.SpecialistGroupName"
    sql = sql & " , SpecialistType.SpecialistTypeName"
    sql = sql & " from Prescription "
    sql = sql & " inner join Visitation on Visitation.VisitationID=Prescription.VisitationID"
    sql = sql & " inner join Patient on Patient.PatientID=Visitation.PatientID"
    sql = sql & " inner join ReceiptType on ReceiptType.ReceiptTypeID=Visitation.ReceiptTypeID"
    sql = sql & " inner join Sponsor on Sponsor.SponsorID=Visitation.SponsorID"
    sql = sql & " inner join SpecialistType on SpecialistType.SpecialistTypeID=Visitation.SpecialistTypeID"
    sql = sql & " inner join SpecialistGroup on SpecialistGroup.SpecialistGroupID=Visitation.SpecialistGroupID"
    sql = sql & " where 1=1 and (Prescription.PrescriptionStatusID='P001' and Prescription.TransProcessValID<>'PrescriptionPro-T003') "
    If Len(filterDict("StartDay")) > 0 Then
        sql = sql & " and Prescription.WorkingDayID >='" & filterDict("StartDay") & "' "
    End If
    If Len(filterDict("EndDay")) > 0 Then
        sql = sql & " and Prescription.WorkingDayID <='" & filterDict("EndDay") & "' "
    End If
    sql = sql & "   and Prescription.BranchID='" & brnch & "'"
    sql = sql & " order by RecentTransDate desc" ', Patient.PatientName asc, Prescription.PatientID asc"

    args = "title=Pending Doctor Prescriptions"
    args = args & ";SubGroupFields=PatientName"
    args = args & ";SubTotalLabelName=*"
    args = args & ";IgnoreFromComputations=*"
    args = args & ";HiddenFields=PatientID|PatientName|SpecialistID|RecentTransDate|PrescriptionNo|VisitationID|Scheme|ReceiptTypeName|SpecialistTypeName|SpecialistGroupName"
    args = args & ";ExtraFields=Visit Detail|Prescription Detail|Control"
    args = args & ";FieldFunctions=Prescription Detail:GetPrescriptionDetail|Control:GetPrescriptionControl|Visit Detail:GetPrescriptionVisitDetail"
    args = args & ";ControlPanelFunction=GetReportControls"
    args = args & ";ShowSummary=No"

    GetPendingDoctorPrescriptions = Array(sql, args)
End Function
Function GetPrescriptionDetail(RECOBJ, fieldNAme)
    Dim sql, rst, ot, maxItems, href

    maxItems = 5
    ot = ""
    ot = ot & "<div class='hdr'>No</div>"
    ot = ot & "<div class='hdr'>Drug</div>"
    ot = ot & "<div class='hdr'>Detail</div>"
    ot = ot & "<div class='hdr'>Status</div>"

    Set rst = CreateObject("ADODB.RecordSet")

    sql = "select Prescription.DrugID, Drug.DrugName, Prescription.PrescInfo1 as [Details], Prescription.PrescriptionStatusID "
    sql = sql & " , PrescriptionStatus.PrescriptionStatusName"
    sql = sql & " from Prescription"
    sql = sql & " inner join Drug on Drug.DrugID=Prescription.DrugID"
    sql = sql & " inner join PrescriptionStatus on PrescriptionStatus.PrescriptionStatusID=Prescription.PrescriptionStatusID"
    sql = sql & " where Prescription.KeyPrefix='" & RECOBJ("PrescriptionNo") & "' "
    sql = sql & "  and (Prescription.PrescriptionStatusID='P001' and Prescription.TransProcessValID<>'PrescriptionPro-T003') "
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF And rst.AbsolutePosition <= maxItems

            ot = ot & "<div>" & rst.AbsolutePosition & "</div>"
            ot = ot & "<div>" & rst.fields("DrugName") & "</div>"
            ot = ot & "<div>" & rst.fields("Details") & "</div>"
            ot = ot & "<div>" & rst.fields("PrescriptionStatusName") & "</div>"

            rst.MoveNext
        Loop
        If rst.recordCount > maxItems Then
            ot = ot & "<div style='grid-column: span 4'>+ " & (rst.recordCount - maxItems) & " More...</div>"
        End If
        href = "wpgDrugSale.asp?PageMode=AddNew&PullUpData=VisitationID||" & RECOBJ("VisitationID")
        ot = ot & "<div style='grid-column: span 4'><span class='link' data-url='" & href & "' style=""color:#a44949"" onclick='open_link(this)'>Dispense Drug(s)</span></div>"
    End If
    rst.Close
    Set rst = Nothing

    ot = "<div style='display:grid;grid-template-columns:auto 1fr 1fr auto'>" & ot & "</div>"

    GetPrescriptionDetail = ot
End Function
Function GetPrescriptionVisitDetail(RECOBJ, fieldNAme)
    Dim ot, href, pDt, sql, rst

    sql = "select top 1 PrescriptionDate from Prescription where KeyPrefix='" & RECOBJ("PrescriptionNo") & "' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        pDt = rst.fields("PrescriptionDate")
    End If
    
    ot = ""
    ot = ot & "<div>" & RECOBJ("SpecialistTypeName") & " / " & RECOBJ("SpecialistGroupName") & "</div>"
    ot = ot & "<div>" & RECOBJ("ReceiptTypeName") & "<span class=""big pad"">&#9830;</span>" & RECOBJ("Scheme") & "</div>"
    If IsDate(pDt) Then
        ot = ot & "<div> Prescribed on: " & FormatDateDetail(pDt) & "</div>"
    End If

    ot = ot & "<div></div>"
    GetPrescriptionVisitDetail = ot
End Function
Function GetPrescriptionControl(RECOBJ, fielName)
    Dim ot, href

    If UCase(RECOBJ("PatientID")) <> "P1" Then
        href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation&FullScreen=Yes&VisitationID=" & RECOBJ("VisitationID")
        'href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation&VisitationID=" & RECOBJ("VisitationID")
        ot = ot & "<div class=""pad""><span class=""link"" style=""color:#0a95ff"" onclick=""open_link(this, true)"" data-url=""" & href & """>Open Folder</span></div>"
    End If

    GetPrescriptionControl = ot
End Function
Function GetUnservedDispenses(filterDict)
    Dim sql, args
    sql = "select DrugSale.DrugSaleID, DrugSale.PatientID"
    sql = sql & " , (case when DrugSale.PatientID='P1' then DrugSale.DrugSaleName else Patient.PatientName end )as [PatientName]"
    sql = sql & " , (case when DrugSale.PatientID='P1' then '[P1 Walk-in] ' + DrugSale.DrugSaleName else '[' + Patient.PatientID + '] '+ Patient.PatientName end )as [Patient]"
    sql = sql & " , DrugSale.DispenseDate "
    sql = sql & " , InsuranceScheme.InsuranceSchemeName as [Scheme] "
    sql = sql & " , DrugSale.ReceiptTypeID "
    sql = sql & " , ReceiptType.ReceiptTypeName "
    sql = sql & " , DrugSale.PostTransactionID "
    sql = sql & " , DrugSale.VisitationID "
    sql = sql & " , DrugSale.WorkingDayID, DrugSale.MainInfo2 as ReceiptNo"
    sql = sql & " , max(DrugSale.DispenseDate) over (partition by DrugSale.PatientID, iif(DrugSale.PatientID='P1', DrugSale.DrugSaleName, Patient.PatientName) ) as RecentTransDate"
    sql = sql & " from DrugSale"
    sql = sql & " left join Patient on Patient.PatientID=DrugSale.PatientID"
    sql = sql & " left join InsuranceScheme on InsuranceScheme.InsuranceSchemeID=DrugSale.InsuranceSchemeID"
    sql = sql & " left join ReceiptType on ReceiptType.ReceiptTypeID=DrugSale.ReceiptTypeID"
    sql = sql & " where 1 = 1 and DrugSale.PostTransactionID='P001' "
    sql = sql & "   and DrugSale.BranchID='" & brnch & "'"

'    If Len(filterDict("StartDay")) > 0 Then
'        sql = sql & "   and DrugSale.WorkingDayID>='" & filterDict("StartDay") & "' "
'    End If

    If Len(filterDict("EndDay")) > 0 Then
        sql = sql & "   and DrugSale.WorkingDayID<='" & filterDict("EndDay") & "' "
    End If

    If Len(filterDict("SystemUserID")) > 0 Then
        If False Then
            sql = sql & "   and DrugSale.SystemUserID='" & filterDict("SystemUserID") & "' " 'user
        Else
            sql = sql & "   and DrugSale.JobScheduleID='" & GetComboNameFld("SystemUser", filterDict("SystemUserID"), "JobScheduleID") & "' " 'jobschedule
        End If
    End If
    sql = sql & " order by RecentTransDate desc, PatientName, [Patient] asc, DrugSale.DispenseDate desc"

    args = "title=Unserved Dispenses"
    args = args & ";SubGroupFields=Patient"
    args = args & ";SubTotalLabelName=*"
    args = args & ";IgnoreFromComputations=*"
    args = args & ";HiddenFields=DrugSaleID|Patient|WorkingDayID|ReceiptNo|ReceiptTypeID|PostTransactionID|VisitationID|PatientID|Scheme|PatientName|DispenseDate|ReceiptTypeName|RecentTransDate"
    args = args & ";ExtraFields=Dispense No|Dispense Info|Dispenses|Returns|Payment Info|Refund|Control"
    'args = args & ";ExtraFields=Dispenses|Returns|Payment Info|Refund|Control"
    args = args & ";FieldFunctions=Dispense Info:GetMyDispenseInfo|Dispenses:GetMyDispenseItems|Returns:GetMyReturnsItems|Payment Info:GetMyDispensePay|Refund:GetMyDispensePayRefund|Control:GetMyDispenseCtrl"
    args = args & "|Dispense No:GetDispenseNo"
    args = args & ";ControlPanelFunction=GetReportControls"
    args = args & ";ShowSummary=No"

    GetUnservedDispenses = Array(sql, args)
End Function
Function GetDispenseSQL(filterDict)
    Dim sql

    sql = ""

    GetDispenseSQL = sql
End Function
Sub AddPageCSS()
    
    With response
        .write "<style>"
        .write "    .num-cell{text-align:right;}"
        .write "    table[name=""rptgen-report"" i]{width:98vw!important;}"
        .write "    .hdr{font-weight:bold;}"
        .write "    .pad{padding:3px;}"
        .write "    .link{padding:2px;text-transform:none!important;font-weight:bold;cursor:pointer;}"
        If Not True Then
            .write "    tr.paid td:nth-child(2){border-left: 3px solid #FFEB3B; background-color:#2d831e1f;}"
            .write "    tr.unpaid td:nth-child(2){border-left: 3px solid #cb0a0a; background-color:#cb0a0a3b;}"
            .write "    tr.posted td:nth-child(2){border-left: 3px solid #2d831e; background-color:#2d831e45;}"
        Else
        
            .write "    tr.unpaid td * {background-color:transparent;}"
            .write "    tr.unpaid td:first-child{border-left: 3px solid #cb0a0a !important;}"
            .write "    tr.unpaid td {background-color:#FCF1F0;border-top:1px solid #ff897e;border-bottom: 1px solid #ff897e;}"
            
            .write "    tr.paid td *{background-color:transparent;}"
            .write "    tr.paid td:first-child{border-left: 3px solid #FFEB3B !important; }"
            .write "    tr.paid td{background-color:#f3fff1; border-top:1px solid #90d385;border-bottom:1px solid #90d385;}"

            .write "    tr.not-posted td *{background-color:transparent;}"
            .write "    tr.not-posted td:first-child{border-left: 3px solid #FFC107 !important}"
            .write "    tr.not-posted td{background-color:#ffc10714;border-top: 1px solid #FFEB3B;border-bottom: 1px solid #FFEB3B;}"
            
            .write "    tr.posted td *{background-color:transparent;}"
            .write "    tr.posted td:first-child{border-left: 3px solid #2d831e !important;}"
            .write "    tr.posted td{background-color:#f3fff1;border-top:1px solid #90d385;border-bottom:1px solid #90d385;}"
        End If
        .write "    div.navigation-select{display:inline-block;}"
        .write "    .big{font-size:12pt;}"
        .write "    div{padding-left:5px;padding-right:5px;}"
        .write "    div.navigation-select, div.navigation-select *{font-size:12px;}"
        .write "</style>"
    End With
End Sub
Function OnAdmission(vst)
    Dim ot, rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    ot = False
    With rst
        sql = "select visitationid from admission where visitationid='" & vst & "' and (admissionstatusid='A001' or admissionstatusid='A007')"
        .open qryPro.FltQry(sql), conn, 3, 4
        If .recordCount > 0 Then
            ot = True
        End If
        .Close
    End With
    OnAdmission = ot
    Set rst = Nothing
End Function
Sub AddPageJS()

    With response
        .write "<script>"
        .write vbCrLf & " function issue_invoice(spn){ "
        .write vbCrLf & "   let url = 'wpgXMLHttp.asp?ProcedureName=GeneratePatientInvoices&inpVisitationID=' + spn.dataset.vst; "
        .write vbCrLf & "   url += ('&TableName=' + spn.dataset.tbl); "
        .write vbCrLf & "   url += '&inp' + spn.dataset.kfld + '=' + spn.dataset.billid; "
        .write vbCrLf & "   "
        .write vbCrLf & "   spn.style.display='none';"
        .write vbCrLf & "   fetch(url)"
        .write vbCrLf & "   .then(response => response.json())"
        .write vbCrLf & "   .then( data=>{"
        .write vbCrLf & "       if(data['status'].toUpperCase() == 'PAID'){"
        .write vbCrLf & "           spn.style.display='';"
        .write vbCrLf & "           spn.innerHTML ='PAID';"
        .write vbCrLf & "           spn.onclick=null;"
        .write vbCrLf & "           spn.style.color='green';"
        .write vbCrLf & "       }"
        .write vbCrLf & "       else{"
        .write vbCrLf & "          for(tmp of data['messages']){ "
        .write vbCrLf & "                 if(tmp.length >0 ){ "
        .write vbCrLf & "                     setTimeout(function(){alert(tmp)}, 200);  "
        .write vbCrLf & "                 } "
        .write vbCrLf & "          } "
        .write vbCrLf & "          spn.style.display='';"
        .write vbCrLf & "       } "
        .write vbCrLf & "   })"
        .write vbCrLf & "  .catch(error => {console.log(error); alert('A problem occured please try again later;');spn.style.display='';})"
        .write vbCrLf & "  "
        .write vbCrLf & " }"
        .write vbCrLf & "  "
        .write vbCrLf & " function post_sales(spn){ "
        .write vbCrLf & "   let url = 'wpgXMLHttp.asp?ProcedureName=PostDrugSale&inpDrugSaleID=' + spn.dataset.drugSaleId; "
        .write vbCrLf & "   spn.style.display='none';"
        .write vbCrLf & "   fetch(url)"
        .write vbCrLf & "   .then(response => response.json())"
        .write vbCrLf & "   .then( data=>{"
        .write vbCrLf & "       if(data['status'].toUpperCase() == 'PAID'){"
        .write vbCrLf & "           spn.style.display='';"
        .write vbCrLf & "           spn.innerHTML ='SERVED';"
        .write vbCrLf & "           spn.onclick=null;"
        .write vbCrLf & "           spn.style.color='green';"
        .write vbCrLf & "           spn.dataset.parentClass = 'posted';"
        .write vbCrLf & "           highlight_parent(spn);"
        .write vbCrLf & "       }"
        .write vbCrLf & "       else{"
        .write vbCrLf & "          for(tmp of data['messages']){ "
        .write vbCrLf & "                 if(tmp.length >0 ){ "
        .write vbCrLf & "                     setTimeout(function(){alert(tmp)}, 200);  "
        .write vbCrLf & "                 } "
        .write vbCrLf & "          } "
        .write vbCrLf & "          spn.style.display=''; "
        .write vbCrLf & "       } "
        .write vbCrLf & "   }) "
        .write vbCrLf & "  .catch(error => {console.log(error); alert('A problem occured please try again later;');spn.style.display='';})"
        .write vbCrLf & "  "
        .write vbCrLf & " } "
        .write vbCrLf & " "
        .write vbCrLf & " function open_link(spn, fullScreen){"
        .write vbCrLf & "   window.open(spn.dataset.url, '_blank', (fullScreen===true?'' :'width=800px,height=700px'));"
        .write vbCrLf & " }"
        .write vbCrLf & " "
        .write vbCrLf & " "
        .write vbCrLf & " function highlight_parent(spn){"
        .write vbCrLf & "   if(!highlight_parent.afftected){ highlight_parent.affected = 0;};"
        .write vbCrLf & "   if(!highlight_parent.runCount){ highlight_parent.runCount = 0;};"
        .write vbCrLf & "   let tr = spn.closest('tr');"
        .write vbCrLf & "   if(!tr.classList.contains(spn.dataset.parentClass)){"
        .write vbCrLf & "       tr.classList.add(spn.dataset.parentClass);"
        .write vbCrLf & "       highlight_parent.affected++;"
        .write vbCrLf & "   }"
        .write vbCrLf & " "
        .write vbCrLf & " }"
        .write vbCrLf & " "
        .write vbCrLf & " "
        .write vbCrLf & " function reload_page(){ "
        .write vbCrLf & "   window.location.reload();"
        .write vbCrLf & " } "
        .write "</script>"
    End With
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>

