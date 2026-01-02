'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim arPeriod, periodStart, periodEnd, medSrvID
arPeriod = getDatePeriodFromDelim(Trim(Request.QueryString("PrintFilter0")))
periodStart = arPeriod(0)
periodEnd = arPeriod(1)
medSrvID = Trim(Request.QueryString("PrintFilter1"))
ShowReport

Sub DisplayCoverLetterL(mth)
    Dim wkDy, prtFlt, totAmt, totClaim
    Dim cellSty, cellStyB, coPayAmt, clmAmt
    Dim dt, bMthNam

    dt = CDate(GetComboName("WorkingMonth", mth))
    bMthNam = MonthName(Month(dt)) & " " & Year(dt)
    DisplayHeader
    response.write "<div style=""page-break-after:always;line-height:40px;margin-top:35px;margin-left:50px;margin-right:50px;"" class=""cover-letter"" id=""cover-letter"">"
    response.write "    <div>"
    response.write "        <div style=""display:flex;font-weight:600;"">OUR TIN:&nbsp; C000158300X</div>"
    response.write "        <div style=""display:flex;font-weight:600;"">EMAIL:&nbsp; raphal.acct@gmail.com</div>"
    response.write "        <div style=""display:flex;font-weight:600;"">TEL NO:&nbsp; 0501255824/0501607940</div>"
    response.write "        <div style=""display:flex;font-weight:600;"">Date " & FormatDate(Now()) & "</div>"
    response.write "    </div>"
    response.write "    <div style='width:75%;margin-right:auto;'>"
    response.write "        <div style=""margin:20px 0;"">TO:</div>"
    response.write "        <div style=""font-weight:600;"">" & GetRecordField("Description") & "</div>"
    response.write "        <div style=""font-weight:600;"">" & GetRecordField("SponsorName") & "</div>"
    response.write "        <div style=""font-weight:600;"">" & GetRecordField("Address") & "</div>"
    response.write "        <div style="""">Dear Sir/Madam</div>"
    response.write "        <div style=""text-decoration:underline;margin:15px 0;"">MEDICAL BILL</div>"
    response.write "        <div style="""">Please find attached the Medical Bill for the month of "
    response.write "<span style='text-decoration:underline;'>" & bMthNam & "</span> in respect of staff of your establishment and their dependants.</div> "
    response.write "        <div style="""">The amount is <span style='/*font-style:italic;*/' id='bill-amt-words'></span> (GHC <span style='font-style:italic;' id='bill-amt'></span>)</div>"
    response.write "        <div style="""">Cheques in settlement of this bill should be in the name of <span style='text-decoration:underline;'>Raphal Medical Center</span>.</div><br>"
    response.write "        <div style="""">Thank You</div><br>"
    response.write "        <div style="""">Yours faithfully</div><br>"
    response.write "        <div style=""display:flex;justify-content: space-between;width:100%;""name=""signature""> "
    response.write "            <div style=""height:70px;width:300px;border-bottom:1px solid black;"" ></div>"
    response.write "            <div style=""height:70px;width:300px;border-bottom:1px solid black;"" ></div>"
    response.write "        </div>"
    response.write "        <div style=""display:flex;justify-content: space-between;width:100%;"">"
    response.write "            <div style="""">MR. TITUS TWUMASI SEDDOH<br>FINANCE MANAGER</div>"
    response.write "            <div style="""">MR. ERIC AMEYEDOWO<br>CLAIMS MANAGER</div>"
    response.write "        </div>"
    response.write "    </div>"
    response.write "</div>"

End Sub
Sub AddSchemeSummary(amt, amtWrds)
    Dim html

    html = ""
    html = html & " <script>"
    html = html & "     addSchemeSummary( '" & amt & "', '" & amtWrds & "')"
    html = html & " "
    html = html & " </script>"
    response.write html
End Sub
Function AddStyles()
    Dim html

    html = html & "<style>"
    html = html & " .cover-letter *{min-height:40px;font-size:18px;text-align:justify;}"
    html = html & " [name=""rptgen-report""], [name=""rptgen-report""] *{text-transform:none!important;}"
    html = html & " [name=""rptgen-report""]:first-child {page-break-after:always;}"
    html = html & " [name=""rptgen-report""] {width:80%;}"
    html = html & " [name=""rptgen-report""] [name=""report-title""] th{text-align:center;}"
    html = html & " @media print{"
    html = html & "     table[name=""rptgen-report""] * {"
    html = html & "         font-size: 12pt!important;"
    html = html & "         padding: 2px !important; "
    html = html & "     } "
    html = html & "    .cover-letter *{font-size:18pt!important;}"
    html = html & "     .no-print{display:none!important;}"
    html = html & "     @page{"
    html = html & "         margin:25px!important;"
    html = html & "         margin-top:25px!important;"
    html = html & "     }"
    html = html & " }"
    html = html & "</style>"

    response.write html
End Function
Sub AddPageJS()
    Dim html
    
    html = html & "<script>"
    html = html & vbCrLf & "    function addSchemeSummary(amt, amtWrds){"
    html = html & vbCrLf & "        let amtEl = document.getElementById(""bill-amt"");"
    html = html & vbCrLf & "        let amtWrdsEl = document.getElementById(""bill-amt-words"");"
    html = html & vbCrLf & "        if(amtEl){"
    html = html & vbCrLf & "            amtEl.innerText=amt;"
    html = html & vbCrLf & "        }"
    html = html & vbCrLf & "        if(amtWrdsEl){"
    html = html & vbCrLf & "            amtWrdsEl.innerText=amtWrds;"
    html = html & vbCrLf & "        }"
    html = html & vbCrLf & "    }"
    html = html & vbCrLf & "</script>"
    
    response.write html
End Sub
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
Sub ShowBillDetails(title, spn, insSch, billMth, exWhCls)
    Dim sql, bGrpSql, rst, tmp, rptGen, fmtNumFlds, tmp2, args, colTot, amt, amtWrds
    
    sql = " with GeneratedBills as ( "
    sql = sql & GetAllBillDetailSQL()
    sql = sql & " ) "

    sql = sql & " , Diag as ("
    'sql = sql & "       select Diagnosis.VisitationID, string_agg(Disease.DiseaseName, ', ') as Diagnosis"
    sql = sql & "       select Visitation.PatientID + '||' + Visitation.WorkingDayID as PKey, string_agg(Disease.DiseaseName, ', ') as Diagnosis"
    sql = sql & "       from Diagnosis inner join Visitation on Visitation.VisitationID=Diagnosis.VisitationID"
    sql = sql & "       inner join Disease on Disease.DiseaseID=Diagnosis.DiseaseID"
    sql = sql & "       where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    '  If Len(medSrvID) > 1 Then
    '      sql = sql & "  and Diagnosis.MedicalServiceID=@medServiceID "
    '  End If
    'sql = sql & "       group by Diagnosis.VisitationID"
    sql = sql & "       group by Visitation.PatientID, Visitation.WorkingDayID "
    sql = sql & "   )"

    sql = sql & " select format(Visitation.VisitDate, 'dddd, dd MMMM yyyy') as [Date of consultation], Patient.PatientName as [Name of Patient] "
    sql = sql & "   , isnull(Principal.PatientName, Patient.PatientName) as [Name of Principal], InsuredPatient.InsuranceNo as [Insurance/Staff No]"
    sql = sql & "   , InsSchemeMode.InsSchemeModeName as [Department], Visitation.VisitationID as [Visit No]"
    sql = sql & "   , Diag.Diagnosis"
    sql = sql & "   , GeneratedBills.Description, GeneratedBills.BillGroupCatID, GeneratedBills.BillGroupCatName as [CAT]"
    sql = sql & "   , MedicalService.MedicalServiceName as [Type of Visit]"
    sql = sql & "   , sum(GeneratedBills.Qty) as Quantity, avg(GeneratedBills.UnitCost) as [Unit Price], sum(GeneratedBills.FinalCost) as [Total] "
    
    sql = sql & " from Visitation "
    sql = sql & " inner join GeneratedBills on GeneratedBills.VisitationID=Visitation.VisitationID "
    sql = sql & " left join InsuredPatient on InsuredPatient.InsuredPatientID=Visitation.InsuredPatientID "
    sql = sql & " left join InsuranceScheme on InsuranceScheme.InsuranceSchemeID=InsuredPatient.InsuranceSchemeID "
    sql = sql & " left join Sponsor on Sponsor.SponsorID=Visitation.SponsorID "
    sql = sql & " left join InsSchemeMode on InsSchemeMode.InsSchemeModeID=InsuredPatient.InsSchemeModeID "
    sql = sql & " left join BenefitType on BenefitType.BenefitTypeID=InsuredPatient.BenefitTypeID "
    sql = sql & " left join InsuredPatient as PrincipalAccount on PrincipalAccount.InsuredPatientID=InsuredPatient.InitialDependantID "
    sql = sql & " left join Patient as Principal on Principal.PatientID=PrincipalAccount.PatientID "
    sql = sql & " left join Patient on Patient.PatientID=Visitation.PatientID "
    sql = sql & " left join MedicalService on MedicalService.MedicalServiceID=Visitation.MedicalServiceID "
    sql = sql & " left join Diag on Diag.PKey=Visitation.PatientID + '||' + Visitation.WorkingDayID "
    'sql = sql & " left join Diag on Diag.VisitationID=Visitation.VisitationID "

    sql = sql & " where 1=1 "
    If exWhCls <> "" Then
        sql = sql & " and " & exWhCls
    End If
    
    sql = sql & " group by Visitation.VisitDate, Visitation.VisitationID, Patient.PatientName, InsuredPatient.InsuranceNo, Diag.Diagnosis, InsuranceScheme.InsuranceSchemeName"
    sql = sql & "   ,Principal.PatientName,BenefitType.BenefitTypeName, Sponsor.SponsorName, InsSchemeMode.InsSchemeModeName"
    sql = sql & "   , GeneratedBills.Description, GeneratedBills.BillGroupCatID, GeneratedBills.BillGroupCatName, MedicalService.MedicalServiceName"
    sql = sql & " having(sum(GeneratedBills.Qty) > 0)"
    sql = sql & " order by InsuranceScheme.InsuranceSchemeName, [Date of consultation], Patient.PatientName, Visitation.VisitationID, GeneratedBills.BillGroupCatID "
    sql = sql & "   , GeneratedBills.BillGroupCatName, GeneratedBills.Description, MedicalService.MedicalServiceName"
    
    sql = Replace(sql, "@billMonthStart", " '" & periodStart & "' ")
    sql = Replace(sql, "@billMonthEnd", " '" & periodEnd & "' ")
    sql = Replace(sql, "@medServiceID", " '" & medSrvID & "' ")
    
    Set rptGen = New PRTGLO_RptGen
    
    args = "title=" & title
    args = args & ";ShowRowTotal=No;PageRecordCount=999999999"
    args = args & ";FormatMoneyFields=UNIT PRICE|TOTAL"
    args = args & ";FormatNumberFields=QUANTITY"
    'args = args & ";IgnoreFromComputations=Date of consultation|NAME OF PATIENT|NAME OF PRINCIPAL|INSURANCE/STAFF NO|DEPARTMENT|VISIT NO|description|CAT|Quantity|Unit Price|Type of Visit"
    args = args & ";InCludeInComputations=Total"
    args = args & ";HiddenFields=Name of Patient|NAME OF PRINCIPAL|INSURANCE/STAFF NO|DEPARTMENT|BillGroupCatID"
    args = args & ";SubGroupFields=Name of Patient|NAME OF PRINCIPAL|INSURANCE/STAFF NO|DEPARTMENT|Diagnosis"
    args = args & ";SubTotalLabelName=SUB TOTAL"
    args = args & ";FieldFunctions=Visit No:GetClaimLink"
    
    Set colTot = rptGen.PrintSQLReport(sql, args)

    amt = colTot("Total")
    amtWrds = ConvertCurrencyToEnglish(amt, "Ghanaian Cedi", "Pesewa")
    AddSchemeSummary FormatNumber(amt, 2), amtWrds
End Sub
Function GetClaimLink(RECOBJ, fieldNAme)
    Dim ot
    
    href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationClaim&PositionForTableName=Visitation&VisitationID=" & RECOBJ(fieldNAme)
    ot = "<a href='" & href & "' style='text-decoration:none;' target='_blank'>" & RECOBJ(fieldNAme) & "</a>"
    GetClaimLink = ot
End Function
Function GetAllBillDetailSQL()
    Dim sql
    
    sql = ""
    sql = sql & " select Visitation.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName "
    sql = sql & "         , BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName, SpecialistType.SpecialistTypeName + ' ['+ VisitType.VisitTypeName +']' as [Description] "
    sql = sql & "         , 1 as Qty, Visitation.VisitCost as UnitCost, Visitation.VisitCost as FinalCost "
    sql = sql & "     from Visitation  "
    sql = sql & "     left join SpecialistType on SpecialistType.SpecialistTypeID=Visitation.SpecialistTypeID "
    sql = sql & "     left join VisitType on VisitType.VisitTypeID=Visitation.VisitTypeID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=SpecialistType.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=SpecialistType.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    If Len(medSrvID) > 1 Then
        sql = sql & "  and Visitation.MedicalServiceID=@medServiceID "
    End If
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select Investigation.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName"
    sql = sql & "         , BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName, LabTest.LabTestName as [Description] "
    sql = sql & "         , sum(Investigation.Qty) as Qty, avg(Investigation.UnitCost) as UnitCost, sum(Investigation.FinalAmt) as FinalCost "
    sql = sql & "     from Investigation "
    sql = sql & "     left join Visitation on Visitation.VisitationID=Investigation.VisitationID "
    sql = sql & "     left join LabTest on LabTest.LabTestID=Investigation.LabTestID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=Investigation.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=Investigation.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    ' If Len(medSrvID) > 1 Then
    '     sql = sql & "  and Investigation.MedicalServiceID=@medServiceID "
    ' End If
    sql = sql & "     group by Investigation.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName "
    sql = sql & "           , LabTest.LabTestName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select Investigation2.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName "
    sql = sql & "         , BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName, LabTest.LabTestName as [Description] "
    sql = sql & "         , sum(Investigation2.Qty) as Qty, avg(Investigation2.UnitCost) as UnitCost, sum(Investigation2.FinalAmt) as FinalCost  "
    sql = sql & "     from Investigation2 "
    sql = sql & "     left join Visitation on Visitation.VisitationID=Investigation2.VisitationID "
    sql = sql & "     left join LabTest on LabTest.LabTestID=Investigation2.LabTestID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=Investigation2.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=Investigation2.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    ' If Len(medSrvID) > 1 Then
    '     sql = sql & "  and Investigation2.MedicalServiceID=@medServiceID "
    ' End If
    sql = sql & "     group by Investigation2.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName "
    sql = sql & "           , LabTest.LabTestName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select DrugSaleItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName "
    sql = sql & "         , BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName, Drug.DrugName as [Description] "
    sql = sql & "         , sum(DrugSaleItems.Qty) as Qty, avg(DrugSaleItems.UnitCost) as UnitCost, sum(DrugSaleItems.FinalAmt) as FinalCost  "
    sql = sql & "     from DrugSaleItems "
    sql = sql & "     left join Visitation on Visitation.VisitationID=DrugSaleItems.VisitationID "
    sql = sql & "     left join Drug on Drug.DrugID=DrugSaleItems.DrugID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=DrugSaleItems.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=DrugSaleItems.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    ' If Len(medSrvID) > 1 Then
    '     sql = sql & " and DrugSaleItems.MedicalServiceID=@medServiceID "
    ' End If
    sql = sql & "     group by DrugSaleItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName"
    sql = sql & "           , Drug.DrugName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select DrugReturnItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName "
    sql = sql & "         , BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName, Drug.DrugName as [Description] "
    sql = sql & "         , -sum(DrugReturnItems.ReturnQty) as Qty, avg(DrugReturnItems.UnitCost) as UnitCost, -sum(DrugReturnItems.FinalAmt) as FinalCost  "
    sql = sql & "     from DrugReturnItems  "
    sql = sql & "     left join Visitation on Visitation.VisitationID=DrugReturnItems.VisitationID "
    sql = sql & "     left join Drug on Drug.DrugID=DrugReturnItems.DrugID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=DrugReturnItems.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=DrugReturnItems.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    '  If Len(medSrvID) > 1 Then
    '      sql = sql & " and DrugReturnItems.MedicalServiceID=@medServiceID "
    '  End If
    sql = sql & "     group by DrugReturnItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName"
    sql = sql & "           , Drug.DrugName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select DrugSaleItems2.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName "
    sql = sql & "         , BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName, Drug.DrugName as [Description] "
    sql = sql & "         , sum(DrugSaleItems2.DispenseAmt1) as Qty, avg(DrugSaleItems2.UnitCost) as UnitCost, sum(DrugSaleItems2.DispenseAmt2) as FinalCost  "
    sql = sql & "     from DrugSaleItems2 "
    sql = sql & "     left join Visitation on Visitation.VisitationID=DrugSaleItems2.VisitationID "
    sql = sql & "     left join Drug on Drug.DrugID=DrugSaleItems2.DrugID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=DrugSaleItems2.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=DrugSaleItems2.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "

    sql = sql & "     group by DrugSaleItems2.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName"
    sql = sql & "           , Drug.DrugName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select DrugReturnItems2.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName"
    sql = sql & "         , BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName, Drug.DrugName as [Description] "
    sql = sql & "         , -sum(DrugReturnItems2.ReturnQty) as Qty, avg(DrugReturnItems2.UnitCost) as UnitCost, -sum(DrugReturnItems2.MainItemValue1) as FinalCost  "
    sql = sql & "     from DrugReturnItems2  "
    sql = sql & "     left join Visitation on Visitation.VisitationID=DrugReturnItems2.VisitationID "
    sql = sql & "     left join Drug on Drug.DrugID=DrugReturnItems2.DrugID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=DrugReturnItems2.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=DrugReturnItems2.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    '  If Len(medSrvID) > 1 Then
    '      sql = sql & " and DrugReturnItems2.MedicalServiceID=@medServiceID "
    '  End If
    sql = sql & "     group by DrugReturnItems2.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName"
    sql = sql & "           , Drug.DrugName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select TreatCharges.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName"
    sql = sql & "         , BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName, Treatment.TreatmentName as [Description] "
    sql = sql & "         , sum(TreatCharges.Qty) as Qty, avg(TreatCharges.UnitCost) as UnitCost, sum(TreatCharges.FinalAmt) as FinalCost  "
    sql = sql & "     from TreatCharges  "
    sql = sql & "     left join Visitation on Visitation.VisitationID=TreatCharges.VisitationID "
    sql = sql & "     left join Treatment on Treatment.TreatmentID=TreatCharges.TreatmentID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=TreatCharges.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=TreatCharges.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    '  If Len(medSrvID) > 1 Then
    '      sql = sql & " and TreatCharges.MedicalServiceID=@medServiceID "
    '  End If
    sql = sql & "     group by TreatCharges.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName "
    sql = sql & "           , Treatment.TreatmentName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select StockIssueItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName "
    sql = sql & "         , BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName, Items.ItemName as [Description] "
    sql = sql & "         , sum(StockIssueItems.Qty) as Qty, avg(StockIssueItems.UnitCost) as UnitCost, sum(StockIssueItems.FinalAmt) as FinalCost  "
    sql = sql & "     from StockIssueItems "
    sql = sql & "     left join Visitation on Visitation.VisitationID=StockIssueItems.VisitationID "
    sql = sql & "     left join Items on Items.ItemID=StockIssueItems.ItemID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=StockIssueItems.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=StockIssueItems.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    sql = sql & "     group by StockIssueItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName"
    sql = sql & "           , Items.ItemName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select StockReturnItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName "
    sql = sql & "         , BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName, Items.ItemName as [Description] "
    sql = sql & "         , -sum(StockReturnItems.ReturnQty) as Qty, avg(StockReturnItems.UnitCost) as UnitCost, -sum(StockReturnItems.FinalAmt) as FinalCost  "
    sql = sql & "     from StockReturnItems  "
    sql = sql & "     left join Visitation on Visitation.VisitationID=StockReturnItems.VisitationID "
    sql = sql & "     left join Items on Items.ItemID=StockReturnItems.ItemID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=StockReturnItems.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=StockReturnItems.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    sql = sql & "     group by StockReturnItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName"
    sql = sql & "           , Items.ItemName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    
    GetAllBillDetailSQL = sql
End Function
Sub DisplayHeader()
    response.write "<style>div.dheader{height:300px;background:url('images/letterhead5.jpg');background-repeat:no-repeat;background-size:contain;} "
    response.write "@media print{div.dheader{content:url('images/letterhead5.jpg')}}</style>"
    
    response.write "<div class=""dheader"">"
    response.write "</div>"
End Sub
Sub ShowReport()
    Dim mth
    Dim title, spn, insSch, billMth, exWhCls

    server.scripttimeout = 1800
    
    'mth = Request.QueryString("PrintFilter")
    mth = FormatWorkingMonth(periodStart)

    AddPageJS

    DisplayCoverLetterL mth
    spn = GetRecordField("SponsorID")
    insSch = ""
    billMth = mth
    exWhCls = " InsuredPatient.InsuredPatientID<>'CANCEL' and Visitation.SponsorID='" & spn & "' "
    If Len(medSrvID) > 1 Then
        exWhCls = exWhCls & " and Visitation.MedicalServiceID='" & medSrvID & "'"
    End If
'    title = (GetRecordField("SponsorName")) & " [Medical Bills for " & GetComboName("WorkingMonth", mth) & "]"
    title = UCase(GetRecordField("SponsorName") & " MEDICAL BILLS FOR " & GetMonthName(mth))

    Call ShowBillDetails(title, spn, insSch, billMth, exWhCls)
    AddStyles
End Sub
Function GetMonthName(mth)
    Dim ot, tmp
    
    tmp = CDate(GetComboName("WorkingMonth", mth))
    
    GetMonthName = MonthName(Month(tmp)) & " " & Year(tmp)
End Function
Function ConvertCurrencyToEnglish(ByVal MyNumber, currencyName, subCurrencyName)
   Dim Temp
   Dim Dollars, Cents
   Dim DecimalPlace, count

   ReDim Place(9)

   Place(2) = " Thousand "
   Place(3) = " Million "
   Place(4) = " Billion "
   Place(5) = " Trillion "

   'Convert MyNumber to a string, trimming extra spaces.
   MyNumber = Trim(CStr(MyNumber))

   'Find decimal place.
   DecimalPlace = InStr(MyNumber, ".")

   'If we find decimal place...
   If DecimalPlace > 0 Then
      'Convert cents
      Temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
      Cents = ConvertTens(Temp)
      'Strip off cents from remainder to convert.
      MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
   End If

   count = 1

   Do While MyNumber <> ""
      'Convert last 3 digits of MyNumber to English dollars.
      Temp = ConvertHundreds(Right(MyNumber, 3))
      If Temp <> "" Then
          Dollars = Temp & Place(count) & ", " & Dollars
      End If
      If Len(MyNumber) > 3 Then
         'Remove last 3 converted digits from MyNumber.
         MyNumber = Left(MyNumber, Len(MyNumber) - 3)
      Else
         MyNumber = ""
      End If
      count = count + 1
   Loop

   'Clean up dollars.
   Select Case Dollars
      Case ""
         Dollars = "No " & currencyName
      Case "One"
         Dollars = "One " & currencyName
      Case Else
         If Right(Dollars, 2) = ", " Then
                Dollars = Mid(Dollars, 1, Len(Dollars) - 2)
         End If
         Dollars = Dollars & " " & currencyName & "s"
   End Select

   'Clean up cents.
   Select Case Cents
      Case ""
         Cents = " And No " & subCurrencyName & "s"
      Case "One"
         Cents = " And One " & subCurrencyName
      Case Else
         Cents = " And " & Cents & " " & subCurrencyName & "s"
   End Select

   ConvertCurrencyToEnglish = Replace(Dollars & Cents, " ,", ",")
End Function
Private Function ConvertHundreds(ByVal MyNumber)
   Dim Result

   'Exit if there is nothing to convert.
   If CInt(MyNumber) = 0 Then Exit Function

   'Append leading zeros to number.
   MyNumber = Right("000" & MyNumber, 3)

   'Do we have a hundreds place digit to convert?
   If Left(MyNumber, 1) <> "0" Then
      Result = ConvertDigit(Left(MyNumber, 1)) & " Hundred "
   End If

   'Do we have a tens place digit to convert?
   If Mid(MyNumber, 2, 1) <> "0" Then
      Result = Result & ConvertTens(Mid(MyNumber, 2))
   Else
      'If not, then convert the ones place digit.
      Result = Result & ConvertDigit(Mid(MyNumber, 3))
   End If

   ConvertHundreds = Trim(Result)
End Function
Private Function ConvertTens(ByVal MyTens)
   Dim Result

   'Is value between 10 and 19?
   If CInt(Left(MyTens, 1)) = 1 Then
      Select Case CInt(MyTens)
         Case 10: Result = "Ten"
         Case 11: Result = "Eleven"
         Case 12: Result = "Twelve"
         Case 13: Result = "Thirteen"
         Case 14: Result = "Fourteen"
         Case 15: Result = "Fifteen"
         Case 16: Result = "Sixteen"
         Case 17: Result = "Seventeen"
         Case 18: Result = "Eighteen"
         Case 19: Result = "Nineteen"
         Case Else
      End Select
   Else
      '... otherwise it's between 20 and 99.
      Select Case CInt(Left(MyTens, 1))
         Case 2: Result = "Twenty "
         Case 3: Result = "Thirty "
         Case 4: Result = "Forty "
         Case 5: Result = "Fifty "
         Case 6: Result = "Sixty "
         Case 7: Result = "Seventy "
         Case 8: Result = "Eighty "
         Case 9: Result = "Ninety "
         Case Else
      End Select
      'Convert ones place digit.
      Result = Result & ConvertDigit(Right(MyTens, 1))
   End If

   ConvertTens = Result
End Function
Private Function ConvertDigit(ByVal MyDigit)
   Select Case CInt(MyDigit)
      Case 1: ConvertDigit = "One"
      Case 2: ConvertDigit = "Two"
      Case 3: ConvertDigit = "Three"
      Case 4: ConvertDigit = "Four"
      Case 5: ConvertDigit = "Five"
      Case 6: ConvertDigit = "Six"
      Case 7: ConvertDigit = "Seven"
      Case 8: ConvertDigit = "Eight"
      Case 9: ConvertDigit = "Nine"
      Case Else: ConvertDigit = ""
   End Select
End Function
Function getDatePeriodFromDelim(strDelimPeriod)
        
    Dim arPeriod, periodStart, periodEnd

    Dim arOut(1)

    arPeriod = Split(strDelimPeriod, "||")

    If UBound(arPeriod) >= 0 Then
        periodStart = arPeriod(0)
    End If

    If UBound(arPeriod) >= 1 Then
        periodEnd = arPeriod(1)
    End If

    periodStart = makeDatePeriod(Trim(periodStart), periodEnd, "0:00:00")
    periodEnd = makeDatePeriod(Trim(periodEnd), periodStart, "23:59:59")

    arOut(0) = periodStart
    arOut(1) = periodEnd

    getDatePeriodFromDelim = arOut

End Function

Function makeDatePeriod(strDateStart, defaultDate, strTime)

    If IsDate(strDateStart) Then
        makeDatePeriod = FormatDate(strDateStart) & " " & Trim(strTime)
    Else

        If IsDate(defaultDate) Then
            makeDatePeriod = FormatDate(defaultDate) & " " & Trim(strTime)
        Else
            makeDatePeriod = FormatDate(Now()) & " " & Trim(strTime)
        End If
    End If

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
