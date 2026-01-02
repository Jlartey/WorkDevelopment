'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim arPeriod, periodStart, periodEnd
arPeriod = getDatePeriodFromDelim(Trim(Request.QueryString("PrintFilter0")))
periodStart = arPeriod(0)
periodEnd = arPeriod(1)

Call ShowReport

Function AddStyles()
    Dim html

    html = html & "<style>"
    html = html & " .cover-letter *{min-height:40px;font-size:18px;text-align:justify;}"
    html = html & " [name=""rptgen-report""], [name=""rptgen-report""] *{text-transform:none!important;}"
    html = html & " [name=""rptgen-report""]:first-child {page-break-after:always;}"
    html = html & " [name=""rptgen-report""] {width:100%;}"
    html = html & " [name=""rptgen-report""] [name=""report-title""] th{text-align:center;}"
    html = html & " @media print{"
    html = html & " .cover-letter *{font-size:19pt!important;}"
    html = html & "     table[name=""rptgen-report""] * {"
    html = html & "         font-size: 12pt!important;"
    html = html & "         padding: 2px !important; "
    html = html & "     } "
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
Sub DisplayCoverLetterL(mth)

    Dim wkDy, prtFlt, totAmt, totClaim
    Dim cellSty, cellStyB, coPayAmt, clmAmt
    Dim dt, bMthNam

    dt = CDate(GetComboName("WorkingMonth", mth))
    bMthNam = MonthName(Month(dt)) & " " & Year(dt)
    DisplayHeader
    response.write "<div style=""page-break-after:always;line-height:40px;margin-top:35px;margin-left:50px;margin-right:50px;"" class=""cover-letter"">"
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
    response.write "        <div style="""">The amount involved is <span style='/*font-style:italic;*/' id='bill-amt-words'></span> (GHC <span style='font-style:italic;' id='bill-amt'></span>)</div>"
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
Function ShowCoPayDetail(rptGen, title, spn, insSch, billMth, exWhCls)
    Dim sql, bGrpSql, rst, tmp, fmtNumFlds, tmp2, args, colTot, amt, amtWrds

    amt = 0
    If HasCopay(spn) Then
        sql = " with GeneratedBills as ( "
        sql = sql & GetAllBillSQL()
        sql = sql & " ) "
        sql = sql & " select format(Visitation.VisitDate, 'dd/MM/yyyy') as [Date], Patient.PatientName as [Name of Patient] "
        sql = sql & "   , InsuredPatient.InsuranceNo as [Insurance/Staff No], Visitation.VisitationID as [Visit No]"
        sql = sql & "   , sum(GeneratedBills.FinalCost) as [Initial Amt], sum(GeneratedBills.FinalCost)-max(Visitation.VisitValue1) as [Final Amt]"
        sql = sql & "   , max(Visitation.VisitValue1) as [Co-Pay Amt]"
        sql = sql & " from Visitation "
        sql = sql & " inner join GeneratedBills on GeneratedBills.VisitationID=Visitation.VisitationID "
        sql = sql & " left join InsuredPatient on InsuredPatient.InsuredPatientID=Visitation.InsuredPatientID "
        sql = sql & " left join InsuranceScheme on InsuranceScheme.InsuranceSchemeID=InsuredPatient.InsuranceSchemeID "
        sql = sql & " left join Sponsor on Sponsor.SponsorID=Visitation.SponsorID "
        sql = sql & " left join InsuredPatient as PrincipalAccount on PrincipalAccount.InsuredPatientID=InsuredPatient.InitialDependantID "
        sql = sql & " left join Patient as Principal on Principal.PatientID=PrincipalAccount.PatientID "
        sql = sql & " left join Patient on Patient.PatientID=Visitation.PatientID "
        sql = sql & " where 1=1 and Visitation.VisitValue1>0"
        If exWhCls <> "" Then
          sql = sql & " and " & exWhCls
        End If
        sql = sql & " group by Visitation.VisitDate, Visitation.VisitationID, Patient.PatientName, InsuredPatient.InsuranceNo"
        sql = sql & " order by [Date], Patient.PatientName"
    
        sql = Replace(sql, "@billMonthStart", " '" & periodStart & "' ")
        sql = Replace(sql, "@billMonthEnd", " '" & periodEnd & "' ")

    
        args = "title=" & title
        args = args & ";ShowRowTotal=No;PageRecordCount=999999999"
        args = args & ";FormatMoneyFields=Initial Amt|Final Amt|Co-Pay Amt"
        'args = args & ";IgnoreFromComputations=Date|NAME OF PATIENT|NAME OF PRINCIPAL|INSURANCE/STAFF NO|DEPARTMENT|VISIT NO|Initial Amt|Final Amt"
        args = args & ";InCludeInComputations=Total|Initial Amt|Co-Pay Amt|Final Amt"
        args = args & ";FieldFunctions=Visit No:GetClaimLink"
        Set colTot = rptGen.PrintSQLReport(sql, args)

        amt = colTot("Co-Pay Amt")

    End If
    ShowCoPayDetail = amt
End Function
Function HasCopay(spn)
    Dim ot, sql, rst

    ot = 0
    sql = "select count(*) as [cnt] from Visitation where VisitValue1>0 and SponsorID='" & spn & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        ot = rst.fields("cnt")
    End If
    rst.Close

    HasCopay = (ot > 0)
End Function
Function ShowBillDetails(rptGen, title, spn, insSch, billMth, exWhCls)
    Dim sql, bGrpSql, rst, tmp, fmtNumFlds, tmp2, args, colTot, amt, amtWrds
    
    bGrpSql = " with GeneratedBills as ( "
    bGrpSql = bGrpSql & GetAllBillSQL()
    bGrpSql = bGrpSql & " ) "
    bGrpSql = bGrpSql & " select distinct GeneratedBills.BillGroupCatID, GeneratedBills.BillGroupCatName "
    bGrpSql = bGrpSql & " from GeneratedBills "
    bGrpSql = bGrpSql & " inner join Visitation on Visitation.VisitationID=GeneratedBills.VisitationID and Visitation.SponsorID='" & spn & "' "
    bGrpSql = bGrpSql & " "
    bGrpSql = bGrpSql & " order by GeneratedBills.BillGroupCatID asc; "
    
    bGrpSql = Replace(bGrpSql, "@billMonthStart", " '" & periodStart & "' ")
    bGrpSql = Replace(bGrpSql, "@billMonthEnd", " '" & periodEnd & "' ")
    
    sql = " with GeneratedBills as ( "
    sql = sql & GetAllBillSQL()
    sql = sql & " ) "
    
    sql = sql & " , Diag as ("
    sql = sql & "       select Visitation.PatientID + '||' + Visitation.WorkingDayID as PKey, string_agg(Disease.DiseaseName, ', ') as Diagnosis "
    sql = sql & "       , visitation.visitationid, CAST(visitation.patientage AS INT) AS AGE"
    sql = sql & "       , CASE WHEN visitation.genderid='GEN01' THEN 'MALE'"
    sql = sql & "       WHEN visitation.genderid='GEN02' THEN 'FEMALE' END AS GENDER"
    sql = sql & "       from Diagnosis inner join Visitation on Visitation.VisitationID=Diagnosis.VisitationID"
    sql = sql & "       inner join Disease on Disease.DiseaseID=Diagnosis.DiseaseID"
    sql = sql & "       where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd"
    sql = sql & "       group by Visitation.PatientID, visitation.visitationid, Visitation.WorkingDayID,visitation.patientage,visitation.genderid "
    sql = sql & "   )"

    sql = sql & " select format(Visitation.VisitDate, 'dd/MM/yyyy') as [Date], Patient.PatientID as [Hospital ID], Patient.PatientName as [Name of Patient],format(Patient.BirthDate, 'dd/MM/yyyy') as [DoB], Diag.Diagnosis, Diag.GENDER, Diag.AGE "
    sql = sql & " , format(admission.AdmissionDate, 'dd/MM/yyyy') as [Adm. Date], format(admission.DischargeDate, 'dd/MM/yyyy') as [Disc. Date]"
    sql = sql & "   , isnull(Principal.PatientName, Patient.PatientName) as [Name of Principal], InsuredPatient.InsuranceNo as [Insurance/Staff No], SpecialistType.SpecialistTypeName as [Service Type]"
    sql = sql & "   , InsSchemeMode.InsSchemeModeName as [Department], Visitation.VisitationID as [Visit No]"

    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(bGrpSql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If fmtNumFlds <> "" Then
                fmtNumFlds = fmtNumFlds & "|"
            End If
            fmtNumFlds = fmtNumFlds & rst.fields("BillGroupCatName")
    
            tmp = ", sum(case when GeneratedBills.BillGroupCatID='" & rst.fields("BillGroupCatID") & "' then GeneratedBills.FinalCost else 0 end) as [" & rst.fields("BillGroupCatName") & "]"
            sql = sql & tmp
            rst.MoveNext
        Loop
        
    End If
    rst.Close
    Set rst = Nothing
    
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
    sql = sql & " left join Admission on Admission.visitationid=visitation.visitationid "
    sql = sql & " left join Diag on Diag.visitationid=visitation.visitationid "
    sql = sql & " left join SpecialistType on SpecialistType.SpecialistTypeID=visitation.SpecialistTypeID "
    sql = sql & " where 1=1"
    If exWhCls <> "" Then
        sql = sql & " and " & exWhCls
    End If
    sql = sql & " group by Visitation.VisitDate, Visitation.VisitationID, Patient.PatientName, InsuredPatient.InsuranceNo, InsuranceScheme.InsuranceSchemeName, Diag.Diagnosis"
    sql = sql & "   ,Principal.PatientName,BenefitType.BenefitTypeName, Sponsor.SponsorName, InsSchemeMode.InsSchemeModeName, Diag.GENDER, Diag.AGE"
    sql = sql & " , Admission.AdmissionDate, Admission.DischargeDate, Patient.BirthDate, Patient.PatientID,  SpecialistType.SpecialistTypeName"
    sql = sql & " order by InsuranceScheme.InsuranceSchemeName, [Date], Patient.PatientName"
    
    sql = Replace(sql, "@billMonthStart", " '" & periodStart & "' ")
    sql = Replace(sql, "@billMonthEnd", " '" & periodEnd & "' ")
    
    args = "title=" & title
    args = args & ";ShowRowTotal=Yes;PageRecordCount=999999999"
    args = args & ";FormatMoneyFields=" & fmtNumFlds
    args = args & ";IncludeInComputations=Total|" & fmtNumFlds
    'args = args & ";IgnoreFromComputations=Date|NAME OF PATIENT|NAME OF PRINCIPAL|INSURANCE/STAFF NO|DEPARTMENT|VISIT NO"
    args = args & ";FieldFunctions=Visit No:GetClaimLink"

    Set colTot = rptGen.PrintSQLReport(sql, args)

    amt = colTot("Total")

    ShowBillDetails = amt
End Function
Function GetClaimLink(RECOBJ, fieldNAme)
    Dim ot, href
    
    href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationClaim&PositionForTableName=Visitation&VisitationID=" & RECOBJ(fieldNAme)
    ot = "<a href='" & href & "' style='text-decoration:none;' target='_blank'>" & RECOBJ(fieldNAme) & "</a>"
    GetClaimLink = ot
End Function
Sub AddSchemeSummary(amt, amtWrds)
    Dim html

    html = ""
    html = html & " <script>"
    html = html & "     addSchemeSummary( '" & amt & "', '" & amtWrds & "')"
    html = html & " "
    html = html & " </script>"
    response.write html
End Sub
Function GetAllBillSQL()
    Dim sql
    
    sql = ""
    sql = sql & " select Visitation.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "         , Visitation.VisitCost as FinalCost "
    sql = sql & "     from Visitation  "
    sql = sql & "     left join SpecialistType on SpecialistType.SpecialistTypeID=Visitation.SpecialistTypeID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=Visitation.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=Visitation.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select Investigation.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "         , sum(Investigation.FinalAmt) as FinalCost "
    sql = sql & "     from Investigation "
    sql = sql & "     left join Visitation on Visitation.VisitationID=Investigation.VisitationID "
    sql = sql & "     left join LabTest on LabTest.LabTestID=Investigation.LabTestID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=Investigation.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=Investigation.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    sql = sql & "     group by Investigation.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select Investigation2.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "         , sum(Investigation2.FinalAmt) as FinalCost  "
    sql = sql & "     from Investigation2 "
    sql = sql & "     left join Visitation on Visitation.VisitationID=Investigation2.VisitationID "
    sql = sql & "     left join LabTest on LabTest.LabTestID=Investigation2.LabTestID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=Investigation2.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=Investigation2.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    sql = sql & "     group by Investigation2.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select DrugSaleItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "         , sum(DrugSaleItems.FinalAmt) as FinalCost  "
    sql = sql & "     from DrugSaleItems "
    sql = sql & "     left join Visitation on Visitation.VisitationID=DrugSaleItems.VisitationID "
    sql = sql & "     left join Drug on Drug.DrugID=DrugSaleItems.DrugID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=DrugSaleItems.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=DrugSaleItems.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    sql = sql & "     group by DrugSaleItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select DrugReturnItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "         , -sum(DrugReturnItems.FinalAmt) as FinalCost  "
    sql = sql & "     from DrugReturnItems  "
    sql = sql & "     left join Visitation on Visitation.VisitationID=DrugReturnItems.VisitationID "
    sql = sql & "     left join Drug on Drug.DrugID=DrugReturnItems.DrugID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=DrugReturnItems.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=DrugReturnItems.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    sql = sql & "     group by DrugReturnItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select DrugSaleItems2.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "         , sum(DrugSaleItems2.DispenseAmt2) as FinalCost  "
    sql = sql & "     from DrugSaleItems2 "
    sql = sql & "     left join Visitation on Visitation.VisitationID=DrugSaleItems2.VisitationID "
    sql = sql & "     left join Drug on Drug.DrugID=DrugSaleItems2.DrugID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=Drug.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=DrugSaleItems2.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    sql = sql & "     group by DrugSaleItems2.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select DrugReturnItems2.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "         , -sum(DrugReturnItems2.MainItemValue1) as FinalCost  "
    sql = sql & "     from DrugReturnItems2  "
    sql = sql & "     left join Visitation on Visitation.VisitationID=DrugReturnItems2.VisitationID "
    sql = sql & "     left join Drug on Drug.DrugID=DrugReturnItems2.DrugID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=DrugReturnItems2.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=DrugReturnItems2.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    sql = sql & "     group by DrugReturnItems2.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select TreatCharges.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "         , sum(TreatCharges.FinalAmt) as FinalCost  "
    sql = sql & "     from TreatCharges  "
    sql = sql & "     left join Visitation on Visitation.VisitationID=TreatCharges.VisitationID "
    sql = sql & "     left join Treatment on Treatment.TreatmentID=TreatCharges.TreatmentID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=TreatCharges.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=TreatCharges.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    sql = sql & "     group by TreatCharges.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     union all "
    sql = sql & "  "
    sql = sql & "  "
    sql = sql & "     select StockIssueItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "           , sum(StockIssueItems.FinalAmt) as FinalCost  "
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
    sql = sql & "     select StockReturnItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "         , -sum(StockReturnItems.FinalAmt) as FinalCost  "
    sql = sql & "     from StockReturnItems  "
    sql = sql & "     left join Visitation on Visitation.VisitationID=StockReturnItems.VisitationID "
    sql = sql & "     left join Items on Items.ItemID=StockReturnItems.ItemID "
    sql = sql & "     left join BillGroup on BillGroup.BillGroupID=StockReturnItems.BillGroupID "
    sql = sql & "     left join BillGroupCat on BillGroupCat.BillGroupCatID=StockReturnItems.BillGroupCatID "
    sql = sql & "     where Visitation.BillProcessDate between @billMonthStart and @billMonthEnd "
    sql = sql & "     group by StockReturnItems.VisitationID, BillGroup.BillGroupID, BillGroup.BillGroupName"
    sql = sql & "           , Items.ItemName, BillGroupCat.BillGroupCatID, BillGroupCat.BillGroupCatName "
    sql = sql & "  "
    GetAllBillSQL = sql
End Function
Sub DisplayHeader()
    response.write "<style>div.dheader{height:300px;background:url('images/letterhead5.jpg');background-repeat:no-repeat;background-size:contain;} "
    response.write "@media print{div.dheader{content:url('images/letterhead5.jpg')}}</style>"
    
    response.write "<div class=""dheader"">"
    'response.write "<div style=""position:relative;display:flex;align-items:center;justify-content:space-between;height:auto;""><img src=""images/letterhead5.jpg"" style=""height: 150px;width:70%!important;max-height: 100%;""> </div>"
    response.write "</div>"
    
    'response.write "<table border=""0""cellspacing=""0"" cellpadding=""0"" width=""" & PrintWidth & """>"
    'response.write "<tr><td>"
    'response.write "<td style=""display:flex; align-items:center; justify-content:space-between;""><img src=""images/letterhead5b.jpg""> <img src=""images/letterhead5c.jpg"" width=""150""></td>"
    'response.write "<img src=""images/logo.jpg"" height=""60"" width=""60"">"
    'response.write "</td>"
    'response.write "<td>"
    'response.write "<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">"
    'response.write "</table>"
    'response.write "</td>"
    'response.write "</tr>"
    'response.write "</table>"
End Sub
Sub ShowReport()
    Dim mth
    Dim title, spn, insSch, billMth, exWhCls, rptGen, amt, amtWrds

    server.scripttimeout = 1800

    amt = 0
    'mth = Request.querystring("PrintFilter")
    mth = FormatWorkingMonth(periodStart)

    AddPageJS

    Set rptGen = New PRTGLO_RptGen

    DisplayCoverLetterL mth

    spn = GetRecordField("SponsorID")
    insSch = ""
    billMth = mth
    exWhCls = " InsuredPatient.InsuredPatientID<>'CANCEL' and Visitation.SponsorID='" & spn & "' "
'    title = GetRecordField("SponsorName") & " MEDICAL BILLS FOR " & GetComboName("WorkingMonth", mth)
    title = UCase(GetRecordField("SponsorName") & " MEDICAL BILLS FOR " & GetMonthName(mth))

    amt = ShowBillDetails(rptGen, title, spn, insSch, billMth, exWhCls)

    spn = GetRecordField("SponsorID")
    insSch = ""
    billMth = mth
    exWhCls = " InsuredPatient.InsuredPatientID<>'CANCEL' and Visitation.SponsorID='" & spn & "' "
    title = UCase("Co-Payment for Medical Bill for [" & GetRecordField("SponsorName") & "]")
    amt = amt - ShowCoPayDetail(rptGen, title, spn, insSch, billMth, exWhCls)

    'amtWrds = ConvertCurrencyToEnglish(amt, "Ghanaian Cedi", "Pesewa")
    amtWrds = GetPaymentWord(amt)
    AddSchemeSummary FormatNumber(amt, 2), amtWrds

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

'newly added functions for converting digits to words
Function GetPaymentWord(inAmt)
Dim amt, fAmt, wAmt, ot
ot = ""
amt = Abs(CDbl(inAmt))
wAmt = Int(amt)
fAmt = Round(amt - wAmt, 2)
ot = ot & GetAmountWord(wAmt) & " GHANA CEDI(S)"

If fAmt > 0 Then
ot = ot & " " & GetAmountWord(100 * fAmt) & " PESEWA(S)"
End If
GetPaymentWord = ot
End Function
Function GetAmountWord(inAmt)
 Dim amt, ot, amtRem, amtUnit
 amt = inAmt
 ot = ""
 If amt >= 1000000000 Then
    amtUnit = "Billion"
    ot = ot & " " & GetLess1000(Int(amt / 1000000000))
    ot = ot & " " & amtUnit
    amtRem = amt - (Int(amt / 1000000000) * 1000000000)
  ElseIf amt >= 1000000 Then
    amtUnit = "Million"
    ot = ot & " " & GetLess1000(Int(amt / 1000000))
    ot = ot & " " & amtUnit
    amtRem = amt Mod 1000000
  ElseIf amt >= 1000 Then
    amtUnit = "Thousand"
    ot = ot & " " & GetLess1000(Int(amt / 1000))
    ot = ot & " " & amtUnit
    amtRem = amt Mod 1000
  Else
    ot = ot & " " & GetLess1000(Int(amt / 1))
    amtRem = 0
  End If
  If amtRem > 0 Then
    ot = ot & " " & GetAmountWord(amtRem)
  End If
  GetAmountWord = ot
End Function

Function GetLess1000(Less1000)
  Dim ot, Less1000Rem
  ot = ""
  If Less1000 >= 100 Then
    ot = ot & " " & GetDigit(CStr(Int(Less1000 / 100)))
    ot = ot & " Hundred"
    Less1000Rem = Less1000 Mod 100
    If Less1000Rem > 0 Then
      ot = ot & " and"
    End If
  ElseIf Less1000 >= 10 Then
    If Less1000 >= 10 And Less1000 <= 19 Then
      Select Case Less1000
        Case 10
         ot = ot & "Ten"
        Case 11
         ot = ot & "Eleven"
        Case 12
         ot = ot & "Twelve"
        Case 13
         ot = ot & "Thirteen"
        Case 14
         ot = ot & "Fourteen"
        Case 15
          ot = ot & "Fifeteen"
        Case 16
          ot = ot & "Sixteen"
        Case 17
          ot = ot & "Seventeen"
        Case 18
          ot = ot & "Eighteen"
        Case 19
          ot = ot & "Nineteen"
        Case Else
        
      End Select
      Less1000Rem = 0
    Else
      ot = ot & " " & GetTens(Int(Less1000 / 10))
      Less1000Rem = Less1000 Mod 10
    End If
  ElseIf Less1000 < 10 Then
    ot = ot & " " & GetDigit(CStr(Less1000))
    Less1000Rem = 0
  End If
  
  If Less1000Rem > 0 Then
    ot = ot & " " & GetLess1000(Less1000Rem)
  End If
  GetLess1000 = ot
End Function

Function GetTens(tens)
 Dim ot
ot = ""
  Select Case tens
    Case 1
    
    Case 2
      ot = ot & "Twenty"
    Case 3
      ot = ot & "Thirty"
    Case 4
      ot = ot & "Forty"
    Case 5
      ot = ot & "Fifty"
    Case 6
      ot = ot & "Sixty"
    Case 7
      ot = ot & "Seventy"
    Case 8
      ot = ot & "Eighty"
    Case 9
      ot = ot & "Ninety"
    Case Else
  End Select
GetTens = ot
End Function

Function GetDigit(digit)
  Dim ot
  ot = ""
  Select Case digit
    Case "0"
     ot = "Zero"
    Case "1"
     ot = "One"
    Case "2"
     ot = "Two"
    Case "3"
      ot = "Three"
    Case "4"
      ot = "Four"
    Case "5"
      ot = "Five"
    Case "6"
      ot = "Six"
    Case "7"
      ot = "Seven"
    Case "8"
      ot = "Eight"
    Case "9"
      ot = "Nine"
    Case "10"
      ot = "Ten"
    Case "11"
      ot = "Eleven"
    Case "12"
      ot = "Twelve"
    Case Else
  End Select
GetDigit = ot
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
