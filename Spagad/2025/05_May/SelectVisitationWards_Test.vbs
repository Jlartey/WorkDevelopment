'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim nm, dur, bDt, gen, pat, sltHt1, sltHt2, dyHt, modMgr, cDt, wkDy
Dim cnt, vDt1, vDt2, rst, pCnt, num, sql2, htStr, dyNm, prtUrl, sltTyp, wdCnt
Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, tb, tbKy, tbNm
Dim recKy, hasPrt, vst, sp, vDt, lnkCnt, nDt, prevDys, wrd, rst0, wrdNm, wrd2, fullScrn, inSql
Dim isWrdNur, addWaitLst, waitBd, dChrgInBd '2 May 2019,6 Jun 2019
Dim addDChrgInLst, insNm
lnkCnt = 0
prevDys = 7

fullScrn = Trim(Request("FullScreen"))
sltTyp = fullScrn
If Len(Trim(sltTyp)) = 0 Then
  sltTyp = "YES"
End If
Set rst = CreateObject("ADODB.Recordset")
Set rst0 = CreateObject("ADODB.Recordset")
tb = "Visitation"
tbKy = "VisitationID"
tbNm = "Ward Patients"

'2 May 2019
wrd = Trim(Request.queryString("Ward"))
isWrdNur = Trim(GetComboName("Ward", jSchd)) '2 May 2019
If Len(isWrdNur) > 0 Then 'Is Ward Nurse, Limit to ward
  wrd = jSchd
End If
If UCase(jSchd) = UCase("W09KYB") Then ''@bless - 27 Jan 2021 >> Kybele update
  wrd = Left(UCase(jSchd), 3)
End If
If UCase(jSchd) = UCase("W001IC") Then ''@Peter - 24 Aug 2022 >> Akosombo Male Medical Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("W002IC") Then ''@Peter - 24 Aug 2022 >> Akosombo Male Surgical Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("W003IC") Then ''@Peter - 24 Aug 2022 >> Akosombo Female Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("W004IC") Then ''@Peter - 24 Aug 2022 >> Obs & Gynae Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("W005IC") Then ''@Peter - 24 Aug 2022 >> Accra Main Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If

If UCase(jSchd) = UCase("W006IC") Then ''@Peter - 24 Aug 2022 >> Akosombo Male Medical Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("W007IC") Then ''@Peter - 24 Aug 2022 >> Akosombo Male Surgical Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("W008IC") Then ''@Peter - 24 Aug 2022 >> Akosombo Female Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("W009IC") Then ''@Peter - 24 Aug 2022 >> Obs & Gynae Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("W010IC") Then ''@Peter - 24 Aug 2022 >> Accra Main Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If

If UCase(jSchd) = UCase("W011IC") Then ''@Peter - 24 Aug 2022 >> Akosombo Male Medical Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("W012IC") Then ''@Peter - 24 Aug 2022 >> Akosombo Male Surgical Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("W013IC") Then ''@Peter - 24 Aug 2022 >> Akosombo Female Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("W014IC") Then ''@Peter - 24 Aug 2022 >> Obs & Gynae Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("W015IC") Then ''@Peter - 24 Aug 2022 >> Accra Main Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If

If UCase(jSchd) = UCase("W016IC") Then ''@Peter - 24 Aug 2022 >> Accra Main Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If

If UCase(jSchd) = UCase("W017IC") Then ''@Peter - 24 Aug 2022 >> Accra Main Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If
If UCase(jSchd) = UCase("M06") Then ''@Peter - 24 Aug 2022 >> Akosombo Male Surgical Ward IC update
  wrd = Left(UCase(jSchd), 4)
End If





















waitBd = GetWaitingBed(wrd)
dChrgInBd = GetDischargeInBed(wrd)

LoadCSS
prtUrl = "wpgVisitation.asp?PageMode=ProcessSelect"
'modMgr = HasModuleMgrAccess(jSchd, tb)
If HasPrintOutAccess(jSchd, "PatientMedicalRecord") Then
  prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation"
ElseIf Len(HasModuleMgrAccess(jSchd, tb)) > 0 Then
  prtUrl = "wpgSelectModuleManager.asp?PositionForTableName=" & tb & "&PositionForCtxTableName=Visitation"
ElseIf HasPrintOutAccess(jSchd, "VisitationRCP") Then
  prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation"
  sltTyp = "YES"
End If
prtUrlKyb = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=KybPatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation"

If UCase(jSchd) = UCase("W09KYB") Then ''@bless - Kybele update
  prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=KybPatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation"
End If
' If (UCase(jSchd))=(UCase("W09")) Then ''@bless - Kybele update
'   prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=KybPatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation"
'   ' If FormatDateDetail(GetComboNameFld("Visitation", vst, "VisitDate")) < FormatDateDetail("1 Feb 2021 08:00:00") Then
'   ' Else ''Use Kybele PatientMedicalRecord from Feb 2021
'   ' End If
' End If

SetPageVariable "IFrameSrc", prtUrl
SetPageVariable "IFrameSrcKyb", prtUrlKyb
SetPageVariable "IFrameSrc2", "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalHistory&PositionForTableName=Patient"
' If UCase(jSchd)=UCase("W09KYB") Then ''@bless - Kybele update
'   SetPageVariable "IFrameSrc2", "wpgPrtPrintLayoutAll.asp?PrintLayoutName=KybPatientMedicalHistory&PositionForTableName=Patient"
' End If

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

If UCase(sltTyp) = "NO" Then 'No Full Screen
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

'Ward()
htStr = htStr & "function WardOnchange(){" & vbCrLf
htStr = htStr & "var ur,dy,fullScrn;" & vbCrLf
htStr = htStr & "dy=GetEleVal('Ward');" & vbCrLf
htStr = htStr & "fullScrn=GetCheckedRadio('inpFullScreen');" & vbCrLf
htStr = htStr & "ur='wpgPrtPrintLayoutAll.asp?PrintLayoutName=SelectVisitationWard&PositionForTableName=WorkingDay';" & vbCrLf
htStr = htStr & "ur=ur + '&WorkingDayID=DAY20160401&Ward=' + dy;" & vbCrLf
htStr = htStr & "window.location.href=processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf

'VisitByName()
htStr = htStr & "function VisitByNameOnchange(){" & vbCrLf
htStr = htStr & "var ur,vst,pat,ele,cmb,fullScrn,sHt,sWd;" & vbCrLf
htStr = htStr & "vst=Helpers.trim(GetEleVal('VisitByName'));" & vbCrLf
htStr = htStr & "form1.VisitByDate.value=vst;" & vbCrLf

'2 May 2019
htStr = htStr & "if (Helpers.ucase(vst)=='WAITINGLIST') {" & vbCrLf
htStr = htStr & "ur = ""wpgPrtPrintLayoutAll.asp?PrintLayoutName=AdmitWaitingList&PositionForTableName=WorkingDay&BedID=" & waitBd & "&WardID=" & wrd & "&WorkingDayID=DAY20180101"";" & vbCrLf
htStr = htStr & "window.open(processurl(ur))" & vbCrLf
htStr = htStr & "}" & vbCrLf

htStr = htStr & "else if (Helpers.ucase(vst)=='DISCHARGEBUTINLIST') {" & vbCrLf
htStr = htStr & "ur = ""wpgPrtPrintLayoutAll.asp?PrintLayoutName=AdmitWaitingList&PositionForTableName=WorkingDay&BedID=" & dChrgInBd & "&WardID=" & wrd & "&WorkingDayID=DAY20180101"";" & vbCrLf
htStr = htStr & "window.open(processurl(ur))" & vbCrLf
htStr = htStr & "}" & vbCrLf

htStr = htStr & "else if (Helpers.len(vst)>0) {" & vbCrLf
htStr = htStr & "fullScrn=GetCheckedRadio('inpFullScreen');" & vbCrLf
htStr = htStr & "ur=GetPageVariable('IFrameSrc');" & vbCrLf
'' switch between Kybele and Existing folders based on date
htStr = htStr & "var jb; jb ='" & Left(jSchd, 3) & "'; " & vbCrLf
htStr = htStr & "if(Helpers.ucase(jb)=='W09') {" & vbCrLf
htStr = htStr & " if (Helpers.ucase(vst) >= Helpers.ucase('V1210201001')) {" & vbCrLf
htStr = htStr & "   ur=GetPageVariable('IFrameSrcKyb'); " & vbCrLf
htStr = htStr & " } ;" & vbCrLf
htStr = htStr & "} ;" & vbCrLf

htStr = htStr & "ur=ur + '&VisitationID=' + vst + '&FullScreen=' + fullScrn;" & vbCrLf
htStr = htStr & "ele = document.getElementById('iFrm1');" & vbCrLf
htStr = htStr & "if (ele) {" & vbCrLf
htStr = htStr & "if (Helpers.ucase(fullScrn)==""YES"") {" & vbCrLf
htStr = htStr & "sHt=screen.availHeight-20;" & vbCrLf
htStr = htStr & "sWd=screen.availWidth-20;" & vbCrLf
htStr = htStr & "window.open(processurl(ur))" & vbCrLf ',""_blank"",formatwinposprt(sWd,sHt));" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "else{" & vbCrLf
htStr = htStr & "ele.src=processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "}" & vbCrLf


htStr = htStr & "ur=GetPageVariable('IFrameSrc2');" & vbCrLf
'' switch between urls
htStr = htStr & "var jb; jb ='" & Left(jSchd, 3) & "'; /* */ " & vbCrLf
htStr = htStr & "if(Helpers.ucase(jb)=='W09') {" & vbCrLf
htStr = htStr & " if (Helpers.ucase(vst) >= Helpers.ucase('V1210201001')) {" & vbCrLf
htStr = htStr & "   ur=GetPageVariable('IFrameSrcKyb'); " & vbCrLf ''V1210202587
htStr = htStr & " } ;" & vbCrLf
htStr = htStr & "} ;" & vbCrLf

htStr = htStr & "pat=Helpers.trim(form1.VisitByName.options[form1.VisitByName.selectedIndex].label);" & vbCrLf
htStr = htStr & "ur=ur + '&PatientID=' + pat;" & vbCrLf
'htStr = htStr & "window.open(processurl(ur))" & vbCrLf  ',"""",formatwinposprt(""1000"",""750""));" & vbCrLf
htStr = htStr & "ele = document.getElementById('iFrm2');" & vbCrLf
htStr = htStr & "if (ele) {" & vbCrLf
htStr = htStr & "ele.src=processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "}" & vbCrLf

'VisitByDate()
htStr = htStr & "function VisitByDateOnchange(){" & vbCrLf
htStr = htStr & "var ur,vst,pat,ele,fullScrn,sHt,sWd;" & vbCrLf
htStr = htStr & "vst=Helpers.trim(GetEleVal('VisitByDate'));" & vbCrLf
htStr = htStr & "form1.VisitByName.value=vst;" & vbCrLf

'2 May 2019
htStr = htStr & "if (Helpers.ucase(vst)=='WAITINGLIST') {" & vbCrLf
htStr = htStr & "ur = ""wpgPrtPrintLayoutAll.asp?PrintLayoutName=AdmitWaitingList&PositionForTableName=WorkingDay&BedID=" & waitBd & "&WardID=" & wrd & "&WorkingDayID=DAY20180101"";" & vbCrLf
htStr = htStr & "window.open(processurl(ur))" & vbCrLf
htStr = htStr & "}" & vbCrLf

htStr = htStr & "else if (Helpers.ucase(vst)=='DISCHARGEBUTINLIST') {" & vbCrLf
htStr = htStr & "ur = ""wpgPrtPrintLayoutAll.asp?PrintLayoutName=AdmitWaitingList&PositionForTableName=WorkingDay&BedID=" & dChrgInBd & "&WardID=" & wrd & "&WorkingDayID=DAY20180101"";" & vbCrLf
htStr = htStr & "window.open(processurl(ur))" & vbCrLf
htStr = htStr & "}" & vbCrLf

htStr = htStr & "else if (Helpers.len(vst)>0) {" & vbCrLf
htStr = htStr & "fullScrn=GetCheckedRadio('inpFullScreen');" & vbCrLf
htStr = htStr & "ur=GetPageVariable('IFrameSrc');" & vbCrLf
'' switch between Kybele and Existing folders based on date
htStr = htStr & "var jb; jb ='" & Left(jSchd, 3) & "'; " & vbCrLf
htStr = htStr & "if(Helpers.ucase(jb)=='W09') {" & vbCrLf
htStr = htStr & " if (Helpers.ucase(vst) >= Helpers.ucase('V1210201001')) {" & vbCrLf
htStr = htStr & "   ur=GetPageVariable('IFrameSrcKyb'); " & vbCrLf
htStr = htStr & " } ;" & vbCrLf
htStr = htStr & "} ;" & vbCrLf

htStr = htStr & "ur=ur + '&VisitationID=' + vst + '&FullScreen=' + fullScrn;" & vbCrLf
htStr = htStr & "ele = document.getElementById('iFrm1');" & vbCrLf
htStr = htStr & "if (ele) {" & vbCrLf
htStr = htStr & "if (Helpers.ucase(fullScrn)==""YES"") {" & vbCrLf
htStr = htStr & "sHt=screen.availHeight-20;" & vbCrLf
htStr = htStr & "sWd=screen.availWidth-20;" & vbCrLf
htStr = htStr & "window.open(processurl(ur))" & vbCrLf ',""_blank"",formatwinposprt(sWd,sHt));" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "else{" & vbCrLf
htStr = htStr & "ele.src=processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "}" & vbCrLf

htStr = htStr & "ur=GetPageVariable('IFrameSrc2');" & vbCrLf
htStr = htStr & "pat=Helpers.trim(form1.VisitByDate.options[form1.VisitByDate.selectedIndex].label);" & vbCrLf
htStr = htStr & "ur=ur + '&PatientID=' + pat;" & vbCrLf
htStr = htStr & "ele = document.getElementById('iFrm2');" & vbCrLf
htStr = htStr & "if (ele) {" & vbCrLf
htStr = htStr & "ele.src=processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "}" & vbCrLf

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

htStr = htStr & "function inpFullScreenOnClick(){" & vbCrLf
htStr = htStr & "var ur,dy,fullScrn;" & vbCrLf
htStr = htStr & "dy=GetEleVal('Ward');" & vbCrLf
htStr = htStr & "fullScrn=GetCheckedRadio('inpFullScreen');" & vbCrLf
htStr = htStr & "ur='wpgPrtPrintLayoutAll.asp?PrintLayoutName=SelectVisitationWard&PositionForTableName=WorkingDay';" & vbCrLf
htStr = htStr & "ur=ur + '&WorkingDayID=DAY20160401&Ward=' + dy + '&FullScreen=' + fullScrn;" & vbCrLf
htStr = htStr & "window.location.href=processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf

htStr = htStr & "function inpShowOnSideOnClick(){" & vbCrLf
'htStr = htStr & "alert(GetEleVal('inpShowOnSide'));" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "</script>"
response.write htStr

response.write "<style>"
response.write ".cmpTdSty {"
response.write "border:1px solid #d0d0d0;"
response.write "border-collapse: collapse;"
response.write "}"
response.write "</style>"

inSql = "select v.patientid,v.patienttypeid,v.benefittypeid,b.benefittypename,v.insuranceno,v.patientrankid,v.insurancetypeid,v.PatientAge" 'GetTableSql("Visitation")
inSql = inSql & " from Visitation as v,BenefitType as b  where v.BenefitTypeID=b.BenefitTypeID "
        
cnt = 0

'Last Visit
'If IsNumeric(dur) Then
  cnt = cnt + 1
  nDt = Now()
  

  addWaitLst = False
  addDChrgInLst = False
  addDChrgInLst = GetHasDischargeInList(wrd)
  If Len(isWrdNur) > 0 Then 'Is Ward Nurse -> Exclude Waiting List
    addWaitLst = GetHasWaitingList(wrd)
    sql = "select Visitation.VisitationID,Visitation.PatientID,Admission.WardID,Admission.BedID,Patient.PatientName"
    sql = sql & ",InsuranceScheme.InsuranceSchemeName from Visitation,Patient,Admission,InsuranceScheme "
    sql = sql & " where Visitation.PatientID=Patient.PatientID and Visitation.VisitationID=Admission.VisitationID "
    sql = sql & " and Admission.BranchID<>'' and Admission.BedNoID<>'000' and Admission.BedNoID<>'999' "
    sql = sql & " and Visitation.InsuranceSchemeID=InsuranceScheme.InsuranceSchemeID "
    sql = sql & "  and Patient.PatientID=Admission.PatientID and Admission.AdmissionStatusID='A001' "
    If UCase(brnch) = UCase("B002") Then
      sql = sql & " and Admission.BranchID='" & brnch & "' order by Patient.PatientName"
    Else
      sql = sql & " and Admission.WardID='" & wrd & "' order by Patient.PatientName"
    End If
    
    sql2 = "select Visitation.VisitationID,Visitation.PatientID,Admission.WardID,Admission.BedID,Patient.PatientName"
    sql2 = sql2 & ",InsuranceScheme.InsuranceSchemeName from Visitation,Patient,Admission,InsuranceScheme "
    sql2 = sql2 & " where Visitation.PatientID=Patient.PatientID and Visitation.VisitationID=Admission.VisitationID "
    sql2 = sql2 & " and Admission.BranchID<>'' and Admission.BedNoID<>'000' and Admission.BedNoID<>'999' "
    sql2 = sql2 & " and Visitation.InsuranceSchemeID=InsuranceScheme.InsuranceSchemeID "
    sql2 = sql2 & "  and Patient.PatientID=Admission.PatientID and Admission.AdmissionStatusID='A001' "
    If UCase(brnch) = UCase("B002") Then
      sql2 = sql2 & " and Admission.BranchID='" & brnch & "' order by Admission.AdmissionDate"
    Else
      sql2 = sql2 & " and Admission.WardID='" & wrd & "' order by Admission.AdmissionDate"
    End If

  Else
    sql = "select Visitation.VisitationID,Visitation.PatientID,Admission.WardID,Admission.BedID,Patient.PatientName"
    sql = sql & ",InsuranceScheme.InsuranceSchemeName, Admission.BlockID from Visitation,Patient,Admission,InsuranceScheme "
    sql = sql & " where Visitation.PatientID=Patient.PatientID and Visitation.VisitationID=Admission.VisitationID "
    sql = sql & " and Admission.BranchID='" & brnch & "' and Admission.BedNoID<>'999' " ' and Admission.BedNoID<>'000' "
    sql = sql & " and Visitation.InsuranceSchemeID=InsuranceScheme.InsuranceSchemeID "
    sql = sql & "  and Patient.PatientID=Admission.PatientID and Admission.AdmissionStatusID='A001' "
    sql = sql & " and Admission.WardID='" & wrd & "' " '' " order by Patient.PatientName"
    ' sql = sql & " and Ward.BlockID='" & brnch & "' order by Patient.PatientName"
    sql = sql & " and Admission.BlockID='" & brnch & "' order by Patient.PatientName"
    
    sql2 = "select Visitation.VisitationID,Visitation.PatientID,Admission.WardID,Admission.BedID,Patient.PatientName"
    sql2 = sql2 & ",InsuranceScheme.InsuranceSchemeName, Admission.BlockID from Visitation,Patient,Admission,InsuranceScheme "
    sql2 = sql2 & " where Visitation.PatientID=Patient.PatientID and Visitation.VisitationID=Admission.VisitationID "
    sql2 = sql2 & " and Admission.BranchID='" & brnch & "' and Admission.BedNoID<>'999' " ' and Admission.BedNoID<>'000' "
    sql2 = sql2 & " and Visitation.InsuranceSchemeID=InsuranceScheme.InsuranceSchemeID "
    sql2 = sql2 & "  and Patient.PatientID=Admission.PatientID and Admission.AdmissionStatusID='A001' "
    sql2 = sql2 & " and Admission.WardID='" & wrd & "' " '' " order by Admission.AdmissionDate"
    ' sql2 = sql2 & " and Ward.BlockID='" & brnch & "' order by Admission.AdmissionDate"
    sql2 = sql2 & " and Admission.BlockID='" & brnch & "' order by Admission.AdmissionDate"
  End If
  pCnt = 0
  
'9 Nov 2017
jb = jSchd
If UCase(sltTyp) = "NO" Then 'No Full Screen
    dyHt = "<select size=""1"" name=""Ward"" id=""Ward"" onchange=""WardOnchange()"">"
    dyHt = dyHt & "<option value=""""></option>"
    
    sDt = FormatDate(DateAdd("d", -1 * (prevDys), nDt)) & " 00:00:00"
    eDt = FormatDate(nDt) & " 23:59:59"
    wdCnt = 0
    With rst
      .open qryPro.FltQry("select distinct WardID from Admission " & GetWardWhCls() & " and AdmissionStatusID='A001' order by WardID"), conn, 3, 4
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          wdCnt = wdCnt + 1
          wrd2 = Trim(.fields("WardID"))
          wrdNm = GetComboName("Ward", wrd2)
          If UCase(wrd) = UCase(wrd2) Then
            dyHt = dyHt & "<option value=""" & wrd2 & """ selected>" & CStr(wdCnt) & "&nbsp;&nbsp;" & wrdNm & "</option>"
          Else
             dyHt = dyHt & "<option value=""" & wrd2 & """>" & CStr(wdCnt) & "&nbsp;&nbsp;" & wrdNm & "</option>"
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
    dyHt = dyHt & "</select>"
    
    sltHt1 = "<select size=""1"" name=""VisitByName"" id=""VisitByName"" onchange=""VisitByNameOnchange()"">"
    sltHt1 = sltHt1 & "<option value=""""></option>"
    
    sltHt2 = "<select size=""1"" name=""VisitByDate"" id=""VisitByDate"" onchange=""VisitByDateOnchange()"">"
    sltHt2 = sltHt2 & "<option value=""""></option>"
    pCnt = 0
    '2 May 2019
      If addWaitLst Then
        sltHt1 = sltHt1 & "<option value=""WAITINGLIST"">WAITING LIST [ASSIGN TO BED]</option>"
      End If
     With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      
      Do While Not .EOF
        pCnt = pCnt + 1
        vDt = ""
        vst = .fields("VisitationID")
        pat = .fields("PatientID")
        nm = .fields("PatientName") 'GetPatientNameVst0(vst, inSql) '
        sp = .fields("BedID")
        insNm = .fields("InsuranceSchemeName")
        '@bless - 19 Dec 2019 >> Show Insurance for Non-Medical Users'
        If (UCase(Left(jb, 3)) = "DPT") Or (UCase(jb) = UCase("M27071")) Or (UCase(jb) = "BILLINGHEAD") Or (UCase(jb) = "CLAIMMANAGER") Or (UCase(jb) = "REGVISITWARD") Or (UCase(jb) = "REGVISITCON") Or (UCase(jb) = "M0198") Then
          sltHt1 = sltHt1 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & " ->" & insNm & "</option>"
        Else
          tm = GetTeamName(vst)
          If UCase(jSchd) = UCase(uName) Then
            setpagemessges tm
          End If
          sltHt1 = sltHt1 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & tm & "</option>"
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  If addDChrgInLst Then
    sltHt1 = sltHt1 & "<option value=""DISCHARGEBUTINLIST"">DISCHARGED-BUT-IN LIST</option>"
  End If
  sltHt1 = sltHt1 & "</select>"
  
  pCnt = 0
  '2 May 2019
  If addWaitLst Then
    sltHt2 = sltHt2 & "<option value=""WAITINGLIST"">WAITING LIST [ASSIGN TO BED]</option>"
  End If
  With rst
    .open qryPro.FltQry(sql2), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        pCnt = pCnt + 1
        vDt = ""
        vst = .fields("VisitationID")
        pat = .fields("PatientID")
        nm = .fields("PatientName")  'GetPatientNameVst0(vst, inSql) '
        sp = .fields("BedID")
        '@bless - 19 Dec 2019 >> Show Insurance for Non-Medical Users'
        insNm = .fields("InsuranceSchemeName")
        If (UCase(Left(jb, 3)) = "DPT") Or (UCase(jb) = UCase("M27071")) Or (UCase(jb) = "BILLINGHEAD") Or (UCase(jb) = "CLAIMMANAGER") Or (UCase(jb) = "REGVISITWARD") Or (UCase(jb) = "REGVISITCON") Or (UCase(jb) = "M0198") Then
          sltHt2 = sltHt2 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & " ->" & insNm & "</option>"
        Else
          tm = GetTeamName(vst)
          sltHt2 = sltHt2 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & tm & "</option>"
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  If addDChrgInLst Then
    sltHt2 = sltHt2 & "<option value=""DISCHARGEBUTINLIST"">DISCHARGED-BUT-IN LIST</option>"
  End If
  sltHt2 = sltHt2 & "</select>"
    
    'Type 1
    response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
     response.write "<tr><td colspan=""8"" align=""left"">"
      response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
      response.write "<tr><td class=""cpHdrTd2"">&nbsp;&nbsp;<u>Visits On&nbsp;:&nbsp;&nbsp;" & FormatDate(vDt1) & "</u>&nbsp;&nbsp;</td>"
  '    lnkCnt = lnkCnt + 1
  '    lnkID = "trslt||lnk" & CStr(lnkCnt)
  '    response.write "<td onclick=""RefreshPage()"" style=""color:#448844"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
  '    response.write "<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Refresh Page&nbsp;&nbsp;</b></td>"
  
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
      
      
      response.write "<td>&nbsp;&nbsp&nbsp;</td>"
      response.write "<td><b>&nbsp;&nbsp&nbsp;Full&nbsp;Screen&nbsp;:</b></td>"
      
      If UCase(sltTyp) = "YES" Then
        response.write "<td><input type=""radio"" id=""inpFullScreen""  name=""inpFullScreen""  value=""Yes"" onclick=""inpFullScreenOnClick()"" checked>Yes</td>"
        response.write "<td><input type=""radio"" id=""inpFullScreen""  name=""inpFullScreen""  value=""No"" onclick=""inpFullScreenOnClick()"">No</td>"
      Else
        response.write "<td><input type=""radio"" id=""inpFullScreen""  name=""inpFullScreen""  value=""Yes"" onclick=""inpFullScreenOnClick()"">Yes</td>"
        response.write "<td><input type=""radio"" id=""inpFullScreen""  name=""inpFullScreen""  value=""No"" onclick=""inpFullScreenOnClick()"" checked>No</td>"
      End If
      
      response.write "</tr>"
      response.write "</table>"
     response.write "</td></tr>"
  
     response.write "<tr>"
      response.write "<td align=""right"">&nbsp;&nbsp;Ward&nbsp;:&nbsp;</td>"
      response.write "<td>" & dyHt & "</td>"
      response.write "<td>&nbsp;</td>"
      response.write "<td align=""right"">By&nbsp;Name&nbsp;:&nbsp;</td>"
      response.write "<td>" & sltHt1 & "</td>"
      response.write "<td>&nbsp;</td>"
      response.write "<td align=""right"">By&nbsp;Time&nbsp;:&nbsp;</td>"
      response.write "<td>" & sltHt2 & "</td>"
     response.write "</tr>"
    response.write "</table>"
    response.write "<iframe id=""iFrm1"" width=""100%"" frameborder=""1"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"">"
    response.write "</iframe>"
  Else ' Full Screen
    dyHt = "<select size=""5"" name=""Ward"" id=""Ward"" onchange=""WardOnchange()"">"
    dyHt = dyHt & "<option value=""""></option>"
    
    sDt = FormatDate(DateAdd("d", -1 * (prevDys), nDt)) & " 00:00:00"
    eDt = FormatDate(nDt) & " 23:59:59"
    cMth = ""
    mth = ""
    With rst
      .open qryPro.FltQry("select distinct WardID from Admission " & GetWardWhCls() & " and AdmissionStatusID='A001' order by WardID"), conn, 3, 4
      If .RecordCount > 0 Then
        .MoveFirst
        wdCnt = 0
        Do While Not .EOF
          wdCnt = wdCnt + 1
          wrd2 = Trim(.fields("WardID"))
          wrdNm = GetComboName("Ward", wrd2)
          If UCase(wrd) = UCase(wrd2) Then
            dyHt = dyHt & "<option value=""" & wrd2 & """ selected>" & CStr(wdCnt) & "&nbsp;&nbsp;" & wrdNm & "</option>"
          Else
             dyHt = dyHt & "<option value=""" & wrd2 & """>" & CStr(wdCnt) & "&nbsp;&nbsp;" & wrdNm & "</option>"
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
    dyHt = dyHt & "</select>"
    
    sltHt1 = "<select size=""12"" name=""VisitByName"" id=""VisitByName"" onchange=""VisitByNameOnchange()"">"
    sltHt1 = sltHt1 & "<option value=""""></option>"
    
    sltHt2 = "<select size=""12"" name=""VisitByDate"" id=""VisitByDate"" onchange=""VisitByDateOnchange()"">"
    sltHt2 = sltHt2 & "<option value=""""></option>"
    pCnt = 0
    '2 May 2019
    If addWaitLst Then
      sltHt1 = sltHt1 & "<option value=""WAITINGLIST"">WAITING LIST [ASSIGN TO BED]</option>"
    End If
    With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      
      Do While Not .EOF
        pCnt = pCnt + 1
        vDt = ""
        vst = .fields("VisitationID")
        pat = .fields("PatientID")
        nm = .fields("PatientName") 'GetPatientNameVst0(vst, inSql) ' .fields("PatientName")
        sp = .fields("BedID")
        '@bless - 19 Dec 2019 >> Show Insurance for Non-Medical Users'
        ' sltHt1 = sltHt1 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & "</option>"
        insNm = .fields("InsuranceSchemeName")
        If (UCase(Left(jb, 3)) = "DPT") Or (UCase(jb) = UCase("M27071")) Or (UCase(jb) = "BILLINGHEAD") Or (UCase(jb) = "CLAIMMANAGER") Or (UCase(jb) = "REGVISITWARD") Or (UCase(jb) = "REGVISITCON") Or (UCase(jb) = "M0198") Then
          sltHt1 = sltHt1 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & " ->" & insNm & "</option>"
        Else
          sltHt1 = sltHt1 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & "</option>"
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  If addDChrgInLst Then
    sltHt1 = sltHt1 & "<option value=""DISCHARGEBUTINLIST"">DISCHARGED-BUT-IN LIST</option>"
  End If
  sltHt1 = sltHt1 & "</select>"
  
  pCnt = 0
  '2 May 2019
  If addWaitLst Then
    sltHt2 = sltHt2 & "<option value=""WAITINGLIST"">WAITING LIST [ASSIGN TO BED]</option>"
  End If
  With rst
    .open qryPro.FltQry(sql2), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      
      Do While Not .EOF
        pCnt = pCnt + 1
        vDt = ""
        vst = .fields("VisitationID")
        pat = .fields("PatientID")
        nm = .fields("PatientName") 'GetPatientNameVst0(vst, inSql) '.fields("PatientName")
        sp = .fields("BedID")
        ' sltHt2 = sltHt2 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & "</option>"
        '@bless - 19 Dec 2019 >> Show Insurance for Non-Medical Users'
        insNm = .fields("InsuranceSchemeName")
        If (UCase(Left(jb, 3)) = "DPT") Or (UCase(jb) = UCase("M27071")) Or (UCase(jb) = "BILLINGHEAD") Or (UCase(jb) = "CLAIMMANAGER") Or (UCase(jb) = "REGVISITWARD") Or (UCase(jb) = "REGVISITCON") Or (UCase(jb) = "M0198") Then
          sltHt2 = sltHt2 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & " ->" & insNm & "</option>"
        Else
          sltHt2 = sltHt2 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & "</option>"
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  If addDChrgInLst Then
    sltHt2 = sltHt2 & "<option value=""DISCHARGEBUTINLIST"">DISCHARGED-BUT-IN LIST</option>"
  End If
  sltHt2 = sltHt2 & "</select>"
    
    'Type 1
    response.write "<table class=""cmpTdSty"" cellpadding=""3"" cellspacing=""0"" width=""100%"">"
     response.write "<tr><td align=""left"" width=""20%"" valign=""top"">"
      response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
       response.write "<tr><td colspan=""2"" align=""center"">"
        response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
        response.write "<tr><td class=""cpHdrTd2"">&nbsp;&nbsp;<u>Visits&nbsp;On&nbsp;:&nbsp;&nbsp;" & Replace(FormatDate(vDt1), " ", "&nbsp;") & "</u>&nbsp;&nbsp;</td>"
         response.write "<td><b>&nbsp&nbsp;Full&nbsp;Screen&nbsp;:</b></td>"
         
         If UCase(sltTyp) = "YES" Then
           response.write "<td><input type=""radio"" id=""inpFullScreen""  name=""inpFullScreen""  value=""Yes"" onclick=""inpFullScreenOnClick()"" checked>Yes</td>"
           response.write "<td><input type=""radio"" id=""inpFullScreen""  name=""inpFullScreen""  value=""No"" onclick=""inpFullScreenOnClick()"">No</td>"
         Else
           response.write "<td><input type=""radio"" id=""inpFullScreen""  name=""inpFullScreen""  value=""Yes"" onclick=""inpFullScreenOnClick()"">Yes</td>"
           response.write "<td><input type=""radio"" id=""inpFullScreen""  name=""inpFullScreen""  value=""No"" onclick=""inpFullScreenOnClick()"" checked>No</td>"
         End If
         
         response.write "</tr>"
         response.write "</table></td></tr>"
        response.write "<tr><td colspan=""2"" align=""center"">"
        response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
         response.write "<tr>"
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
        response.write "</tr>"
        response.write "</table>"
       response.write "</td></tr>"
    
       response.write "<tr>"
        response.write "<td align=""right"" valign=""top"">Ward&nbsp;:&nbsp;</td>"
        response.write "<td>" & dyHt & "</td>"
        response.write "</tr>"
         
         response.write "<tr>"
        response.write "<td align=""right"" valign=""top"">By&nbsp;Name&nbsp;:&nbsp;</td>"
        response.write "<td>" & sltHt1 & "</td>"
        response.write "</tr>"
         
         response.write "<tr>"
        response.write "<td align=""right"" valign=""top"">By&nbsp;Time&nbsp;:&nbsp;</td>"
        response.write "<td>" & sltHt2 & "</td>"
       response.write "</tr>"
      response.write "</table>"
      
      response.write "</td>"
      response.write "<td valign=""top"">"
    response.write "<iframe id=""iFrm1"" width=""100%"" frameborder=""1"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"">"
    response.write "</iframe>"
    response.write "</td></tr></table>"
  End If
  Set rst = Nothing
'End If
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
Sub LoadCSS()
  Dim Str
  Str = ""
  Str = Str & "<style type='text/css' id=""styPrt"">"
  Str = Str & ".cpHdrTd{font-size:14pt;font-weight:bold}"
  Str = Str & ".cpHdrTr{background-color:#eeeeee}"
  Str = Str & ".cpHdrTd2{font-size:12pt;font-weight:bold}"
  Str = Str & ".cpHdrTr2{background-color:#eeeeee}" 'fafafa
  Str = Str & "</style>"
  response.write Str
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
Function HasVisits(dy)
  Dim rst, sql, ot
  ot = False
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    .maxrecords = 1
    sql = "select VisitationID from Visitation "
    sql = sql & " where WorkingDayId='" & dy & "'"
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      ot = True
    End If
    .Close
  End With
  HasVisits = ot
  Set rst = Nothing
End Function
Function GetWardWhCls()
  Dim jb, dpt, ot, wrd
  jb = jSchd
  dpt = depid
  wrd = Trim(Request("WardID"))
  ot = " where WardID='' "
  If Len(GetComboName("Ward", jb)) > 0 Then 'Ward Nurse
    ot = " where WardID='" & jb & "'"
  ElseIf IsClinicNurse(jb) Then  'Clinic Nurse
    ot = " where BranchID='" & brnch & "'"
  ElseIf (UCase(Left(jb, 3)) = "M03") Or (UCase(jb) = "M0603") Or (UCase(jb) = "M0210") Or (UCase(jb) = "M0211") Or (UCase(jb) = "M0209") Then 'Doctor/Pharmacist/Theatre/Anaethesia/Public Health
    ot = " where WardID<>''"
    If Len(wrd) > 0 Then
      ot = " where WardID='" & wrd & "'"
    End If
  ElseIf UCase(Left(jb, 3)) = "M02" Then 'OPD Nurse
    'ot = " where WardID=''"
    ot = " where BranchID='" & brnch & "'" 'Temp until assigned to ward profile
  ElseIf (UCase(jb) = "BILLINGHEAD") Or (UCase(jb) = "CLAIMMANAGER") Or (UCase(jb) = "REGVISITWARD") Or (UCase(jb) = "REGVISITCON") Then 'BILLING
    ot = " where WardID<>''"
    If Len(wrd) > 0 Then
      ot = " where WardID='" & wrd & "'"
    End If
  ElseIf (UCase(jb) = "MANAGEMENT") Then 'Management
    ot = " where WardID<>''"
    If Len(wrd) > 0 Then
      ot = " where WardID='" & wrd & "'"
    End If
  ElseIf (UCase(jb) = "S93") Or (UCase(jb) = "S13") Or (UCase(jb) = "LabCashier") Then 'Labs: MnC, Main
    ot = " where BranchID='" & brnch & "' "
  ElseIf (UCase(jb) = "S95") Or (UCase(jb) = "S22") Then 'Pharmacies: MnC, Main
    ot = " where BranchID='" & brnch & "' "
  ElseIf (UCase(jb) = "MEDICALRECORDS") Then
    ot = " where BranchID='" & brnch & "'"
  ElseIf (UCase(jb) = "PATIENTSERVICES") Then
    ot = " where BranchID='" & brnch & "'"
  Else
  
  End If
  
  GetWardWhCls = ot
End Function

Function IsClinicNurse(jb)
  Dim ot, lst, arr, ul, num
  ot = False
  lst = "S02||S09||S26||S15||S30||S28||S07"
  arr = Split(lst, "||")
  ul = UBound(arr)
  For num = 0 To ul
    If UCase(Trim(arr(num))) = UCase(Trim(jb)) Then
      ot = True
      Exit For
    End If
  Next
  IsClinicNurse = ot
End Function
'GetPatientNameVst0
Function GetPatientNameVst0(vst, inSql)
  Dim rs, sql, rs2, insno, rnkid, rnknm, rnkSt, Rel, relnm, patid, ptId, title, FirstName, SurName, OtherName, age, ot
  Dim insTyp
  Set rs = CreateObject("ADODB.Recordset")
  Set rs2 = CreateObject("ADODB.Recordset")
  ot = ""

  sql = inSql & " and v.Visitationid='" & vst & "'"
  With rs
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      patid = UCase(.fields("patientid"))
      ptId = UCase(.fields("patienttypeid"))
      Rel = UCase(.fields("benefittypeid"))
      relnm = UCase(.fields("benefittypename"))
      insno = .fields("insuranceno")
      rnkid = .fields("patientrankid")
      rnkSt = UCase(GetComboNameFld("patientrank", rnkid, "patientrankstatusid"))
      age = CDbl(.fields("PatientAge"))
      insTyp = .fields("insurancetypeid")
      sql = "select Titleid,surname,firstname,othername from Patient where Patientid='" & patid & "'"
      With rs2
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
          .MoveFirst
          title = GetComboName("Title", .fields("titleid"))
          SurName = .fields("surname")
          FirstName = .fields("firstname")
          OtherName = .fields("othername")
          ot = title & " " & FirstName & ", " & SurName & ", " & OtherName
          If UCase(insTyp) = "I014" And Len(insno) > 0 And insno <> "-" Then 'Entitled
            ot = "" 'insno
            If age < 60 And Rel = "X" Then
                Rel = "B006"
            End If
            rnknm = ""
            If rnkid <> "63" Then 'has rank, so set rank
              rnknm = GetComboNameFld("patientrank", rnkid, "rankabbreviate")
              If Rel = "X" Then
                If rnkSt = "P001" Then
                  rnknm = rnknm & "(Rtd)"
                ElseIf rnkSt = "P002" Then
                  rnknm = "Ex-" & rnknm
                End If
              End If
              If (Rel = "B006") Or (Rel = "X") Or (Rel = "B001") Then 'Not a dependant
                'ot = insno & " " & rnknm & " " & FirstName & ", " & SurName & ", " & OtherName
                ot = rnknm & " " & FirstName & ", " & SurName & ", " & OtherName
                rankName = rnknm
                ServiceNo = insno
              Else
                'ot = relnm & " " & insno & " " & rnknm & " " & FirstName & ", (" & SurName & " " & OtherName & ")"
                ot = relnm & " " & rnknm & " " & FirstName & ", (" & SurName & " " & OtherName & ")"
              End If
            End If
            If ptId = "CE" Or ptId = "EF" Then
              ot = "C/E " & title & " " & FirstName & ", " & SurName & ", " & OtherName
              If ptId = "EF" Then
                ot = relnm & " " & ot
              End If
            End If
          End If
        End If
        .Close
      End With
    End If
    .Close
  End With
  GetPatientNameVst0 = ot
  Set rs = Nothing
  Set rs2 = Nothing
End Function

'2 May 2019
Function GetHasWaitingList(wrd)
  Dim rst, sql, ot
  ot = False
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    sql = "select PatientID from Admission "
    sql = sql & " where Wardid='" & wrd & "' and BedNoId='000' and AdmissionStatusID='A001'"
    .maxrecords = 1
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      ot = True
    End If
    .Close
  End With
  GetHasWaitingList = ot
  Set rst = Nothing
End Function
'6 Jun 2019
Function GetHasDischargeInList(wrd)
  Dim rst, sql, ot
  ot = False
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    sql = "select PatientID from Admission "
    sql = sql & " where Wardid='" & wrd & "' and BedNoId='999' and AdmissionStatusID='A001'"
    .maxrecords = 1
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      ot = True
    End If
    .Close
  End With
  GetHasDischargeInList = ot
  Set rst = Nothing
End Function

Function GetWaitingBed(wrd)
  Dim rst, sql, ot
  ot = False
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    sql = "select BedID from Bed "
    sql = sql & " where Wardid='" & wrd & "' and BedNoId='000'"
    .maxrecords = 1
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      ot = .fields("BedID")
    End If
    .Close
  End With
  GetWaitingBed = ot
End Function
Function GetDischargeInBed(wrd)
  Dim rst, sql, ot
  ot = False
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    sql = "select BedID from Bed "
    sql = sql & " where Wardid='" & wrd & "' and BedNoId='999'"
    .maxrecords = 1
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      ot = .fields("BedID")
    End If
    .Close
  End With
  GetDischargeInBed = ot
End Function
Function GetTeamName(vst)
  Dim rst, sql, ot
  ot = ""
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    sql = "select * From FollowupVisit where VisitationID='" & vst & "' "
    sql = sql & " And FollowupVisitStatID='F001' " ''active
    sql = sql & " order by WorkingDayID "
    .maxrecords = 1
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      ot = ot & " ->" & GetComboName("FollowupVisitType", .fields("FollowupVisitTypeID"))
    End If
    .Close
  End With
  GetTeamName = ot
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
