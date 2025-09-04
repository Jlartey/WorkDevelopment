'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim nm, dur, bDt, gen, pat, sltHt1, sltHt2, dyHt, modMgr, cDt, wkDy
Dim cnt, vDt1, vDt2, rst, pCnt, num, sql2, htStr, dyNm, prtUrl, sltTyp
Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, tb, tbKy, tbNm
Dim recKy, hasPrt, vst, sp, vDt, lnkCnt, nDt, prevDys, wrd, rst0, wrdNm, wrd2, fullScrn
lnkCnt = 0
prevDys = 365

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

LoadCSS
prtUrl = "wpgVisitation.asp?PageMode=ProcessSelect"
'modMgr = HasModuleMgrAccess(jSchd, tb)
If UCase(jSchd) = "QUALITYCOMPLIANCE" Then
  prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation"
ElseIf HasPrintOutAccess(jSchd, "PatientMedicalRecord") Then
  prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalRecord&PositionForTableName=WorkingDay"
ElseIf Len(HasModuleMgrAccess(jSchd, tb)) > 0 Then
  prtUrl = "wpgSelectModuleManager.asp?PositionForTableName=" & tb
ElseIf HasPrintOutAccess(jSchd, "VisitationRCP") Then
  prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All"
  sltTyp = "YES"
End If

SetPageVariable "IFrameSrc", prtUrl
SetPageVariable "IFrameSrc2", "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalHistory&PositionForTableName=Patient"
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
htStr = htStr & "if (Helpers.len(vst)>0) {" & vbCrLf
htStr = htStr & "fullScrn=GetCheckedRadio('inpFullScreen');" & vbCrLf
htStr = htStr & "ur=GetPageVariable('IFrameSrc');" & vbCrLf
htStr = htStr & "ur=ur + '&VisitationID=' + vst + '&FullScreen=' + fullScrn;" & vbCrLf
htStr = htStr & "ele = document.getElementById('iFrm1');" & vbCrLf
htStr = htStr & "if (ele) {" & vbCrLf
htStr = htStr & "if (Helpers.ucase(fullScrn)==""YES"") {" & vbCrLf
htStr = htStr & "sHt=screen.availHeight-20;" & vbCrLf
htStr = htStr & "sWd=screen.availWidth-20;" & vbCrLf
htStr = htStr & "window.open(processurl(ur));" & vbCrLf ',""_blank"",formatwinposprt(sWd,sHt));" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "else{" & vbCrLf
htStr = htStr & "ele.src=processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "}" & vbCrLf

htStr = htStr & "ur=GetPageVariable('IFrameSrc2');" & vbCrLf
htStr = htStr & "pat=Helpers.trim(form1.VisitByName.options[form1.VisitByName.selectedIndex].label);" & vbCrLf
htStr = htStr & "ur=ur + '&PatientID=' + pat;" & vbCrLf
'htStr = htStr & "window.open(processurl(ur));" & vbCrLf  ',"""",formatwinposprt(""1000"",""750""));" & vbCrLf
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
htStr = htStr & "if (Helpers.len(vst)>0) {" & vbCrLf
htStr = htStr & "fullScrn=GetCheckedRadio('inpFullScreen');" & vbCrLf
htStr = htStr & "ur=GetPageVariable('IFrameSrc');" & vbCrLf
htStr = htStr & "ur=ur + '&VisitationID=' + vst + '&FullScreen=' + fullScrn;" & vbCrLf
htStr = htStr & "ele = document.getElementById('iFrm1');" & vbCrLf
htStr = htStr & "if (ele) {" & vbCrLf
htStr = htStr & "if (Helpers.ucase(fullScrn)==""YES"") {" & vbCrLf
htStr = htStr & "sHt=screen.availHeight-20;" & vbCrLf
htStr = htStr & "sWd=screen.availWidth-20;" & vbCrLf
htStr = htStr & "window.open(processurl(ur));" & vbCrLf ',""_blank"",formatwinposprt(sWd,sHt));" & vbCrLf
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

          
wrd = Trim(Request.QueryString("Ward"))

cnt = 0

'Last Visit
'If IsNumeric(dur) Then
  cnt = cnt + 1
  nDt = Now()
  
  sql = "select Visitation.VisitationID,Visitation.PatientID,Admission.WardID,Admission.BedID,Patient.PatientName from Visitation,Patient,Admission "
  sql = sql & " where Visitation.PatientID=Patient.PatientID and Visitation.VisitationID=Admission.VisitationID "
  sql = sql & "  and Patient.PatientID=Admission.PatientID and Admission.AdmissionStatusID='A001' and Admission.WardID='" & wrd & "' order by Patient.PatientName"
  
  sql2 = "select Visitation.VisitationID,Visitation.PatientID,Admission.WardID,Admission.BedID,Patient.PatientName from Visitation,Patient,Admission "
  sql2 = sql2 & " where Visitation.PatientID=Patient.PatientID and Visitation.VisitationID=Admission.VisitationID "
  sql2 = sql2 & "  and Patient.PatientID=Admission.PatientID and Admission.AdmissionStatusID='A001' and Admission.WardID='" & wrd & "' order by Admission.AdmissionDate"
  
  pCnt = 0
  
'  'Later : Can also query distinct workingdayid on visitation to get only relevant days for affected visits.
'  dyHt = "<select size=""1"" name=""Ward"" id=""Ward"" onchange=""WardOnchange()"">"
'  dyHt = dyHt & "<option value=""""></option>"
'  With rst0
'    .Open qryPro.FltQry("select distinct WardID from Admission where AdmissionStatusID='A001'"), conn, 3, 4
'    If .RecordCount > 0 Then
'      .movefirst
'      Do While Not .EOF
'        wrd2 = Trim(.fields("WardID"))
'        wrdNm = GetComboName("Ward", wrd2)
'        If UCase(wrd) = UCase(wrd2) Then
'          dyHt = dyHt & "<option value=""" & wrd2 & """ selected>" & wrdNm & "</option>"
'        Else
'           dyHt = dyHt & "<option value=""" & wrd2 & """>" & wrdNm & "</option>"
'        End If
'        .movenext
'      Loop
'    End If
'    .Close
'  End With
'  dyHt = dyHt & "</select>"
'
'  sltHt1 = "<select size=""1"" name=""VisitByName"" id=""VisitByName"" onchange=""VisitByNameOnchange()"">"
'  sltHt1 = sltHt1 & "<option value=""""></option>"
'
'  sltHt2 = "<select size=""1"" name=""VisitByDate"" id=""VisitByDate"" onchange=""VisitByDateOnchange()"">"
'  sltHt2 = sltHt2 & "<option value=""""></option>"
'
'  With rst
'    .Open qryPro.FltQry(sql), conn, 3, 4
'    If .RecordCount > 0 Then
'      .movefirst
'
'      Do While Not .EOF
'        pCnt = pCnt + 1
'        vDt = ""
'        vst = .fields("VisitationID")
'        pat = .fields("PatientID")
'        nm = .fields("PatientName")
'        sp = .fields("BedID")
'        sltHt1 = sltHt1 & "<option value=""" & vst & """>" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & "</option>"
'        .movenext
'      Loop
'    End If
'    .Close
'  End With
'  sltHt1 = sltHt1 & "</select>"
'
'  pCnt = 0
'  With rst
'    .Open qryPro.FltQry(sql2), conn, 3, 4
'    If .RecordCount > 0 Then
'      .movefirst
'      Do While Not .EOF
'        pCnt = pCnt + 1
'        vDt = ""
'        vst = .fields("VisitationID")
'        nm = .fields("PatientName")
'        pat = .fields("PatientID")
'        nm = .fields("PatientName")
'        sp = .fields("BedID")
'        sltHt2 = sltHt2 & "<option value=""" & vst & """>" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & "</option>"
'        .movenext
'      Loop
'    End If
'    .Close
'  End With
'  sltHt2 = sltHt2 & "</select>"
  
''Type 1
'  response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
'   response.write "<tr><td colspan=""8"" align=""left"">"
'    response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
'    response.write "<tr><td class=""cpHdrTd2"">&nbsp;&nbsp;<u>Patients&nbsp;On&nbsp;:&nbsp;&nbsp;" & Replace(GetComboName("Ward", wrd), " ", "&nbsp;") & "</u>&nbsp;&nbsp;</td>"
'
''    lnkCnt = lnkCnt + 1
''    lnkID = "trslt||lnk" & CStr(lnkCnt)
''    response.write "<td onclick=""RefreshPage()"" style=""color:#448844"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
''    response.write "<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Refresh Page&nbsp;&nbsp;</b></td>"
'
'    lnkCnt = lnkCnt + 1
'    lnkID = "trslt||lnk" & CStr(lnkCnt)
'    response.write "<td onclick=""RefreshPage()"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
'    response.write "<b>Refresh&nbsp;</b></td>"
'
'    lnkCnt = lnkCnt + 1
'    lnkID = "trslt||lnk" & CStr(lnkCnt)
'    response.write "<td onclick=""cmdPrtBackOnClick()"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
'    response.write "<b>&nbsp;<<&nbsp;</b></td>"
'
'    lnkCnt = lnkCnt + 1
'    lnkID = "trslt||lnk" & CStr(lnkCnt)
'    response.write "<td onclick=""cmdPrintOnClick2()"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
'    response.write "<b>&nbsp;Print&nbsp;</b></td>"
'
'    lnkCnt = lnkCnt + 1
'    lnkID = "trslt||lnk" & CStr(lnkCnt)
'    response.write "<td onclick=""cmdPrtForwardOnClick()"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
'    response.write "<b>&nbsp;>>&nbsp;</b></td>"
'
'    fullScrn = Trim(Request("FullScreen"))
'    response.write "<td>&nbsp;&nbsp&nbsp;</td>"
'    response.write "<td><b>&nbsp;&nbsp&nbsp;Full&nbsp;Screen&nbsp;:</b></td>"
'    If UCase(fullScrn) = "NO" Then
'      response.write "<td><input type=""radio"" id=""inpFullScreen""  name=""inpFullScreen""  value=""Yes"" onclick=""inpFullScreenOnClick()"">Yes</td>"
'      response.write "<td><input type=""radio"" id=""inpFullScreen""  name=""inpFullScreen""  value=""No"" onclick=""inpFullScreenOnClick()"" checked>No</td>"
'    Else
'      response.write "<td><input type=""radio"" id=""inpFullScreen""  name=""inpFullScreen""  value=""Yes"" onclick=""inpFullScreenOnClick()"" checked>Yes</td>"
'      response.write "<td><input type=""radio"" id=""inpFullScreen""  name=""inpFullScreen""  value=""No"" onclick=""inpFullScreenOnClick()"">No</td>"
'    End If
'    'response.write "<td>&nbsp;&nbsp&nbsp;</td>"
'    'response.write "<td><b>&nbsp;&nbsp&nbsp;&nbsp;&nbsp&nbsp;Show&nbsp;on&nbsp;Side &nbsp;:</b></td>"
'    'response.write "<td><input type=""radio"" id=""inpShowOnSide""  name=""inpShowOnSide""  value=""Yes"" onclick=""inpShowOnSideOnClick()"">Yes</td>"
'    'response.write "<td><input type=""radio"" id=""inpShowOnSide""  name=""inpShowOnSide""  value=""No"" onclick=""inpShowOnSideOnClick()"">No</td>"
'    response.write "</tr>"
'    response.write "</table>"
'   response.write "</td></tr>"
'
'   response.write "<tr>"
'    response.write "<td align=""right"">Ward&nbsp;:&nbsp;</td>"
'    response.write "<td>" & dyHt & "</td>"
'    response.write "<td>&nbsp;</td>"
'    response.write "<td align=""right"">By&nbsp;Name&nbsp;:&nbsp;</td>"
'    response.write "<td>" & sltHt1 & "</td>"
'    response.write "<td>&nbsp;</td>"
'    response.write "<td align=""right"">By&nbsp;Time&nbsp;:&nbsp;</td>"
'    response.write "<td>" & sltHt2 & "</td>"
'   response.write "</tr>"
'  response.write "</table>"
'  response.write "<iframe id=""iFrm1"" width=""100%"" frameborder=""1"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"">"
'  response.write "</iframe>"
'
''  'Type 2
''  response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
''   response.write "<tr>"
''   response.write "<td colspan=""3"">"
''    response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
''    response.write "<tr><td class=""cpHdrTd2"">&nbsp;&nbsp;<u>Patients&nbsp;On&nbsp;:&nbsp;&nbsp;" & Replace(GetComboName("Ward", wrd), " ", "&nbsp;") & "</u>&nbsp;&nbsp;</td>"
''    lnkCnt = lnkCnt + 1
''    lnkID = "trslt||lnk" & CStr(lnkCnt)
''    response.write "<td onclick=""RefreshPage()"" style=""color:#448844"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
''    response.write "<b>&nbsp;&nbsp;Refresh&nbsp;Page&nbsp;&nbsp;</b></td>"
''    response.write "</tr>"
''    response.write "</table>"
''    response.write "</td>"
''    'By Time
''    response.write "<td>&nbsp;</td>"
''    response.write "<td align=""right"">By&nbsp;Time&nbsp;:&nbsp;</td>"
''    response.write "<td>" & sltHt2 & "</td>"
''    response.write "<td rowspan=""2"">&nbsp;In Full Screen</td>"
''    response.write "</tr>"
''    'Day
''   response.write "<tr>"
''    response.write "<td align=""right"">Ward&nbsp;:&nbsp;</td>"
''    response.write "<td>" & dyHt & "</td>"
''    response.write "<td>&nbsp;</td>"
''    'By Name
''    response.write "<td>&nbsp;</td>"
''    response.write "<td align=""right"">By&nbsp;Name&nbsp;:&nbsp;</td>"
''    response.write "<td>" & sltHt1 & "</td>"
''
''   response.write "</tr>"
''  response.write "</table>"

'9 Nov 2017
If UCase(sltTyp) = "NO" Then 'No Full Screen
    dyHt = "<select size=""1"" name=""Ward"" id=""Ward"" onchange=""WardOnchange()"">"
    dyHt = dyHt & "<option value=""""></option>"
    
    sDt = FormatDate(DateAdd("d", -1 * (prevDys), nDt)) & " 00:00:00"
    edt = FormatDate(nDt) & " 23:59:59"
    
    With rst
      .open qryPro.FltQry("select distinct WardID from Admission " & GetWardWhCls() & " and AdmissionStatusID='A001'"), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          wrd2 = Trim(.fields("WardID"))
          wrdNm = GetComboName("Ward", wrd2)
          If UCase(wrd) = UCase(wrd2) Then
            dyHt = dyHt & "<option value=""" & wrd2 & """ selected>" & wrdNm & "</option>"
          Else
             dyHt = dyHt & "<option value=""" & wrd2 & """>" & wrdNm & "</option>"
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
     With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      
      Do While Not .EOF
        pCnt = pCnt + 1
        vDt = ""
        vst = .fields("VisitationID")
        pat = .fields("PatientID")
        nm = .fields("PatientName")
        sp = .fields("BedID")
        sltHt1 = sltHt1 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & "</option>"
        .MoveNext
      Loop
    End If
    .Close
  End With
  sltHt1 = sltHt1 & "</select>"
  
  pCnt = 0
  With rst
    .open qryPro.FltQry(sql2), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      Do While Not .EOF
        pCnt = pCnt + 1
        vDt = ""
        vst = .fields("VisitationID")
        nm = .fields("PatientName")
        pat = .fields("PatientID")
        nm = .fields("PatientName")
        sp = .fields("BedID")
        sltHt2 = sltHt2 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & "</option>"
        .MoveNext
      Loop
    End If
    .Close
  End With
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
    edt = FormatDate(nDt) & " 23:59:59"
    cMth = ""
    mth = ""
    With rst
      .open qryPro.FltQry("select distinct WardID from Admission " & GetWardWhCls() & " and AdmissionStatusID='A001'"), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          wrd2 = Trim(.fields("WardID"))
          wrdNm = GetComboName("Ward", wrd2)
          If UCase(wrd) = UCase(wrd2) Then
            dyHt = dyHt & "<option value=""" & wrd2 & """ selected>" & wrdNm & "</option>"
          Else
             dyHt = dyHt & "<option value=""" & wrd2 & """>" & wrdNm & "</option>"
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
    With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      
      Do While Not .EOF
        pCnt = pCnt + 1
        vDt = ""
        vst = .fields("VisitationID")
        pat = .fields("PatientID")
        nm = .fields("PatientName")
        sp = .fields("BedID")
        sltHt1 = sltHt1 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & "</option>"
        .MoveNext
      Loop
    End If
    .Close
  End With
  sltHt1 = sltHt1 & "</select>"
  
  pCnt = 0
  With rst
    .open qryPro.FltQry(sql2), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      Do While Not .EOF
        pCnt = pCnt + 1
        vDt = ""
        vst = .fields("VisitationID")
        nm = .fields("PatientName")
        pat = .fields("PatientID")
        nm = .fields("PatientName")
        sp = .fields("BedID")
        sltHt2 = sltHt2 & "<option value=""" & vst & """>" & CStr(pCnt) & "&nbsp;&nbsp;" & nm & " ->" & pat & " ->" & GetComboName("Bed", sp) & "</option>"
        .MoveNext
      Loop
    End If
    .Close
  End With
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
  Dim str
  str = ""
  str = str & "<style type='text/css' id=""styPrt"">"
  str = str & ".cpHdrTd{font-size:14pt;font-weight:bold}"
  str = str & ".cpHdrTr{background-color:#eeeeee}"
  str = str & ".cpHdrTd2{font-size:12pt;font-weight:bold}"
  str = str & ".cpHdrTr2{background-color:#eeeeee}" 'fafafa
  str = str & "</style>"
  response.write str
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
      .movefirst
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
  ot = " where WardID=''"
  If Len(GetComboName("Ward", jb)) > 0 Then 'Ward Nurse
    ot = " where WardID='" & jb & "'"
    ot = " where BranchID='" & brnch & "'"
  ElseIf Len(GetComboName("Ward", Replace(jb, "head", "", 1, -1, 1))) > 0 Then 'Ward Nurse
    ot = " where WardID='" & jb & "'"
    ot = " where BranchID='" & brnch & "'"
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
  ElseIf (UCase(jb) = "BILLINGHEAD") Or (UCase(jb) = "CLAIMMANAGER") Or (UCase(jb) = "REGVISITWARD") Or (UCase(jb) = "REGVISITCON") Or (UCase(jb) = "QUALITYCOMPLIANCE") Then  'BILLING
    ot = " where WardID<>''"
    If Len(wrd) > 0 Then
      ot = " where WardID='" & wrd & "'"
    End If
  ElseIf (UCase(jb) = "MANAGEMENT") Or (UCase(jb) = UCase("chiefexecutiveofficer")) Or (UCase(jb) = "MANAGEMENT") Or (UCase(jb) = "CREDITCONTROL") Then 'Management
    ot = " where WardID<>''"
    If Len(wrd) > 0 Then
      ot = " where WardID='" & wrd & "'"
    End If
  ElseIf (UCase(jb) = "S93") Or (UCase(jb) = "S13") Or (UCase(jb) = "LabCashier") Then 'Labs: MnC, Main
    ot = " where BranchID='" & brnch & "' "
  ElseIf (UCase(jb) = "DPT005") Or (UCase(jb) = "DPT011") Then  'Labhead and radhead  '13TH MARCH 2023 Monica updated on 13th July,2023
    ot = " where BranchID='" & brnch & "' "
    ElseIf (UCase(jb) = "S19") Then  'Radiology '17th October 2022 Monica
    ot = " where BranchID='" & brnch & "' "
  ElseIf (UCase(jb) = "S95") Or (UCase(jb) = "S22") Then 'Pharmacies: MnC, Main
    ot = " where BranchID='" & brnch & "' "
  ElseIf (UCase(jb) = "MEDICALRECORDS") Then
    ot = " where BranchID='" & brnch & "'"
  ElseIf (UCase(jb) = "PATIENTSERVICES") Then
    ot = " where BranchID='" & brnch & "'"
    ElseIf (UCase(jb) = "PATIENTSERVICESHEAD") Then ' Patientserviceshead by monica
    ot = " where BranchID='" & brnch & "'"
 ElseIf (UCase(jb) = "M13") Then
    ot = " where BranchID='" & brnch & "'"
  Else
  
  End If
  
  GetWardWhCls = ot
End Function

Function IsClinicNurse(jb)
  Dim ot, lst, arr, ul, num
  ot = False
  lst = "S09||S26||S15||S30||S28||S07"
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
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
