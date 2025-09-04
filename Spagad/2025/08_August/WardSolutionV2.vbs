'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim nm, dur, bDt, gen, pat, sltHt1, sltHt2, dyHt, modMgr, cDt, wkDy
Dim cnt, vDt1, vDt2, rst, pCnt, num, sql2, htStr, dyNm, prtUrl, sltTyp
Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, tb, tbKy, tbNm
Dim recky, hasPrt, vst, sp, vDt, lnkCnt, nDt, prevDys, wrd, rst0, wrdNm, wrd2, fullScrn
lnkCnt = 0
prevDys = 90

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
If HasPrintOutAccess(jSchd, "PatientMedicalRecord") Then
  prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation"
ElseIf Len(HasModuleMgrAccess(jSchd, tb)) > 0 Then
  prtUrl = "wpgSelectModuleManager.asp?PositionForTableName=" & tb & "&PositionForCtxTableName=Visitation"
ElseIf HasPrintOutAccess(jSchd, "VisitationRCP") Then
  prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation"
  sltTyp = "YES"
End If

If UCase(Left(jSchd, 3)) = "M02" Or UCase(Left(jSchd, 2)) = "W0" Then
    prtUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PromptChangeBed&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation"
End If

SetPageVariable "IFrameSrc", prtUrl
SetPageVariable "IFrameSrc2", "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalHistory&PositionForTableName=Patient"

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

If UCase(sltTyp) = "NO" Then
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

htStr = htStr & "function RefreshPage(){" & vbCrLf
htStr = htStr & "window.location.reload();" & vbCrLf
htStr = htStr & "}" & vbCrLf

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

htStr = htStr & "function WardOnchange(){" & vbCrLf
htStr = htStr & "var ur,dy,fullScrn;" & vbCrLf
htStr = htStr & "dy=GetEleVal('Ward');" & vbCrLf
htStr = htStr & "fullScrn=GetCheckedRadio('inpFullScreen');" & vbCrLf
htStr = htStr & "ur='wpgPrtPrintLayoutAll.asp?PrintLayoutName=SelectVisitationWard&PositionForTableName=WorkingDay';" & vbCrLf
htStr = htStr & "ur=ur + '&WorkingDayID=DAY20160401&Ward=' + dy;" & vbCrLf
htStr = htStr & "window.location.href=processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf

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
htStr = htStr & "window.open(processurl(ur));" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "else{" & vbCrLf
htStr = htStr & "ele.src=processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "ur=GetPageVariable('IFrameSrc2');" & vbCrLf
htStr = htStr & "pat=Helpers.trim(form1.VisitByName.options[form1.VisitByName.selectedIndex].label);" & vbCrLf
htStr = htStr & "ur=ur + '&PatientID=' + pat;" & vbCrLf
htStr = htStr & "ele = document.getElementById('iFrm2');" & vbCrLf
htStr = htStr & "if (ele) {" & vbCrLf
htStr = htStr & "ele.src=processurl(ur);" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "}" & vbCrLf
htStr = htStr & "}" & vbCrLf

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
htStr = htStr & "window.open(processurl(ur));" & vbCrLf
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
If Len(GetWardIDFromJobSchedule(jSchd)) > 0 Then
  wrd = GetWardIDFromJobSchedule(jSchd) ' Force ward to match nurse's assigned ward
End If

cnt = 0

sql = "select Visitation.VisitationID,Visitation.PatientID,Admission.WardID,Admission.BedID,Patient.PatientName from Visitation,Patient,Admission "
sql = sql & " where Visitation.PatientID=Patient.PatientID and Visitation.VisitationID=Admission.VisitationID "
sql = sql & "  and Patient.PatientID=Admission.PatientID and Admission.AdmissionStatusID='A001' "
sql = sql & "  and Admission.TransProcessValID in ('AdmissionPro-T001', 'AdmissionPro-T002', 'AdmissionPro-T014') "
sql = sql & "  and Admission.WardID='" & wrd & "' order by Patient.PatientName"

sql2 = "select Visitation.VisitationID,Visitation.PatientID,Admission.WardID,Admission.BedID,Patient.PatientName from Visitation,Patient,Admission "
sql2 = sql2 & " where Visitation.PatientID=Patient.PatientID and Visitation.VisitationID=Admission.VisitationID "
sql2 = sql2 & "  and Patient.PatientID=Admission.PatientID and Admission.AdmissionStatusID='A001' "
sql2 = sql2 & "  and Admission.TransProcessValID in ('AdmissionPro-T001', 'AdmissionPro-T002', 'AdmissionPro-T014') "
sql2 = sql2 & "  and Admission.WardID='" & wrd & "' order by Admission.AdmissionDate"

pCnt = 0

If UCase(sltTyp) = "NO" Then
    dyHt = "<select size=""1"" name=""Ward"" id=""Ward"" onchange=""WardOnchange()"">"
    If Len(GetWardIDFromJobSchedule(jSchd)) > 0 Then
      dyHt = dyHt & "<option value=""" & GetWardIDFromJobSchedule(jSchd) & """ selected>" & GetComboName("Ward", GetWardIDFromJobSchedule(jSchd)) & "</option>"
    Else
      dyHt = dyHt & "<option value=""""></option>"
      sDt = FormatDate(DateAdd("d", -1 * (prevDys), nDt)) & " 00:00:00"
      eDt = FormatDate(nDt) & " 23:59:59"
      With rst
        .open "select distinct WardID from Admission " & GetWardWhCls() & " and AdmissionStatusID='A001' and Admission.TransProcessValID in ('AdmissionPro-T001', 'AdmissionPro-T002', 'AdmissionPro-T014') ", conn, 3, 4
        If .recordCount > 0 Then
          .MoveFirst
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
    End If
    dyHt = dyHt & "</select>"

    sltHt1 = "<select size=""1"" name=""VisitByName"" id=""VisitByName"" onchange=""VisitByNameOnchange()"">"
    sltHt1 = sltHt1 & "<option value=""""></option>"

    sltHt2 = "<select size=""1"" name=""VisitByDate"" id=""VisitByDate"" onchange=""VisitByDateOnchange()"">"
    sltHt2 = sltHt2 & "<option value=""""></option>"
    pCnt = 0
    With rst
      .open sql, conn, 3, 4
      If .recordCount > 0 Then
        .MoveFirst
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
      .open sql2, conn, 3, 4
      If .recordCount > 0 Then
        .MoveFirst
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

    response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
    response.write "<tr><td colspan=""8"" align=""left"">"
    response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
    response.write "<tr><td class=""cpHdrTd2"">&nbsp;&nbsp;<u>Visits On&nbsp;:&nbsp;&nbsp;" & FormatDate(vDt1) & "</u>&nbsp;&nbsp;</td>"
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
Else
    dyHt = "<select size=""5"" name=""Ward"" id=""Ward"" onchange=""WardOnchange()"">"
    If Len(GetWardIDFromJobSchedule(jSchd)) > 0 Then
      dyHt = dyHt & "<option value=""" & GetWardIDFromJobSchedule(jSchd) & """ selected>" & GetComboName("Ward", GetWardIDFromJobSchedule(jSchd)) & "</option>"
    Else
      dyHt = dyHt & "<option value=""""></option>"
      sDt = FormatDate(DateAdd("d", -1 * (prevDys), nDt)) & " 00:00:00"
      eDt = FormatDate(nDt) & " 23:59:59"
      cMth = ""
      mth = ""
      With rst
        .open "select distinct WardID from Admission " & GetWardWhCls() & " and AdmissionStatusID='A001' and Admission.TransProcessValID in ('AdmissionPro-T001', 'AdmissionPro-T002', 'AdmissionPro-T014') ", conn, 3, 4
        If .recordCount > 0 Then
          .MoveFirst
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
    End If
    dyHt = dyHt & "</select>"

    sltHt1 = "<select size=""12"" name=""VisitByName"" id=""VisitByName"" onchange=""VisitByNameOnchange()"">"
    sltHt1 = sltHt1 & "<option value=""""></option>"

    sltHt2 = "<select size=""12"" name=""VisitByDate"" id=""VisitByDate"" onchange=""VisitByDateOnchange()"">"
    sltHt2 = sltHt2 & "<option value=""""></option>"
    pCnt = 0
    With rst
      .open sql, conn, 3, 4
      If .recordCount > 0 Then
        .MoveFirst
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
      .open sql2, conn, 3, 4
      If .recordCount > 0 Then
        .MoveFirst
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

Sub AddUrlLink(lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth)
  Dim plusMinus, imgName, lnkOpClNavPop, align
  plusMinus = ""
  imgName = ""
  align = ""
  lnkOpClNavPop = inout & "||" & navPop & "||800||600||CLOSE"
  AddPrtNavLink lnkID, plusMinus, imgName, lnkText, lnkUrl, lnkOpClNavPop, fntSize, fntColor, bgColor, align, wdth
End Sub

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
      Else
        dat2 = FormatDate(CDate(dat1)) & " 23:59:59"
        dat1 = FormatDate(CDate(dat1)) & " 00:00:00"
      End If
    Else
      If IsDate(dat2) Then
        dat1 = FormatDate(CDate(dat2)) & " 0:00:00"
        dat2 = FormatDate(CDate(dat2)) & " 23:59:59"
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
  str = str & ".cpHdrTr2{background-color:#eeeeee}"
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

Function HasVisits(dy)
  Dim rst, sql, ot
  ot = False
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    .maxrecords = 1
    sql = "select VisitationID from Visitation "
    sql = sql & " where WorkingDayId='" & dy & "'"
    .open sql, conn, 3, 4
    If .recordCount > 0 Then
      .MoveFirst
      ot = True
    End If
    .Close
  End With
  HasVisits = ot
  Set rst = Nothing
End Function

Function GetWardIDFromJobSchedule(jb)
  Dim ot
  ot = ""
  Select Case UCase(jb)
    Case "W001": ot = "W01"
    Case "W002": ot = "W02"
    Case "W003": ot = "W03"
    Case "W004": ot = "W04"
    Case "W005": ot = "W05"
    Case "W006": ot = "W06"
    Case "W007": ot = "W07"
    Case "W008": ot = "W08"
    Case "W009": ot = "W09"
    Case "W010": ot = "W10"
    Case "W011": ot = "W11"
    Case "W012": ot = "W12"
    Case "W013": ot = "W13"
    ' W014 to W017 have no corresponding ward IDs
  End Select
  GetWardIDFromJobSchedule = ot
End Function

Function GetWardWhCls()
  Dim jb, dpt, ot, wrd, wardID
  jb = jSchd
  dpt = depID
  wrd = Trim(Request("WardID"))
  ot = " where WardID=''"
  wardID = GetWardIDFromJobSchedule(jb)
  If Len(wardID) > 0 Then
    ot = " where WardID='" & wardID & "'"
  ElseIf UCase(Left(jb, 2)) = "W0" Then
    ot = " where BranchID='" & brnch & "'"
  ElseIf IsClinicNurse(jb) Then
    ot = " where BranchID='" & brnch & "'"
  ElseIf (UCase(jb) = "S19A") Or (UCase(Left(jb, 3)) = "M03") Or (UCase(Left(jb, 3)) = "M06") Or (UCase(jb) = "M0603") Or (UCase(jb) = "M0210") Or (UCase(jb) = "M0211") Or (UCase(jb) = "M0209") Then
    ot = " where WardID<>''"
    If Len(wrd) > 0 Then
      ot = " where WardID='" & wrd & "'"
    End If
    If UCase(jb) = "M0327" Then
        ot = " where WardID IN ('W09', 'W10' ) "
    End If
  ElseIf UCase(Left(jb, 3)) = "M02" Then
    ot = " where BranchID='" & brnch & "'"
    If UCase(jb) = "M0227" Then
        ot = " where WardID IN ('W09', 'W10' ) "
    End If
  ElseIf (UCase(jb) = "BILLINGHEAD") Or (UCase(jb) = "CLAIMMANAGER") Or (UCase(jb) = "DPT014") Or (UCase(jb) = "REGVISITWARD") Or (UCase(jb) = "REGVISITCON") Then
    ot = " where WardID<>''"
    If Len(wrd) > 0 Then
      ot = " where WardID='" & wrd & "'"
    End If
  ElseIf (UCase(jb) = "MANAGEMENT") Then
    ot = " where WardID<>''"
    If Len(wrd) > 0 Then
      ot = " where WardID='" & wrd & "'"
    End If
  ElseIf (UCase(jb) = "S93") Or (UCase(jb) = "S13") Or (UCase(jb) = "LabCashier") Or (UCase(jb) = "DPT005") Then
    ot = " where BranchID='" & brnch & "' "
  ElseIf (UCase(jb) = "S95") Or (UCase(jb) = "S22") Or (UCase(jb) = "S22A") Then
    ot = " where BranchID='" & brnch & "' "
  ElseIf (UCase(jb) = "M13") Then
    ot = " where BranchID='" & brnch & "' "
  ElseIf (UCase(jb) = "DPT011") Or (UCase(jb) = "S19") Then
    ot = " where BranchID='" & brnch & "'"
  ElseIf (UCase(jb) = "DPT010") Or (UCase(jb) = "DPT022") Then
    ot = " where BranchID='" & brnch & "'"
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
