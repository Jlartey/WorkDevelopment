Dim tdDim, tdDim2, iFUrl, cmpSrcKy, cmpTabKy, msgTxt, iFUrl2
Dim emrDat, emrReq, vst, emrTab, lnkUrl, svd
emrDat = UCase(Trim(Request("EMRDataID")))
emrReq = UCase(Trim(Request("EMRRequestID")))
emrTab = UCase(Trim(Request("EMRCompTabID")))
svd = Trim(Request("cmdSave"))
vst = Trim(Request("VisitationID"))
spTypID = GetComboNameFld("Visitation", vst, "SpecialistTypeID")
wkDayID = getWorkingDayID()

If (Len(vst) > 0) Then
    If Not IsDate(GetPageVariable("EntryOpenDate")) Then
        SetPageVariable "EntryOpenDate", FormatDateDetail(Now)
    End If

    If gBillUtils.GetBillAmt(vst, "Visitation") Then
        lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=UpaidBillPrompt&BillType=Visitation"
        response.redirect lnkUrl
        response.enD
    ElseIf BlockConsult(vst) Then
        lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=BlockPrompt&TableName=Visitation"
        response.redirect lnkUrl
        response.enD
    ElseIf RequiresApproval(vst) Then
        lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=BlockPrompt&TableName=Visitation"
        lnkUrl = lnkUrl & "&PromptType=RequiresApproval&VisitationID=" & vst
        response.redirect lnkUrl
        response.enD
    
    Else
        Select Case UCase(emrDat)
            Case "TH080" 'Medical Outcome
                ProcessMedicalOutcome vst
            Case "TH051" 'Nurses notes
                If (UCase(spTypID) = "S100") Or (UCase(spTypID) = "S088") Then
                    'display nurses notes
                ElseIf IsNurseSchedule(jschd) And (Not EMRFormCompleted(vst, "NUR006", uName)) Then 'fall risk
                    lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=NoDiagnosisPrompt"
                    lnkUrl = lnkUrl & "&PreferredEMRDataID=" & emrDat & "&PreferredEMRCompTabID=" & emrTab & "&CompTableKeyID=EMRComponentID&VisitationID=" & vst
                    lnkUrl = lnkUrl & "&TargetEMRDataID=NUR006&TargetEMRCompTabID="
                    lnkUrl = lnkUrl & "&PromptType=MissingEMR"
                    
                    response.redirect ProcessCliUrl(lnkUrl)
                    response.enD
                ElseIf IsNurseSchedule(jschd) And (Not EMRFormCompleted(vst, "NUR012", uName)) Then  'braden scale
                        lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=NoDiagnosisPrompt"
                        lnkUrl = lnkUrl & "&PreferredEMRDataID=" & emrDat & "&PreferredEMRCompTabID=" & emrTab & "&CompTableKeyID=EMRComponentID&VisitationID=" & vst
                        lnkUrl = lnkUrl & "&TargetEMRDataID=NUR012&TargetEMRCompTabID="
                        lnkUrl = lnkUrl & "&PromptType=MissingEMR"
                        
                        response.redirect ProcessCliUrl(lnkUrl)
                        response.enD
                ElseIf IsNurseSchedule(jschd) And (Not EMRFormCompleted(vst, "NUR008", uName)) Then 'pain assessment
                        lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=NoDiagnosisPrompt"
                        lnkUrl = lnkUrl & "&PreferredEMRDataID=" & emrDat & "&PreferredEMRCompTabID=" & emrTab & "&CompTableKeyID=EMRComponentID&VisitationID=" & vst
                        lnkUrl = lnkUrl & "&TargetEMRDataID=NUR008&TargetEMRCompTabID="
                        lnkUrl = lnkUrl & "&PromptType=MissingEMR"
                        
                        response.redirect ProcessCliUrl(lnkUrl)
                        response.enD
                ElseIf IsNurseSchedule(jschd) And (Not EMRFormCompleted(vst, "NUR001", uName)) Then 'sBar
                        lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=NoDiagnosisPrompt"
                        lnkUrl = lnkUrl & "&PreferredEMRDataID=" & emrDat & "&PreferredEMRCompTabID=" & emrTab & "&CompTableKeyID=EMRComponentID&VisitationID=" & vst
                        lnkUrl = lnkUrl & "&TargetEMRDataID=NUR001&TargetEMRCompTabID="
                        lnkUrl = lnkUrl & "&PromptType=MissingEMR"
                        
                        response.redirect ProcessCliUrl(lnkUrl)
                        response.enD
                End If
        End Select
    End If
End If


Function IsNurseSchedule(jsc)
    Dim ot
    
    If InStr(1, jsc, "W0", 1) > 0 Then
        ot = True
    ElseIf InStr(1, jsc, "M02", 1) > 0 Then
        ot = True
    End If
    
    IsNurseSchedule = ot
End Function
Function EMRFormCompleted(vst, emrDat, uName)
    Dim ot, sql, rst
    
    ot = False
    
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "select top 1 * from EMRRequestItems where VisitationID='" & vst & "' and EMRDataID='" & emrDat & "' and WorkingDayID='" & wkDayID & "'"
        
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        ot = True
    End If
    
    rst.Close
    Set rst = Nothing
    
    EMRFormCompleted = ot
End Function
Function getWorkingDayID()
    Dim curDay, iMth, iDay
    iDay = CStr(day(Now()))
    yr = "DAY" & CStr(year(Now()))
    iMth = Month(FormatDate(Now()))
    
    If iMth < 10 Then
        iMth = "0" & CStr(iMth)
    End If
    If iDay < 10 Then
        iDay = "0" & CStr(iDay)
    End If
    
    curDay = yr & iMth & iDay
    getWorkingDayID = curDay
End Function
Function RequiresApproval(vst)
    Dim ot
    ot = False
    
    ot = (UCase(GetComboNameFld("Visitation", vst, "TransProcessValID")) = UCase("VisitationPro-T004"))
    RequiresApproval = ot
End Function
Function BlockConsult(vst)
    Dim ot, kPfx, tmp
    ot = False
    
    kPfx = GetComboNameFld("SpecialistGroup", GetComboNameFld("Visitation", vst, "SpecialistGroupID"), "KeyPrefix")
    If Len(kPfx) > 0 Then
        kPfx = Split(kPfx, "||")
        If UBound(kPfx) >= 0 Then
            '0 - block consult
            If UCase(Trim(kPfx(0))) = "YES" Then
                ot = True
            End If
        End If
    End If
    BlockConsult = ot
End Function

Sub ProcessMedicalOutcome(vst)
  Dim lnkUrl, rDir
  rDir = True
  If HasAdmission(vst) Then 'In-Patient
    If Not (DiagnosisExist2(vst) And DischargeSummExist(vst)) Then
      lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=NoDiagnosisPrompt&PromptType=NoIPDDischargePrompt"
      lnkUrl = lnkUrl & "&EMRDataID=" & emrDat & "&EMRCompTabID=" & emrTab & "&CompTableKeyID=EMRComponentID&VisitationID=" & vst
      response.redirect ProcessCliUrl(lnkUrl)
    Else
      cmpSrcKy = "TS003"
      cmpSrcKy = "FOH012"
      cmpTabKy = ""
      tdDim = " width=""700"" height=""800"""
      iFUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=EMRSelector"
      iFUrl = iFUrl & "&EMRDataID=" & cmpSrcKy & "&EMRCompTabID=" & cmpTabKy & "&CompTableKeyID=EMRComponentID&VisitationID=" & vst
      response.write "<tr>"
      response.write "<td align=""left"">"
      response.write "<iframe name=""iFrm1"" " & tdDim & " frameborder=""0"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"" src=""" & ProcessCliUrl(iFUrl) & """></iframe>"
      response.write "</td></tr>"
      response.write "</table>"
    End If
  Else 'OPD
    If Not DiagnosisExist2(vst) Then
      lnkUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=NoDiagnosisPrompt&PromptType=NoOPDDischargePrompt"
      lnkUrl = lnkUrl & "&EMRDataID=" & emrDat & "&EMRCompTabID=" & emrTab & "&CompTableKeyID=EMRComponentID&VisitationID=" & vst
      response.redirect ProcessCliUrl(lnkUrl)
    Else
      cmpSrcKy = "TH060"
      cmpTabKy = "TH06005"
      tdDim = " width=""600"" height=""300"""
      iFUrl = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=EMRSelector"
      iFUrl = iFUrl & "&EMRDataID=" & cmpSrcKy & "&EMRCompTabID=" & cmpTabKy & "&CompTableKeyID=EMRComponentID&VisitationID=" & vst

      response.write "<table cellpadding=""5"" border=""0"" cellspacing=""0"" width=""100%"">"
      response.write "<tr>"
      response.write "<td align=""left"">"
      response.write "<iframe name=""iFrm1"" " & tdDim & " frameborder=""0"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"" src=""" & ProcessCliUrl(iFUrl) & """></iframe>"
      response.write "</td></tr>"
      response.write "</table>"
    End If
  End If
End Sub
Function HasAdmission(vst)
  Dim rs, ot, sql
  Set rs = CreateObject("ADODB.Recordset")
  ot = False
  sql = "select PatientID from Admission where Visitationid='" & vst & "' and AdmissionStatusID<>'A003'"
  With rs
    .maxrecords = 1
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      ot = True
    End If
    .Close
  End With
  Set rs = Nothing
  HasAdmission = ot
End Function
Function DiagnosisExist2(vst)
  Dim rs, ot, sql
  Set rs = CreateObject("ADODB.Recordset")
  ot = False
  sql = "select DiseaseID from Diagnosis where Visitationid='" & vst & "'"
  With rs
    .maxrecords = 1
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      ot = True
    End If
    .Close
  End With
  Set rs = Nothing
  DiagnosisExist2 = ot
End Function

Function DischargeSummExist(vst)
  Dim rs, ot, sql
  Set rs = CreateObject("ADODB.Recordset")
  ot = False
  sql = "select EMRDataID from EMRRequestItems where Visitationid='" & vst & "' and EMRDataID in ('TS003', 'FOH012')"
  With rs
    .maxrecords = 1
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      ot = True
    End If
    .Close
  End With
  Set rs = Nothing
  DischargeSummExist = ot
End Function
