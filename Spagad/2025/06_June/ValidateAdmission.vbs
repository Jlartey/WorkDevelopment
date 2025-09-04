'************************************************************************
'ProcessText
'************************************************************************
processcode

Sub processcode()
Dim tbl, ot, vld, pos, vLst, fLst, pos1, pat, pos2, vl2, pos3, vl3, aSt, amtval, dDt
Dim arr, num, ul, arr2, num2, ul2, vl, eps, dt, exmpt, recNo, recTyp, amtTbl, amtfld, kyFld
Dim ins, bd, dy, cKy, tblMode, bedID

tblMode = objPage.PageMode

vld = True
tbl = UCase(GetPageVariable("TableName"))
If (tbl = "ADMISSION") Then
  eps = Trim(Request("inpVisitationID"))
  pat = Trim(Request("inpPatientID"))
  cKy = Trim(Request("inpAdmissionID"))
  bedID = Trim(Request("inpBedID"))

  ins = Trim(Request("inpInsuranceSchemeID"))
  dt = Trim(Request("inpAdmissionDate"))
  dDt = Trim(Request("inpDischargeDate"))
  recNo = Trim(Request("inpAdmissionInfo1")) 'Receipt No.
  recTyp = Trim(Request("inpReceiptTypeID")) 'Receipt Type.
  aSt = Trim(Request("inpAdmissionStatusID"))
  bd = Trim(Request("inpBedCharge"))
  dy = Trim(Request("inpNoOfDays"))
  amtval = 0
  If IsNumeric(bd) And IsNumeric(dy) Then
    amtval = CDbl(CDbl(bd) * CDbl(dy))
  End If
  amtTbl = "Admission"
  amtfld = "AdmissionAmt1"
  kyFld = "BedID"
  ' rNo = GetTreatReceiptNo(eps)
  ' recNo = recNo & "," & rNo
  recNo = GetAllReceiptNo(eps) ''@bless - 06 Feb 2023 >> Include all receipts

  spGrp = GetComboNameFld("Visitation", eps, "SpecialistGroupID")

  If Not VisitBillPeriodValid(eps) Then
    vld = False
  ' ElseIf VisitHasExistingAdmission(eps) Then ''@bless - to do
  '   vld = False
  'ElseIf DischargePendingPay() Then
    'vld = False

  ElseIf Glob_IsWalkinConsultation(spGrp) Then
    vld = False
    SetPageMessages "Policy Alert! Admitting this folder [Guest/Walk-in Visit] is not allowed"
  ElseIf UCase(initVst) = "SUB" Then
    vld = False
    SetPageMessages "POLICY ALERT:: Cannot admit on a sub encounter. For billing purpose ONLY "
  ElseIf Not AllowGenderAgeRestriction(eps, bedID) Then
    vld = False
  ElseIf OnActiveAdmission(pat, eps) And UCase(tblMode) = UCase("NEWMODE") Then
    vld = False
  ElseIf OnAdmission(eps) And UCase(tblMode) = UCase("NEWMODE") Then
      vld = False
      SetPageMessages "Patient already on Admission"
  ElseIf (Len(recNo) > 0) And Not ReceiptNosValid(recNo) Then
      vld = False
  ElseIf HasAdmitExempt(ins, amtTbl, kyFld) Then
    vld = False
  ElseIf Not DischargeValid(aSt, dDt, dy) Then
    vld = False
  ElseIf UCase(aSt) = "A001" Then 'On Admission
    vld = True
  ElseIf Not HasPatientPaidByVal(amtTbl, amtfld, amtval, kyFld, recNo, recTyp, pat, cKy, "ADMISSION", dt, eps, dDt) Then
    vld = False
  End If

  ' If UCase("admin") = UCase(uName) Or UCase(jSchd) = UCase(uName) Then
  '   vld = False
  '   SetPageMessages "[Administrator] User Account cannot do transaction"
  ' End If

  If Not vld Then
    If objPage.rtnHdlProcessPoint Then
       objPage.hdlProcessPoint = False
    End If
  End If
End If
End Sub


Function GetTreatReceiptNo(vst)
    Dim rs, ot, sql, rNo, cnt
    Set rs = CreateObject("ADODB.Recordset")
    ot = ""
    cnt = 0

    sql = "select * from ConsultReview where visitationid='" & vst & "'"
    With rs
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .movefirst
            Do While Not .EOF
                If Not IsNull(.fields("KeyPrefix")) Then
                    rNo = Trim(.fields("KeyPrefix"))
                    If Len(rNo) > 0 Then
                        ot = ot & "," & rNo
                    End If
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rs = Nothing
    GetTreatReceiptNo = ot
End Function

''@bless - 06 Feb 2023 //All receipts
Function GetAllReceiptNo(vst)
    Dim rs, ot, sql, rNo, cnt
    Set rs = CreateObject("ADODB.Recordset")
    ot = "'INVALID-RECEIPT-ID'"
    ''Visitation
    sql = "select * from Visitation where visitationid='" & vst & "'"
    With rs
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          If Not IsNull(.fields("VisitInfo1")) Then
            rNo = Trim(.fields("VisitInfo1"))
            If Len(rNo) > 0 Then
              If Len(Trim(ot)) > 0 Then
                ot = ot & ",'" & rNo & "'"
              Else
                ot = "'" & rNo & "'"
              End If
            End If
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With

    ''Admission
    sql = "select * from Admission where visitationid='" & vst & "'"
    With rs
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          If Not IsNull(.fields("AdmissionInfo1")) Then
            rNo = Trim(.fields("AdmissionInfo1"))
            If Len(rNo) > 0 Then
              If Len(Trim(ot)) > 0 Then
                ot = ot & ",'" & rNo & "'"
              Else
                ot = "'" & rNo & "'"
              End If
            End If
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With

    ''DrugSale
    sql = "select * from DrugSale where visitationid='" & vst & "'"
    With rs
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          If Not IsNull(.fields("MainInfo2")) Then
            rNo = Trim(.fields("MainInfo2"))
            If Len(rNo) > 0 Then
              If Len(Trim(ot)) > 0 Then
                ot = ot & ",'" & rNo & "'"
              Else
                ot = "'" & rNo & "'"
              End If
            End If
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With

    ''LabRequest
    sql = "select * from LabRequest where visitationid='" & vst & "'"
    With rs
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          If Not IsNull(.fields("ReceiptInfo1")) Then
            rNo = Trim(.fields("ReceiptInfo1"))
            If Len(rNo) > 0 Then
              If Len(Trim(ot)) > 0 Then
                ot = ot & ",'" & rNo & "'"
              Else
                ot = "'" & rNo & "'"
              End If
            End If
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With

    ''ConsultReview
    sql = "select * from ConsultReview where visitationid='" & vst & "'"
    With rs
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          If Not IsNull(.fields("KeyPrefix")) Then
            rNo = Trim(.fields("KeyPrefix"))
            If Len(rNo) > 0 Then
              If Len(Trim(ot)) > 0 Then
                ot = ot & ",'" & rNo & "'"
              Else
                ot = "'" & rNo & "'"
              End If
            End If
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
    rec = ""
    ''Receipt
    sql = "select * from Receipt where ReceiptID IN (" & ot & ") "
    With rs
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          If Not IsNull(.fields("ReceiptID")) Then
            rNo = Trim(.fields("ReceiptID"))
            If Len(rNo) > 0 Then
              If Len(Trim(rec)) > 0 Then
                rec = rec & "," & rNo
              Else
                rec = rNo
              End If
            End If
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rs = Nothing
    GetAllReceiptNo = rec
End Function


Function AllowGenderAgeRestriction(vst, bedID)
    Dim rst, sql, ot, age, gend, bdMd
    Set rst = CreateObject("ADODB.Recordset")
    ot = True
    sql = "select * from Visitation where VisitationID='" & vst & "' "
    ' sql = sql & " and ='" & "" & "' "
    bdMd = GetComboNameFld("Bed", bedID, "BedModeID")
    Const MALE = "GEN01"
    Const FML = "GEN02"
    Const NOGEND = "G001"
    Const ADLT = "01"
    Const PAED = "02"
    Const BABY = "03"
    msg = "RESTRICTION ALERT:: " & bdMd

    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        If Not IsNull(rst.fields("PatientID")) Then
            age = rst.fields("PatientAge")
            gend = rst.fields("GenderID")
            gendNm = UCase(GetComboName("Gender", gend))
            ageNm = rst.fields("VisitInfo6")
            ' msg = msg & " >> " & gend & " >> " & age & " >> "
            ''ADULT
            If UCase(bdMd) = UCase(MALE & ADLT) Then ''Male, Adult
              If UCase(gend) <> MALE Then
                ot = False
                msg = msg & "MALE" & " Bed against " & gendNm & " Patient"
              ElseIf CDbl(age) < 12 Then
                ot = False
                msg = msg & "ADULT " & " Bed against " & ageNm & " Old Patient"
              End If
            ElseIf UCase(bdMd) = UCase(FML & ADLT) Then ''Female, Adult
              If UCase(gend) <> FML Then
                ot = False
                msg = msg & "FEMALE" & " Bed against " & gendNm & " Patient"
              ElseIf CDbl(age) < 12 Then
                ot = False
                msg = msg & "ADULT " & " Bed against " & ageNm & " Old Patient"
              End If
            ElseIf UCase(bdMd) = UCase(FML) Then ''Female
              If UCase(gend) <> FML Then
                ot = False
                msg = msg & "FEMALE" & " Bed against " & gendNm & " Patient"
              End If
            ''PAEDIATRIC
            ElseIf UCase(bdMd) = UCase(MALE & PAED) Then ''Male, Paediatric
              If UCase(gend) <> MALE Then
                ot = False
                msg = msg & "MALE" & " Bed against " & gendNm & " Patient"
              ElseIf CDbl(age) >= 12 Then
                ot = False
                msg = msg & "PAEDIATRIC " & " Bed against " & ageNm & " Old Patient"
              End If
            ElseIf UCase(bdMd) = UCase(FML & PAED) Then ''Female, Paediatric
              If UCase(gend) <> FML Then
                ot = False
                msg = msg & "FEMALE" & " Bed against " & gendNm & " Patient"
              ElseIf CDbl(age) >= 12 Then
                ot = False
                msg = msg & "PAEDIATRIC " & " Bed against " & ageNm & " Old Patient"
              End If
            ElseIf UCase(bdMd) = UCase(FML) Then ''Female
              If UCase(gend) <> FML Then
                ot = False
                msg = msg & "FEMALE" & " Bed against " & gendNm & " Patient"
              End If
            ''BABY
            ElseIf UCase(bdMd) = UCase(MALE & BABY) Then ''Male, baby
              If UCase(gend) <> MALE Then
                ot = False
                msg = msg & "MALE" & " Bed against " & gendNm & " Patient"
              ElseIf CDbl(age) > 6 Then
                ot = False
                msg = msg & "BABY " & " Bed against " & ageNm & " Old Patient"
              End If
            ElseIf UCase(bdMd) = UCase(FML & BABY) Then ''Female, baby
              If UCase(gend) <> FML Then
                ot = False
                msg = msg & "FEMALE" & " Bed against " & gendNm & " Patient"
              ElseIf CDbl(age) > 6 Then
                ot = False
                msg = msg & "BABY " & " Bed against " & ageNm & " Old Patient"
              End If
            ElseIf UCase(bdMd) = UCase(FML) Then ''Female
              If UCase(gend) <> FML Then
                ot = False
                msg = msg & "FEMALE" & " Bed against " & gendNm & " Patient"
              End If
            ''NO GENDER
            ElseIf UCase(bdMd) = UCase(NOGEND & BABY) Then ''no gender, baby
              If CDbl(age) > 6 Then
                ot = False
                msg = msg & "BABY " & " Bed against " & ageNm & " Old Patient"
              End If
            ElseIf UCase(bdMd) = UCase(NOGEND & PAED) Then ''no gender, paed
              If CDbl(age) >= 12 Then
                ot = False
                msg = msg & "PAEDIATRIC " & " Bed against " & ageNm & " Old Patient"
              End If
            ElseIf UCase(bdMd) = UCase(NOGEND & ADULT) Then ''no gender, adult
              If CDbl(age) < 12 Then
                ot = False
                msg = msg & "ADULT" & " Bed against " & ageNm & " Old Patient"
              End If
            ''MALE
            ElseIf UCase(bdMd) = UCase(MALE) Then ''Male Only
              If UCase(gend) <> MALE Then
                ot = False
                msg = msg & "MALE" & " Bed against " & gendNm & " Patient"
              End If
            ''FEMALE
            ElseIf UCase(bdMd) = UCase(FML) Then ''Female Only
              If UCase(gend) <> FML Then
                ot = False
                msg = msg & "FEMALE" & " Bed against " & gendNm & " Patient"
              End If
            End If
        End If
    End If
    rst.Close
    If Not ot Then
      SetPageMessages msg
    End If
    Set rst = Nothing
    AllowGenderAgeRestriction = ot
End Function

'On Admission
Function OnActiveAdmission(pat, vst)
  Dim ot, rst, sql
  Set rst = CreateObject("ADODB.Recordset")
  ot = False
  With rst
  sql = "select admissionid, visitationid from admission where PatientID='" & pat & "' "
  sql = sql & " and (admissionstatusid='A001' or admissionstatusid='A007')"
  sql = sql & " and visitationid<>'" & vst & "'"

  sql = "select admissionid, visitationid, WardID, AdmissionDate from admission where (admissionstatusid='A001' or admissionstatusid='A007') "
  sql = sql & " and patientid='" & pat & "' order by AdmissionDate "

  ''25 May 2022
  sql = "select admissionid, visitationid, WardID, AdmissionDate from admission where (AdmissionStatusID IN ('A001','A007')) "
  ' sql = sql & " and patientid='" & pat & "' And (VisitationID<>'" & vst & "-C' Or VisitationID<>'" & vst & "') order by AdmissionDate "
  sql = sql & " and patientid='" & pat & "' And (VisitationID<>'" & vst & "' Or VisitationID<>AdmissionID) order by AdmissionDate "

  ''30 May 2022
  sql = "select admissionid, visitationid, WardID, AdmissionDate from admission where (AdmissionStatusID IN ('A001','A007')) "
  sql = sql & " and patientid='" & pat & "' And (VisitationID<>'" & vst & "' Or VisitationID<>AdmissionID) "
  sql = sql & " And VisitationID NOT LIKE '%-C' order by AdmissionDate "
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    ot = True
    msg = "POLICY ALERT :: Patient is having an ACTIVE admission on Visit No.: [" & UCase(.fields("visitationid")) & "] "
    msg = msg & " @ " & GetComboName("Ward", rst.fields("WardID")) & " since " & FormatDate(rst.fields("AdmissionDate"))
    SetPageMessages msg
  ' Else
  '   ' ot = True
  '   ' SetPageMessages "No passed"
  End If
  .Close
  End With
  OnActiveAdmission = ot
  Set rst = Nothing
End Function

Function VisitHasExistingAdmission(eps)
  Dim rst, sql, ot
  Set rst = CreateObject("ADODB.Recordset")
  ot = False

  If UCase(tblMode) = "NEWMODE" Then
    ' lstWhCls2 = " and (a.TransProcessStatID='T009' or a.TransProcessStatID='T010' or a.TransProcessStatID='T013' or a.TransProcessStatID='T004' or a.TransProcessStatID='T005' or a.TransProcessStatID='T006' or a.TransProcessStatID='T007' or a.TransProcessStatID='T014')"
    lstWhCls2 = " and IsDate(a.DischargeDate) <= 0 "
    sql = "select a.visitationID,a.WardID,a.BedID,a.PatientID,a.MainInfo3,a.TransProcessStatID,a.AdmissionDate,a.DischargeDate,a.AdmissionID,a.InsuranceSchemeID,a.TransProcessValID "
    sql = sql & " from Admission as a,Visitation as v "
    sql = sql & " where v.visitationID=a.visitationID and v.VisitationID='" & eps & "' " & lstWhCls2
    sql = sql & " And v.InitialVisitationID<>'SUB' " & lstWhCls2
    '29 Apr 2020 @bless - Fix no discharge dates'
    sql = sql & " and a.AdmissionStatusID<>'A003' order by a.AdmissionDate desc, a.WardID "
    ' sql = sql & " and (a.AdmissionStatusID<>'A003' or (a.AdmissionStatusID='A003' and a.AdmissionDate >= v.VisitDate)) "
    ' sql = sql & " order by a.AdmissionDate desc, a.WardID "

    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        rst.movefirst

        If Not IsNull(rst.fields("WardID")) Then
            ot = True
            wd = rst.fields("WardID")
            dt = rst.fields("AdmissionDate")
            SetPageMessages "There's already a valid admission for this visit at [" & GetComboName("Ward", wd) & "] since " & FormatDate(dt) & " "
        End If
    End If
    rst.Close
  End If
  SetPageMessages "Exists: " & ot & " :: " & sql

  ot = True
  Set rst = Nothing
  VisitHasExistingAdmission = ot
End Function


Function DischargePendingPay()
  Dim strAdmissionID, strCurrenstStateID, strAdmissionStatusID, ot
  ot = False
  strAdmissionID = Trim(Request("inpAdmissionID"))
  strAdmissionStatusID = Trim(Request("inpAdmissionStatusID"))
  strCurrenstStateID = GetComboNameFld("Admission", strAdmissionID, "TransProcessStatID")
  If UCase(strCurrenstStateID) = "T002" Then ' discharged, pending payment
      SetPageMessages "Update to [Discharged Pending Payment] not allowed."
      ot = True
  Else
    If UCase(strAdmissionStatusID) = "A007" Then ' discharged, pending payment
      SetPageMessages "[Discharged Pending Payment] not allowed from Admission."
      ot = True
    End If
  End If
  DischargePendingPay = ot
End Function
Function DischargeValid(aSt, dDt, dy)
  Dim ot, cDt
  cDt = CStr(Now())
  ot = True
  If UCase(aSt) <> UCase("A001") Then
    If Not IsDate(dDt) Then
      ot = False
      SetPageMessages "The Discharge Date [" & FormatDateDetail(dDt) & "] is not VALID"
    ElseIf Not IsNumeric(dy) Then
      ot = False
      SetPageMessages "The Number of Days [" & dy & "] must be a NUMBER"
    ElseIf CDbl(dy) < 0 Then
      ot = False
      SetPageMessages "The Number of Days [" & dy & "] must be a POSITIVE NUMBER"
    ElseIf CDate(dDt) > CDate(cDt) Then
      ot = False
      SetPageMessages "The Discharge Date [" & FormatDateDetail(dDt) & "] cannot be later than Current Date [" & FormatDateDetail(cDt) & "]"
    End If
    ''@bless - 16 Jun 2022 //Block any status for doctor
    If Left(UCase(jSchd), 3) = UCase("M03") Then
      ot = False
      SetPageMessages "Admission status must be [On Admission] "
    End If
  End If
  DischargeValid = ot
End Function
Function BillMonthValidVisit(vst)
  Dim ot, bPrd
  ot = False
  If Len(Trim(vst)) > 0 Then
    bPrd = Trim(GetComboNameFld("Visitation", Trim(vst), "BillMonthID"))
    If (UCase(bPrd) = "NONE") Or (UCase(bPrd) = "") Then
      ot = True
    Else
      SetPageMessages "This Visit No. [" & UCase(vst) & "] has been Billed. It can not be used in this transaction."
    End If
  End If
  BillMonthValidVisit = ot
End Function
Function VisitBillPeriodValidOLD(vst)
  Dim ot, bPrd
  ot = False
  If Len(Trim(vst)) > 0 Then
    bPrd = Trim(GetComboNameFld("Visitation", Trim(vst), "BillMonthID"))
    If UCase(bPrd) = "NONE" Then
      ot = True
    Else
      SetPageMessages "This Visit No. [" & UCase(vst) & "] has been Billed. It can not be used in this transaction."
    End If
  End If
  VisitBillPeriodValidOLD = ot
End Function
Function VisitBillPeriodValid(vst)
  Dim rst, rst3, sql, dt, ot, vMd, vGrp, vMth, bMth
  Set rst = CreateObject("ADODB.Recordset")
  Set rst3 = CreateObject("ADODB.Recordset")
  ot = False
  If Len(Trim(vst)) > 0 Then
    With rst
      sql = "select * from visitation where visitationid='" & vst & "'"
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        dt = FormatDateDetail(Now())
        vGrp = .fields("VettingGroupID")
        vMth = .fields("WorkingMonthID")
        bMth = Trim(.fields("BillMonthID"))
        'If UCase(bMth) = "NONE" Then '11 Dec 2014
          'ot = True
          'sql = "select * from BillPeriod where BillPeriodStatusid='B001' and BillPeriodDate>='" & CStr(dt) & "' order by BillPeriodDate desc"
          sql = "select * from BillPeriod where BillPeriodStatusid='B001' and BillPeriodTypeID='" & vGrp & "'"
          sql = sql & " and BillMonthID='" & vMth & "' and BillPeriodDate>='" & CStr(dt) & "' order by BillPeriodDate desc"
          rst3.maxrecords = 1
          rst3.open qryPro.FltQry(sql), conn, 3, 4
          If rst3.RecordCount > 0 Then
            rst3.movefirst
            ot = True
          Else
            If OnAdmission(vst) Then
              ot = True
            Else
              SetPageMessages "The Billing Period for this Visit No. [" & UCase(vst) & "] has been closed. Please contact the Claim Manager."
            End If
          End If
          rst3.Close
        'Else
        '  SetPageMessages "This Visit No. [" & UCase(vst) & "] has been Billed. Please contact the Claim Manager to open it for transactions."
        'End If
      Else
       ot = True
      End If
      .Close
    End With
  End If
  VisitBillPeriodValid = ot
  Set rst = Nothing
  Set rst2 = Nothing
  Set rst3 = Nothing
End Function
'On Admission
Function OnAdmission(vst)
  Dim ot, rst, sql
  Set rst = CreateObject("ADODB.Recordset")
  ot = False
  With rst
  sql = "select visitationid from admission where visitationid='" & vst & "' and (admissionstatusid='A001' or admissionstatusid='A007')"
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  ot = True
  End If
  .Close
  End With
  OnAdmission = ot
  Set rst = Nothing
End Function
Function IsNHISDrug(drg)
  Dim ot, ky
  ot = False
  ky = Trim(drg)
  If Len(ky) = 9 Then
    If IsNumeric(Right(ky, 1)) Then
      ot = True
    End If
  End If
  IsNHISDrug = ot
End Function
Function VisitNoExempt1(vst)
Dim arr, ul, num, lst, ot
ot = False
lst = "E01||E02"
arr = Split(lst, "||")
ul = UBound(arr)
For num = 0 To ul
If UCase(Trim(arr(num))) = UCase(Trim(vst)) Then
ot = True
Exit For
End If
Next
VisitNoExempt1 = ot
End Function
Function VisitNoExempt2(vst)
Dim arr, ul, num, lst, ot
ot = False
lst = "0054444||0079669||0185870||0208211||0158043||0099346||0186748||0207940||P1||P2"
arr = Split(lst, "||")
ul = UBound(arr)
For num = 0 To ul
If UCase(Trim(arr(num))) = UCase(Trim(vst)) Then
ot = True
Exit For
End If
Next
VisitNoExempt2 = ot
End Function

Function AdmitExemptExist(drg, ins)
  Dim ot, ky1, ky2, bDt, bDt1, bDt2, gen, patID, rst, sql, qty, rQty
  ot = False
  Set rst = CreateObject("ADODB.Recordset")

  sql = "SELECT * FROM InsurAdmitExempt WHERE InsuranceSchemeid='" & ins & "' AND " & "Bedid='" & drg & "'"
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      ot = True
    End If
    .Close
  End With
  AdmitExemptExist = ot
  Set rst = Nothing
End Function

'/////////////////////////
Function HasPatientPaidByVal(amtTbl, amtfld, amtval, kyFld, recNo, recTyp, pat, kVl, kTb, bDt, vst, dDt)
  Dim ot, pos, pos1, pos2, rTyp, fLst, vLst, arr, ul, num, vl, amt, totAmt
  Dim vld, balAmt, totBal, rec, arr2, ul2, num2, PID, currAmt
  ot = False
  rTyp = Trim(recTyp)
  If UCase(rTyp) = "R002" Then 'Credit
    ot = True
  Else
    totAmt = 0
    vld = True
    amt = CStr(amtval)
    totAmt = GetAdmissionAmount(pat, recNo, vst, dDt)
    If Len(amt) > 0 Then
      If IsNumeric(amt) Then
        totAmt = totAmt + CDbl(amt)
      Else
        SetPageMessages amtfld & " is not a number."
        vld = False
      End If
    Else
      SetPageMessages amtfld & " is blank."
      vld = False

    End If
    If vld Then
      totBal = 0
     
      totBal = Round(totBal, 2)
      totAmt = Round(totAmt, 2)
      If CDbl(totBal) >= CDbl(totAmt) Then
        ot = True
      Else
        If HasPatientCredit(pat, bDt) Then
          ot = True
        Else
          If ul2 < 0 Then
            SetPageMessages "No Receipt No. has been given for the current CASH transaction."
          End If
          'SetPageMessages "The current balanace [" & FormatNumber(totBal, 2, , , -1) & "] on Receipt No.[" & recNo & "] cannot pay for the Bill Amount[" & FormatNumber(totAmt, 2, , , -1) & "]"
          SetPageMessages "The current Receipt No.[" & recNo & "] cannot pay for the outstanding balance of [" & FormatNumber(totAmt, 2, , , -1) & "]"
          ot = False
          ot = True ''@bless - 02 Dec 2023 //Allow
        End If
      End If
    End If
  End If
  HasPatientPaidByVal = ot
End Function
Function HasPatientCredit(PID, bDt)
  Dim rs, ot, sql, dt
    Set rs = CreateObject("ADODB.Recordset")
    ot = False
    dt = CStr(bDt)
    sql = "select patientid from PatientCredit where patientid='" & PID & "' and PatientCreditStatID='P002' and ReceiptDate2>='" & dt & "'"
    With rs
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        ot = True
      End If
      .Close
    End With
    Set rs = Nothing
    HasPatientCredit = ot
End Function
'GetUsedByCurrTrans
Function GetUsedByCurrTrans(rec, kVl, kTb)
  Dim ot, rst, sql, bid, rst2
  Set rst = CreateObject("ADODB.Recordset")
  Set rst2 = CreateObject("ADODB.Recordset")
  ot = 0
  With rst
    sql = "select PatientBillID from PatientBill where keyprefix='" & kVl & "' and tableID='" & kTb & "'"
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      Do While Not .EOF
        bid = .fields("PatientBillID")
        sql = "select paidamount from PatientReceipt2 where ReceiptID='" & rec & "' and PatientBillID='" & bid & "'"
        rst2.open qryPro.FltQry(sql), conn, 3, 4
        If rst2.RecordCount > 0 Then
          rst2.movefirst
          ot = ot + CDbl(rst2.fields("paidamount"))
        End If
        rst2.Close
        .MoveNext
      Loop
    End If
    .Close
  End With
  GetUsedByCurrTrans = ot
  Set rst = Nothing
  Set rst2 = Nothing
End Function
'/////////////////////////
Function HasAdmitExempt(ins, tbl, fld)
  Dim ot, pos, pos1, pos2, fLst, vLst, arr, ul, num
  Dim vld, arr2, ul2, num2
  ot = False
  vl = Trim(Request("inp" & fld))
  If Len(vl) > 0 Then
    If AdmitExemptExist(vl, ins) Then
      ot = True
      SetPageMessages "Admission->" & GetComboName("Bed", vl) & "[" & UCase(vl) & "] is exempted by [" & GetComboName("InsuranceScheme", ins) & "]"
    End If
  Else
    SetPageMessages fld & " is blank."
  End If
  HasAdmitExempt = ot
End Function
Function ApplyPrescribeCheck(ins)
  Dim ot, inZn
  ot = False
  If Len(Trim(ins)) > 0 Then
    inZn = GetComboNameFld("InsuranceScheme", ins, "InsuranceZoneID")
    If UCase(Trim(Right(inZn, 4))) = "-INS" Then 'Insurance
      ot = True
    ElseIf UCase(Trim(Right(inZn, 6))) = "-SSNIT" Then 'SSNIT
      ot = True
    End If
  End If
  ApplyPrescribeCheck = ot
End Function
'////////////////////
Function GetArrVal(lst, dlm, pos)
Dim arr, ul, num, ot
ot = ""
arr = Split(lst, dlm)
ul = UBound(arr)
If ul >= pos Then
    ot = arr(pos)
End If
GetArrVal = ot
End Function

Function GetAdmissionAmount(pat, recNo, vst, dDt)
  Dim ot, vCst
  ot = 0
  vCst = GetComboNameFld("Visitation", vst, "VisitCost")
  If IsNumeric(vCst) Then
    ot = ot + CDbl(vCst)
  End If
  'ot=ot + GetAdmission(vst)
  ot = ot + GetDrug(vst)
  ot = ot + GetNonDrug(vst)
  ot = ot + GetLab(vst)
  'ot=ot + GetXray(vst)
  ot = ot + GetTreat(vst)
  ot = ot - GetPayments(pat, recNo, vst, dDt)
  ot = ot - GetWaiver(vst)
  ot = ot + GetUsedPayments(pat, recNo, vst, dDt)
  GetAdmissionAmount = ot
End Function
'GetWaiver
Function GetWaiver(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select sum(PaidAmount) as tot from patientwaiveritems where visitationid='" & vst & "'"
  ot = 0
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
    .movefirst

      Do While Not .EOF
      If Not IsNull(.fields("tot")) Then
        tot = .fields("tot")
        ot = ot + tot
      End If
      .MoveNext
      Loop

    End If
    .Close
  End With
  GetWaiver = ot
  Set rst = Nothing
End Function
'GetPayments
Function GetPayments(pat, recNo, vst, dDt)
Dim rst, sql, ot, cnt, hdr, dsc, pd, cn, bal, dt, sDt, eDt, cn2
Dim arr, ul, num, whcls, r, rCnt, sqlOk, sql2
Set rst = CreateObject("ADODB.Recordset")
ot = 0
cnt = 0
dt = GetComboNameFld("Visitation", vst, "VisitDate")
sDt = FormatDate(dt) & " 0:00:00"
eDt = Now()
If Not IsNull(dDt) Then
  If IsDate(dDt) Then
    If CDate(dDt) > CDate(sDt) Then
      eDt = FormatDate(dDt) & " 23:59:59"
    End If
  End If
End If
'Receipt No
sqlOk = False
If (UCase(pat) = "P1") Or (UCase(pat) = "P2") Then
  whcls = ""
  arr = Split(recNo, ",")
  ul = UBound(arr)
  rCnt = 0
  For num = 0 To ul
    r = Trim(arr(num))
    If Len(r) > 0 Then
      sqlOk = True
      rCnt = rCnt + 1
      If rCnt = 1 Then
        whcls = whcls & " where "
      Else
        whcls = whcls & " or "
      End If
      whcls = whcls & " (PatientID='" & pat & "' and ReceiptID='" & r & "') "
    End If
  Next
  If Len(Trim(whcls)) > 0 Then
    sqlOk = True
    sql = "select * from Receipt "
    sql = sql & " " & whcls
    sql = sql & " order by receiptDate"
  End If
Else
'  sqlOk = True
'  whcls = ""
'  arr = Split(recNo, ",")
'  ul = UBound(arr)
'  For num = 0 To ul
'    r = Trim(arr(num))
'    If Len(r) > 0 Then
'      whcls = whcls & " or (PatientID='" & pat & "' and ReceiptID='" & r & "') "
'    End If
'  Next
'  sql = "select * from Receipt where (Patientid='" & pat & "' and receiptdate between '" & sDt & "' and '" & eDt & "')"
'  sql = sql & " " & whcls
'  sql = sql & " order by receiptDate"

  sqlOk = True
  sql2 = "select distinct receiptid from patientreceipt2 where visitationid='" & vst & "' and PatientID='" & pat & "'"
  whcls = ""
  arr = Split(recNo, ",")
  ul = UBound(arr)
  For num = 0 To ul
    r = Trim(arr(num))
    If Len(r) > 0 Then
      whcls = whcls & " or (PatientID='" & pat & "' and ReceiptID='" & r & "') "
    End If
  Next
  sql = "select * from Receipt where (receiptID in (" & sql2 & "))"
  sql = sql & " " & whcls
  sql = sql & " order by receiptDate"
End If

cnt = 0
If sqlOk Then
  With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  .movefirst

  Do While Not .EOF
  cnt = cnt + 1

  dsc = .fields("Remarks")
  pd = .fields("ReceiptAmount1")
  cn = .fields("paidamounT")
  If CDbl(cn) = 0 Then
    cn2 = "-"
  ElseIf CDbl(cn) < 0 Then
    cn = 0
    cn2 = "-"
  Else
    cn2 = FormatNumber(CStr(cn), 2, , , -1)
  End If
  bal = CDbl(pd) - CDbl(cn)
  ot = ot + bal
  .MoveNext
  Loop
  End If
  .Close
  End With
End If 'Sql
GetPayments = ot
Set rst = Nothing
End Function
'GetUsedPayments
Function GetUsedPayments(pat, recNo, vst, dDt)
Dim rst, rst2, sql, ot, cnt, hdr, dsc, pd, cn, bal, dt, sDt, eDt
Dim cnt2, cn2, rec, usd, uCnt, sql2
Dim arr, ul, num, whcls, r, rCnt, sqlOk
Set rst = CreateObject("ADODB.Recordset")
Set rst2 = CreateObject("ADODB.Recordset")
ot = 0
cnt = 0
dt = GetComboNameFld("Visitation", vst, "VisitDate")
sDt = FormatDate(dt) & " 0:00:00"
eDt = Now()
uCnt = 0
If Not IsNull(dDt) Then
  If IsDate(dDt) Then
    If CDate(dDt) > CDate(sDt) Then
      eDt = FormatDate(dDt) & " 23:59:59"
    End If
  End If
End If

'Receipt No
sqlOk = False
If (UCase(pat) = "P1") Or (UCase(pat) = "P2") Then
  whcls = ""
  arr = Split(recNo, ",")
  ul = UBound(arr)
  rCnt = 0
  For num = 0 To ul
    r = Trim(arr(num))
    If Len(r) > 0 Then
      sqlOk = True
      rCnt = rCnt + 1
      If rCnt = 1 Then
        whcls = whcls & " where "
      Else
        whcls = whcls & " or "
      End If
      whcls = whcls & " (PatientID='" & pat & "' and ReceiptID='" & r & "') "
    End If
  Next
  If Len(Trim(whcls)) > 0 Then
    sqlOk = True
    sql = "select * from Receipt "
    sql = sql & " " & whcls
    sql = sql & " order by receiptDate"
  End If
Else
'  sqlOk = True
'  whcls = ""
'  arr = Split(recNo, ",")
'  ul = UBound(arr)
'  For num = 0 To ul
'    r = Trim(arr(num))
'    If Len(r) > 0 Then
'      whcls = whcls & " or (PatientID='" & pat & "' and ReceiptID='" & r & "') "
'    End If
'  Next
'  sql = "select * from Receipt where (Patientid='" & pat & "' and receiptdate between '" & sDt & "' and '" & eDt & "')"
'  sql = sql & " " & whcls
'  sql = sql & " order by receiptDate"

  sqlOk = True
  sql2 = "select distinct receiptid from patientreceipt2 where visitationid='" & vst & "' and PatientID='" & pat & "'"
  whcls = ""
  arr = Split(recNo, ",")
  ul = UBound(arr)
  For num = 0 To ul
    r = Trim(arr(num))
    If Len(r) > 0 Then
      whcls = whcls & " or (PatientID='" & pat & "' and ReceiptID='" & r & "') "
    End If
  Next
  sql = "select * from Receipt where (receiptID in (" & sql2 & "))"
  sql = sql & " " & whcls
  sql = sql & " order by receiptDate"
End If 'P1,P2
If sqlOk Then
  cnt = 0
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst

      Do While Not .EOF
        cnt = cnt + 1
        rec = .fields("ReceiptID")
        dsc = ""
        pd = .fields("ReceiptAmount1")
        cn = .fields("paidamounT")
        If CDbl(cn) = 0 Then
          cn2 = "-"
        ElseIf CDbl(cn) < 0 Then
          cn = 0
          cn2 = "-"
        Else
          cn2 = FormatNumber(CStr(cn), 2, , , -1)
        End If
        bal = CDbl(pd) - CDbl(cn)
        usd = 0
        cnt2 = 0
        sql = "select * from PatientReceipt2 where Receiptid='" & rec & "' and VisitationID<>'" & vst & "' order by receiptDate2"
        rst2.open qryPro.FltQry(sql), conn, 3, 4
        If rst2.RecordCount > 0 Then
          rst2.movefirst
          Do While Not rst2.EOF
            cnt2 = cnt2 + 1
            usd = usd + rst2.fields("PaidAmount")
            rst2.MoveNext
          Loop
        End If
        rst2.Close
        ot = ot + usd
        .MoveNext
      Loop
    End If
    .Close
  End With
End If 'SqlOk
GetUsedPayments = ot
Set rst = Nothing
Set rst2 = Nothing
End Function
'GetAdmission
Function getAdmission(vst)
  Dim rst, sql, ot, cnt, hdr, adm, chg, dys
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select * from admission where visitationid='" & vst & "'"
  ot = 0
  cnt = 0
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst

      Do While Not .EOF
        cnt = cnt + 1
        adm = .fields("admissionid")
        chg = 0 '.fields("bedcharge")
        dys = 0 '.fields("noofdays")
        ot = ot + (chg * dys)
        .MoveNext
      Loop

    End If
    .Close
  End With
  getAdmission = ot
  Set rst = Nothing
End Function
'GetDrug
Function GetDrug(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot, rQty, drg, fQty, tot2, rTot
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select drugid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from drugsaleitems where visitationid='" & vst & "' group by drugid"
  ot = 0
  cnt = 0
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      'Pharmacy
      Do While Not .EOF
        cnt = cnt + 1
        unt = .fields("unt")
        qty = .fields("qty")
        drg = .fields("drugid")
        tot2 = .fields("tot") 'Addedd 1 Oct 2015

        rQty = GetReturnQty(vst, drg)
        fQty = qty - rQty
        If fQty > 0 Then
          If rQty > 0 Then 'Addedd 1 Oct 2015
            rTot = GetReturnTot(vst, drg)
            unt = (tot2 - rTot) / fQty
          End If
          tot = fQty * unt '.Fields("tot")
          ot = ot + tot
        End If

        .MoveNext
      Loop

    End If
    .Close
  End With
  GetDrug = ot
  Set rst = Nothing
End Function
'GetReturnTot   'Addedd 1 Oct 2015
Function GetReturnTot(vst, dg)
Dim rstTblSql, sql, ot
Set rstTblSql = CreateObject("ADODB.Recordset")
ot = 0
With rstTblSql
sql = "select sum(FinalAmt) as sm from drugreturnitems where visitationid='" & vst & "' and drugid='" & dg & "'"
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
.movefirst
ot = .fields("sm")
If IsNull(ot) Then
ot = 0
End If
End If
.Close
End With
Set rstTblSql = Nothing
GetReturnTot = ot
End Function
'GetReturnQty
Function GetReturnQty(vst, dg)
  Dim rstTblSql, sql, ot
  Set rstTblSql = CreateObject("ADODB.Recordset")
  ot = 0
  With rstTblSql
    sql = "select sum(returnqty) as sm from drugreturnitems where visitationid='" & vst & "' and drugid='" & dg & "'"
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      ot = .fields("sm")
      If IsNull(ot) Then
        ot = 0
      End If
    End If
    .Close
  End With
  Set rstTblSql = Nothing
  GetReturnQty = ot
End Function
'GetNonDrug
Function GetNonDrug(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select itemid,sum(qty) as qty,avg(retailunitcost) as unt,sum(finalamt) as tot from stockissueitems where visitationid='" & vst & "' group by itemid"
  ot = 0
  cnt = 0
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      Do While Not .EOF
        cnt = cnt + 1
        unt = .fields("unt")
        qty = .fields("qty")
        tot = .fields("tot")
        ot = ot + tot
        .MoveNext
      Loop

    End If
    .Close
  End With
  GetNonDrug = ot
  Set rst = Nothing
End Function

'GetLab
Function GetLab(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from investigation where visitationid='" & vst & "'"
  sql = sql & " group by labtestid"
  cnt = 0
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      Do While Not .EOF
        cnt = cnt + 1
        unt = .fields("unt")
        qty = .fields("qty")
        tot = .fields("tot")
        ot = ot + tot
        .MoveNext
      Loop

    End If
    .Close
  End With
  GetLab = ot
  Set rst = Nothing
End Function

'GetXRay
Function GetXRay(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from investigation where visitationid='" & vst & "'"
  sql = sql & " and (testcategoryid='T006' or testcategoryid='T007' or testcategoryid='T008') group by labtestid"
  ot = 0
  cnt = 0
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst

      Do While Not .EOF
        cnt = cnt + 1
        unt = .fields("unt")
        qty = .fields("qty")
        tot = .fields("tot")
        ot = ot + tot

        .MoveNext
      Loop

    End If
    .Close
  End With
  GetXRay = ot
  Set rst = Nothing
End Function

'GetTreat
Function GetTreat(vst)
  Dim rst, sql, ot, cnt, hdr, adm, unt, qty, tot
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select treatmentid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from treatcharges where visitationid='" & vst & "' group by treatmentid"
  cnt = 0
  ot = 0
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      Do While Not .EOF
        cnt = cnt + 1
        unt = .fields("unt")
        qty = .fields("qty")
        tot = .fields("tot")
        ot = ot + tot
        .MoveNext
      Loop
    End If
    .Close
  End With
  GetTreat = ot
  Set rst = Nothing
End Function
Function ReceiptNosValid(recNo)
  Dim arr, ul, num, lst, ot, arr2, ul2, num2, lst2, rec, cnt, recOk
  ot = False
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
              ot = False
              recOk = False
              SetPageMessages "Duplicate Receipt No.[" & rec & "] has been found. This is not allowed"
              Exit For
            End If
          Next
          If recOk Then
            cnt = cnt + 1
            If cnt > 1 Then
              lst2 = lst2 & ","
            End If
            lst2 = lst2 & rec
            ot = True
          Else
            ot = False
            Exit For
          End If
        Else
          ot = False
          Exit For
        End If
      Else
'        ot = False @ Peter - 19 Sep 2023 Set ot = True Was unable to save admission for testing purposes.
'        SetPageMessages "A blank Receipt No. after a comma(,).[" & lst2 & ",]"
ot = True
        Exit For
      End If
    Next
  Else
    ot = True
  End If
  ReceiptNosValid = ot
End Function
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
      SetPageMessages "The character [ " & ch & " ] cannot be used in Receipt No. [" & rec & "]"
      Exit For
    End If
  Next
  ReceiptCharValid = ot
End Function
Function ReceiptNameValid(recNo, patNm)
  Dim arr, ul, num, lst, ot, arr2, ul2, num2, lst2, rec, cnt, recOk, recNm
  ot = False
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
              ot = False
              recOk = False
              SetPageMessages "Duplicate Receipt No.[" & rec & "] has been found. This is not allowed"
              Exit For
            End If
          Next
          If recOk Then
            recNm = Trim(GetComboName("Receipt", rec))
            If (Len(recNm) > 0) And (Len(patNm) > 0) Then
              If HasTextMatch(recNm, patNm, 2) Then
                cnt = cnt + 1
                If cnt > 1 Then
                  lst2 = lst2 & ","
                End If
                lst2 = lst2 & rec
                ot = True
              Else
                ot = False
                SetPageMessages "The Receipt Name for [" & rec & "] does not match the Patient Name"
                Exit For
              End If
            Else
              ot = False
              SetPageMessages "Either the Receipt Name for [" & rec & "] or Patient Name is blank."
              Exit For
            End If
          Else
            ot = False
            Exit For
          End If
        Else
          ot = False
          Exit For
        End If
      Else
        ot = False
        SetPageMessages "A blank Receipt No. after a comma(,).[" & lst2 & ",]"
        Exit For
      End If
    Next
  Else
    ot = True
  End If
  ReceiptNameValid = ot
End Function
Function HasTextMatch(recNm, patNm, cnt)
  Dim ot, arr, ul, num, tot, fndCnt, nm, pos
  ot = False
  tot = 0
  fndCnt = 0
  arr = Split(patNm, " ")
  ul = UBound(arr)
  For num = 0 To ul
    nm = Trim(arr(num))
    If Len(nm) > 1 Then
      pos = InStr(1, UCase(recNm), UCase(nm))
      If pos > 0 Then
        fndCnt = fndCnt + 1
      End If
    End If
  Next
  If fndCnt >= cnt Then
    ot = True
  End If
  HasTextMatch = ot
End Function
Function HasTextMatch2(recNm, patNm, cnt)
  Dim ot, arr, ul, num, arr2, ul2, num2, tot, fndCnt, nm, pos, rNm
  ot = False
  tot = 0
  fndCnt = 0
  arr = Split(patNm, " ")
  ul = UBound(arr)

  arr2 = Split(recNm, " ")
  ul2 = UBound(arr2)

  For num = 0 To ul
    nm = Trim(arr(num))
    If Len(nm) > 1 Then
      For num2 = 0 To ul2
        rNm = Trim(arr2(num2))
        If Len(rNm) > 1 Then
          If UCase(nm) = UCase(rNm) Then
            fndCnt = fndCnt + 1
            Exit For
          End If
        End If
      Next
    End If
  Next
  If fndCnt >= cnt Then
    ot = True
  End If
  HasTextMatch2 = ot
End Function
Function ReceiptDateValid(recNo)
  Dim arr, ul, num, lst, ot, arr2, ul2, num2, lst2, rec, cnt, recOk, recDt, pat, hrs
  ot = False
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
              ot = False
              recOk = False
              SetPageMessages "Duplicate Receipt No.[" & rec & "] has been found. This is not allowed"
              Exit For
            End If
          Next
          If recOk Then
            recDt = Trim(GetComboNameFld("Receipt", rec, "ReceiptDate"))
            pat = Trim(GetComboNameFld("Receipt", rec, "PatientID"))
            If IsDate(recDt) And (Len(pat) > 0) Then
              hrs = DateDiff("h", CDate(recDt), Now())
              If (UCase(pat) = "P1") Or (UCase(pat) = "P2") Then 'Walk In
                If hrs <= 6 Then  '6 Hours
                  ot = True
                Else
                  ot = False
                  SetPageMessages "The Receipt Date[" & FormatDateDetail(recDt) & "] for [" & rec & "] is too old."
                  Exit For
                End If
              Else
                If hrs <= 720 Then '720 hours [30 Days]
                  ot = True
                Else
                  ot = False
                  SetPageMessages "The Receipt Date[" & FormatDateDetail(recDt) & "] for [" & rec & "] is too old."
                  Exit For
                End If
              End If
            Else
              ot = False
              SetPageMessages "Either the Receipt Date for [" & rec & "] is not Valid or Patient No is blank."
              Exit For
            End If
          Else
            ot = False
            Exit For
          End If
        Else
          ot = False
          Exit For
        End If
      Else
        ot = False
        SetPageMessages "A blank Receipt No. after a comma(,).[" & lst2 & ",]"
        Exit For
      End If
    Next
  Else
    ot = True
  End If
  ReceiptDateValid = ot
End Function
