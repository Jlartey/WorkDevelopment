'************************************************************************
'ProcessText
'************************************************************************
ProcessCode

Sub ProcessCode()
  Dim tbl, recExist, insSch, insno, inspat, vst, bd, adm, aSt
  tbl = UCase(GetPageVariable("TableName"))
  vst = Trim(request("inpVisitationID"))
  adm = Trim(request("inpAdmissionID"))
  aSt = Trim(request("inpAdmissionStatusID"))
  If (tbl = "ADMISSION") Then

    If UCase(GetComboNameFld("systemUser", uName, "BranchID")) <> "B001" Then 'check for all branches, except MAIN' @bash

      recExist = AdmitExist()
      bd = GetComboNameFld("Admission", adm, "BedID")
      If recExist Or (UCase(vst) = "E01") Or BedOccupied(vst, bd, aSt) Then
        If objPage.rtnHdlProcessPoint Then
           objPage.hdlProcessPoint = False
        End If
      End If
      
    End If

  End If
End Sub


Function BedOccupied(vst, bd, aSt)
Dim ot, ptId, patID, rst, sql, inspatid, vst2
Dim adm, adm2
ot = False
Set rst = CreateObject("ADODB.Recordset")
If UCase(aSt) = "A001" Then 'Seeking On Admission
  sql = "SELECT patientid,visitationID,Admissionid FROM admission WHERE visitationid<>'" & vst & "' and bedid='" & bd & "' and admissionstatusid='A001'"
  sql = sql & " and BedID<>'B082' and BedNoID<>'000'" 'B082 NHIS Detention,000->Waiting List
  With rst
    .open sql, conn, 3, 4
    If .RecordCount > 0 Then
      ot = True
      ptId = .fields("PatientID")
      adm2 = .fields("admissionid")
      vst2 = .fields("visitationid")
      SetPageMessages "A patient [" & ptId & "] with Visit No [" & vst2 & "] and Admission No [" & adm2 & "] is ocuppying the current bed. Discharge him/her first."
    End If
    .Close
  End With
End If
BedOccupied = ot
Set rst = Nothing
End Function

Function AdmitExist()
  Dim ot, ptId, patID, rst, sql, inspatid, vst2
  Dim vst, adm, adm2, pgMd

  '****** @bash: alllow discharges 30/may/2018'
  pgMd = objPage.pagemode
  If UCase(pgMd) = "EDITMODE" Then
      If UCase(request("inpAdmissionStatusID")) = "A002" Then
        AdmitExist = False
        Exit Function
      End If
  End If
  '********************'


  ot = True
  Set rst = server.CreateObject("ADODB.Recordset")
  vst = Trim(request("inpVisitationID"))
  adm = Trim(request("inpAdmissionID"))
  sql = "SELECT patientid,visitationID,admissionid FROM admission WHERE visitationid='" & vst & "' and PatientID<>'P1' and PatientID<>'P2'"
  With rst
    .open sql, conn, 3, 4
    If .RecordCount = 0 Then
      ot = False
    Else
      ptId = .fields("PatientID")
      adm2 = .fields("admissionid")
      If UCase(adm) <> UCase(adm2) Then
        SetPageMessages "A Patient [" & ptId & "] with Visit No. [" & vst & "] and Admission No. [" & adm2 & "] already exists.."
      Else
        ot = False
      End If
    End If
    .Close
  End With
  AdmitExist = ot
  Set rst = Nothing
End Function
