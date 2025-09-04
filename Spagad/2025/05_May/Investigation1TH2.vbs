'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Sub DisplayHead2()
AddReportHeader
Dim str
str = ""
str = str & "<tr><td align=""center"" height=""100""></td></tr>" '"<tr><td align=""center"" height=""10""></td></tr>"
If UCase(Trim(GetRecordField("TestGroupID"))) = "B13" Then '17th March 2022
  str = str & "<tr><td align=""center"" style=""font-family:sans-serif;font-weight:bold;font-size:14pt"">MEDICAL LABORATORY REPORT</td></tr>"
Else
  str = str & "<tr><td align=""center"" style=""font-family:sans-serif;font-weight:bold;font-size:14pt"">MEDICAL RADIOLOGY REPORT</td></tr>"
End If
str = str & "<tr><td align=""center""><hr color=""#999999"" size=""1""></td></tr>"
str = str & "<tr><td align=""center"">"
str = str & "<table border=""0"" width=""650"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial"">"
str = str & "<tr>"
str = str & "<td name=""tdLabelInpLabReceiptID"" id=""tdLabelInpLabReceiptID"" style=""font-weight: bold"">LAB No.</td>"
str = str & "<td width=""10""></td>"
str = str & "<td style=""font-size:11pt"" name=""tdInputInpLabReceiptID"" id=""tdInputInpLabReceiptID""><b>" & GetRecordField("LabReQUESTID") & "</b></td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdLabelInpLabReceiptName"" id=""tdLabelInpLabReceiptName"" style=""font-weight: bold"">Patient Name</td>"
str = str & "<td width=""10""></td>"

If VisitNoExempt(GetRecordField("visitationid")) Or VisitNoExempt(Trim(GetRecordField("Patientid"))) Then
str = str & "<td style=""font-size:11pt"" name=""tdInputInpLabReceiptName"" id=""tdInputInpLabReceiptName""><b>" & GetRequestName(GetRecordField("LabRequestID")) & "</b></td>"
Else
str = str & "<td style=""font-size:11pt"" name=""tdInputInpLabReceiptName"" id=""tdInputInpLabReceiptName""><b>" & GetRecordField("pATIENTName") & "</b></td>"
End If

str = str & "<td width=""10""></td>"
str = str & "</tr>"

str = str & "<tr>"
str = str & "<td name=""tdLabelInpAge"" id=""tdLabelInpAge"" style=""font-weight: bold"">Age</td>"
str = str & "<td width=""10""></td>"

If VisitNoExempt(GetRecordField("visitationid")) Or VisitNoExempt(Trim(GetRecordField("Patientid"))) Then
str = str & "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & GetRequestAge(GetRecordField("LabRequestID")) & "</td>"
Else
str = str & "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & GetPatientAge(GetRecordField("VisitationID")) & "</td>"
End If

str = str & "<td width=""10""></td>"
str = str & "<td name=""tdLabelInpGenderID"" id=""tdLabelInpGenderID"" style=""font-weight: bold"">Gender</td>"
str = str & "<td width=""10""></td>"

If VisitNoExempt(GetRecordField("visitationid")) Or VisitNoExempt(Trim(GetRecordField("Patientid"))) Then
str = str & "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & GetRequestGender(GetRecordField("LabRequestID")) & "</td>"
Else
str = str & "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & GetRecordField("GenderName") & "</td>"
End If
str = str & "<td width=""10""></td>"
str = str & "</tr>"

str = str & "<tr>"
str = str & "<td name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"" style=""font-weight: bold"">Receipt Type</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpReceiptTypeID"" id=""tdInputInpReceiptTypeID""></td>" ' & GetRecordField("ReceiptTypeName") & "</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdLabelInpInsuranceSchemeID"" id=""tdLabelInpInsuranceSchemeID"" style=""font-weight: bold"">Organization</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpInsuranceSchemeID"" id=""tdInputInpInsuranceSchemeID""></td>" ' & GetRecordField("InsuranceSchemeName") & "</td>"
str = str & "<td width=""10""></td>"
str = str & "</tr>"

str = str & "<tr>"
str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Episode No.</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (GetRecordField("VisitationID")) & "</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"" style=""font-weight: bold"">Patient Tel:</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & (GetRecordField("ContactNo")) & "</td>"
str = str & "<td width=""10""></td>"
str = str & "</tr>"

str = str & "<tr>"
str = str & "<td name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"" style=""font-weight: bold"">Manual Path. No</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpReceiptTypeID"" id=""tdInputInpReceiptTypeID"">" & GetRecordField("ReceiptInfo2") & "</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdLabelInpInsuranceSchemeID"" id=""tdLabelInpInsuranceSchemeID"" style=""font-weight: bold""></td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpInsuranceSchemeID"" id=""tdInputInpInsuranceSchemeID""></td>"
str = str & "<td width=""10""></td>"
str = str & "</tr>"
str = str & "</table>"
str = str & "</td></tr>"
'Doctor Info
str = str & "<tr><td align=""center""><hr color=""#999999"" size=""1""></td></tr>"
str = str & "<tr><td align=""center"">"
str = str & "<table border=""0"" width=""650"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial"">"
If UCase(Trim(GetRecordField("TestGroupID"))) = "B13" Then '6th May 2024 ' Frank
str = str & "<tr>"
str = str & "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Requested By</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">" & GetRecordField("DoctorName") & "</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Sample Collection Date</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & FormatDateDetail(GetRecordField("RequestDate")) & "</td>"
str = str & "<td width=""10""></td></tr>"

str = str & "<tr>"
str = str & "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Requested From</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID""></td>" '& GetRecordField("HospitalName") & "</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Received Date</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">&nbsp;&nbsp;&nbsp;&nbsp;</td>"
str = str & "<td width=""10""></td></tr>"

str = str & "<tr>"
str = str & "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Diagnosis</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">&nbsp;" & cliDia & "</td>"
str = str & "<td width=""10""></td>"
End If
str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Report Date</td>"
str = str & "<td width=""10""></td>"

If IsDate(lstDate) Then
  str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & FormatDateDetail(lstDate) & "</td>"
Else
  str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">&nbsp;</td>"
End If
str = str & "<td width=""10""></td></tr>"

str = str & "<tr style=""display:none"">"
str = str & "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Sample Collected At</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">SSNIT:" & GetRecordField("BranchName") & "</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">&nbsp;</td>"
str = str & "<td width=""10""></td>"
str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">&nbsp;</td>"
str = str & "<td width=""10""></td></tr>"
str = str & "</table>"
str = str & "</td></tr>"
'Requested
str = str & "<tr><td align=""center""><hr color=""#999999"" size=""1""></td></tr>"
str = str & "<tr><td align=""left"">"
str = str & "<table border=""0"" width=""650"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial"">"
str = str & "<tr><td><b>REQUESTED:</b>&nbsp;&nbsp;</td><td>" & reqTest & "</td></tr></table>"
str = str & "</td></tr>"
response.write str
End Sub

'GetRequestedTest
Function GetRequestedTest(rec)
Dim ot, rst, sql, cnt, nm, dt, lbTch, dia
ot = ""
sql = "select investigation.labtestid,investigation.labtechid,investigation.requestdate1,investigation.clinicaldiagnosis,labtest.labtestname "
sql = sql & " from investigation,labtest"
sql = sql & " where investigation.labtestid=labtest.labtestid and investigation.labrequestid='" & rec & "'"
sql = sql & " order by investigation.labtestid"
Set rst = CreateObject("ADODB.Recordset")
With rst
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
  .MoveFirst
  cnt = 0
  Do While Not .EOF
    cnt = cnt + 1
    If cnt > 1 Then
      ot = ot & ", "
    End If
    nm = .fields("labtestname")
    dt = .fields("RequestDate1")
    lbTch = .fields("labtechid")
    dia = .fields("clinicaldiagnosis")
    ot = ot & Replace(nm, " ", "&nbsp;")
    If dt > lstDate Then
      lstDate = dt
      lstLbTech = lbTch
    End If
    If Not IsNull(dia) Then
      If Len(Trim(dia)) > 0 Then
        cliDia = "" 'dia
      End If
    End If
    .MoveNext
  Loop
End If
.Close
End With

'///////////Investigation2 /////////////
sql = "select investigation2.labtestid,investigation2.labtechid,investigation2.requestdate1,investigation2.clinicaldiagnosis,labtest.labtestname "
sql = sql & " from investigation2,labtest"
sql = sql & " where investigation2.labtestid=labtest.labtestid and investigation2.labrequestid='" & rec & "'"
sql = sql & " order by investigation2.labtestid"
Set rst = CreateObject("ADODB.Recordset")
With rst
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
  .MoveFirst
  'cnt = 0 continue from above
  Do While Not .EOF
    cnt = cnt + 1
    If cnt > 1 Then
      ot = ot & ", "
    End If
    nm = .fields("labtestname")
    dt = .fields("RequestDate1")
    lbTch = .fields("labtechid")
    dia = .fields("clinicaldiagnosis")
    ot = ot & Replace(nm, " ", "&nbsp;")
    If dt > lstDate Then
      lstDate = dt
      lstLbTech = lbTch
    End If
    If Not IsNull(dia) Then
      If Len(Trim(dia)) > 0 Then
        cliDia = "" 'dia
      End If
    End If
    .MoveNext
  Loop
End If
.Close
End With
GetRequestedTest = ot
End Function

Function VisitNoExempt(vst)
Dim arr, ul, num, lst, ot
ot = False
lst = "0054444||0079669||0185870||0208211||0158043||0099346||0186748||0207940||E01||P1||E02||P2||V02||V03||V04||V05||E02"
arr = Split(lst, "||")
ul = UBound(arr)
For num = 0 To ul
If UCase(Trim(arr(num))) = UCase(Trim(vst)) Then
ot = True
Exit For
End If
Next
VisitNoExempt = ot
End Function


'GetPatientAge
Function GetPatientAge(vst)
Dim ot, rst, sql
sql = "select patientage from visitation where visitationid='" & vst & "'"
Set rst = server.CreateObject("ADODB.Recordset")
With rst
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
ot = .fields("patientage")
End If
.Close
GetPatientAge = CInt(ot)
End With
End Function

'GetRequestAge
Function GetRequestAge(vst)
Dim ot, rst, sql
sql = "select requestage from labrequest where labrequestid='" & vst & "'"
Set rst = server.CreateObject("ADODB.Recordset")
With rst
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
ot = .fields("requestage")
End If
.Close
GetRequestAge = CInt(ot)
End With
End Function

'GetRequestName
Function GetRequestName(vst)
Dim ot, rst, sql
sql = "select requestname from labrequest where labrequestid='" & vst & "'"
Set rst = server.CreateObject("ADODB.Recordset")
With rst
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
ot = .fields("requestname")
End If
.Close
GetRequestName = ot
End With
End Function

'GetRequestGender
Function GetRequestGender(vst)
Dim ot, rst, sql
sql = "select requestgenderid from labrequest where labrequestid='" & vst & "'"
Set rst = server.CreateObject("ADODB.Recordset")
With rst
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
ot = GetComboName("RequestGender", .fields("RequestGenderid"))
End If
.Close
GetRequestGender = ot
End With
End Function

Sub DisplayFoot()
Dim str
str = ""
str = str & "<tr>"
str = str & " <td valign=""bottom"" align=""center"" height=""10"">"
str = str & "<table height=""10"" id=""tblFooter"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
str = str & "<tr><td width=""110""></td>"
str = str & "<td colspan=""6"" bgcolor=""#FFFFFF"" height=""10"" style=""font-size: 8pt"" align=""right"">"
str = str & "Hospital Laboratory.Copyright@2013<br>" 'Software by : Spagad Technologies : 020-9426112<br> "
str = str & "</td>"
str = str & "</tr>"
str = str & "</table>"
str = str & "</td>"
str = str & "</tr>"
response.write str
End Sub
'DisplayLabResults
Sub DisplayLabResults()
Dim rst, sql, rec, lnPerPg, i, j, num, num2, fnd, rst2
Dim lnLbt, lbt, tct, str, cmp, sTy, out, lth, colm, colCnt, tdAttr
Dim arrLbt(27, 5), arrPg(27), numLbt, numPg, pos, pdCnt, cPdPos
Dim blkVl, blkCmp, blkCnt, lbDoc
lnPerPg = 27
numLbt = 0
numPg = 0
lth = 50
colCnt = 6
'Init array
For i = 1 To lnPerPg
For j = 1 To 5
arrLbt(i, j) = ""
arrPg(i) = 0
Next
Next
Set rst = server.CreateObject("ADODB.Recordset")
Set rst2 = server.CreateObject("ADODB.Recordset")
rec = Trim(Request.queryString("LABREQUESTID"))
lbDoc = Trim(Request.queryString("LabByDoctorID"))
lbt = Trim(GetRecordField("LabTestID"))
If Len(rec) > 0 Then
sql = "SELECT testcategoryid,labtestid from investigation"
sql = sql & " where LABREQUESTid='" & rec & "' and labtestid='" & lbt & "' and requeststatusid='RRD002' "
sql = sql & " group by testcategoryid,labtestid order by testcategoryid,labtestid"
With rst
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
.MoveFirst
reqTest = GetRequestedTest(rec)
Do While Not .EOF
lbt = .fields("labtestid")
numLbt = numLbt + 1
arrLbt(numLbt, 1) = .fields("testcategoryid")
arrLbt(numLbt, 2) = .fields("labtestid")
arrLbt(numLbt, 3) = "0"

sql = "SELECT count(labtestid) as cnt from labresults"
  sql = sql & " where labrequestid='" & rec & "' and labtestid='" & lbt & "'"
  With rst2
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  .MoveFirst
  If IsNumeric(.fields("cnt")) Then
  arrLbt(numLbt, 3) = CStr(.fields("cnt"))
  End If
  End If
  .Close
  End With

.MoveNext
Loop
End If
.Close
End With
'Place Lab Test on pages
For num = 1 To numLbt
fnd = False
lnLbt = CInt(arrLbt(num, 3)) + 2
If numPg = 0 Then
numPg = numPg + 1
arrPg(numPg) = lnLbt
arrLbt(num, 4) = CStr(numPg)
fnd = True
Else
For num2 = 1 To numPg
If (lnPerPg >= (arrPg(num2) + lnLbt)) Then
arrPg(num2) = arrPg(num2) + lnLbt
arrLbt(num, 4) = CStr(num2)
fnd = True
Exit For
End If
Next
End If
If Not fnd Then
numPg = numPg + 1
arrPg(numPg) = lnLbt
arrLbt(num, 4) = CStr(numPg)
fnd = True
End If
Next
'Display Pages
response.write "<tr><td width=""100%"" bgcolor=""white"" align=""right"">"
response.write "<p align=""center"">"
response.write "</td></tr>"
For num = 1 To numPg
DisplayHead2
response.write "<tr><td valign=""top"" align=""center"" height=""81"" width=""100%""><table width=""100%"" border=""0"" cellpadding=""0"" cellSpacing=""0"">"
For num2 = 1 To numLbt
If CInt(arrLbt(num2, 4)) = num Then
tct = arrLbt(num2, 1)
lbt = arrLbt(num2, 2)

'Block C/S sensitivity when no isolate
blkCmp = False
blkCnt = 0

sql = "select column1,column2,column3,column4,column5,column6,testcomponentid from labresults "
sql = sql & " where LABREQUESTid='" & rec & "' and labtestid='" & lbt & "' order by comppos"
str = ""
With rst
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
.MoveFirst
sTy = "font-size:10pt; border-collapse:collapse"
str = str & "<tr><td>"
str = str & "<table border=""0"" style=""" & sTy & """ valign=""top"" width=""100%"" cellpadding=""2"" cellSpacing=""0"">"
str = str & "<tr>"
str = str & "<td height=""10"" colspan=""7"" align=""center""></td>"
str = str & "</tr>"
str = str & "<tr>"
sTy = "font-weight:bold; font-size:10pt; border-bottom-style: solid; border-bottom-width: 1px"
str = str & "<td colspan=""7"" align=""left"" style=""" & sTy & """>" & GetComboName("labtest", lbt) & " -> [" & GetComboName("TestCategory", tct) & "]  </td>"
str = str & "</tr>"
cPdPos = 0
Do While Not .EOF
cmp = .fields("testcomponentid")
out = GetComboName("Testcomponent", cmp)
If Len(out) < lth Then
out = Replace(out, " ", " ")
End If
sTy = ""
pos = InStr(1, UCase(out), "PARAMETER")
If (pos > 0) Or (UCase(cmp) = "L0516") Then
  cPdPos = 0
  sTy = "style=""font-weight:bold; font-size:10pt; border-bottom-style: solid; border-bottom-width: 1px"""
Else
  cPdPos = cPdPos + 1
  If cPdPos = 1 Then
    sTy = "style=""padding-top: 3px"""
  End If
End If
'Block C/S sensitivity when no isolate
blkCmp = False
If blkCnt > 0 Then
  blkCnt = blkCnt - 1
  blkCmp = True
End If
If (UCase(cmp) = "L0402") Or (UCase(cmp) = "L0410") Then
  blkVl = Trim(.fields("column1"))
  If UCase(blkVl) = "" Then
    blkCmp = True
    blkCnt = 6
  ElseIf (UCase(blkVl) = "12") Or (UCase(blkVl) = "71") Or (UCase(blkVl) = "72") Or (UCase(blkVl) = "75") Or (UCase(blkVl) = "77") Or (UCase(blkVl) = "80") Then
    blkCnt = 6
  ElseIf (UCase(blkVl) = "T006") Then
    blkCnt = 6
  End If
End If
If Not blkCmp Then
  str = str & "<tr>"
  str = str & "<td valign=""top""" & sTy & ">" & out & "</td>"
  For colm = 1 To colCnt
  tdColSp = "1"
  tdAlign = ""
  tdAttr = " valign=""top"""
  out = GetCompFldVal("TestComponentID", cmp, "Column" & CStr(colm), .fields("column" & CStr(colm)))
  If Len(tdAlign) > 0 Then
  tdAttr = tdAttr & " align=""" & tdAlign & """ "
  Else
  If Len(Trim(out)) >= 3 Then
  If UCase(Left(Trim(out), 3)) = "<B>" Then
  tdAttr = tdAttr & " align=""left"" "
  Else
  tdAttr = tdAttr & " align=""center"" "
  End If
  Else
  tdAttr = tdAttr & " align=""center"" "
  End If
  End If
  If Len(out) < lth Then
  out = Replace(out, " ", " ")
  End If
  If Len(Trim(out)) = 0 Then
  out = " "
  End If
  If CInt(tdColSp) = 1 Then
  str = str & "<td " & sTy & " " & tdAttr & ">" & out & "</td>"
  Else
  str = str & "<td " & sTy & " " & tdAttr & " colspan=""" & tdColSp & """>" & out & "</td>"
  colm = colm + CInt(tdColSp) - 1
  End If
  Next
  str = str & "</tr>"
End If 'blkCmp
.MoveNext
Loop
str = str & "</table>"
str = str & "</td></tr>"
response.write str
End If
.Close
End With
End If
Next


'Fill Blank empty space
For pdCnt = arrPg(num) To lnPerPg
 response.write "<tr>"
 response.write "<td colspan=""7"" align=""center"">&nbsp;</td>"
 response.write "</tr>"
Next
If num = numPg Then
  response.write "<tr>"
  response.write "<td colspan=""7"" align=""right""><br>Electronically Signed By:  " & GetComboName("LabTech", lstLbTech) & "  <br> </td>"
  response.write "</tr>"
End If
response.write "<tr>"
response.write "<td colspan=""7"" align=""center""><br><u>Page " & CStr(num) & " of " & CStr(numPg) & "</u></td>"
response.write "</tr>"
response.write "</table></td></tr>"
If num = numPg Then
DisplayFoot
End If
Next
End If
Set rst = Nothing
Set rst2 = Nothing
End Sub

'DisplayLabResults2
Sub DisplayLabResults2()
Dim rst, sql, rec, lnPerPg, i, j, num, num2, fnd, rst2
Dim lnLbt, lbt, tct, str, cmp, sTy, out, lth, colm, colCnt, tdAttr
Dim arrLbt(27, 5), arrPg(27), numLbt, numPg, pos, pdCnt, cPdPos
Dim blkVl, blkCmp, blkCnt, lbDoc
lnPerPg = 27
numLbt = 0
numPg = 0
lth = 50
colCnt = 6
'Init array
For i = 1 To lnPerPg
For j = 1 To 5
arrLbt(i, j) = ""
arrPg(i) = 0
Next
Next
Set rst = server.CreateObject("ADODB.Recordset")
Set rst2 = server.CreateObject("ADODB.Recordset")

rec = Trim(Request.queryString("LABREQUESTID"))
lbDoc = Trim(Request.queryString("LabByDoctorID"))
lbt = Trim(GetRecordField("LabTestID"))
If Len(rec) > 0 Then
sql = "SELECT testcategoryid,labtestid from investigation2"
sql = sql & " where LABREQUESTid='" & rec & "' and labtestid='" & lbt & "'" ' and requeststatusid='RRD002' "
sql = sql & " group by testcategoryid,labtestid order by testcategoryid,labtestid"
With rst
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
.MoveFirst
reqTest = GetRequestedTest(rec)
Do While Not .EOF
lbt = .fields("labtestid")
numLbt = numLbt + 1
arrLbt(numLbt, 1) = .fields("testcategoryid")
arrLbt(numLbt, 2) = .fields("labtestid")
arrLbt(numLbt, 3) = "0"

sql = "SELECT count(labtestid) as cnt from labresults"
  sql = sql & " where labrequestid='" & rec & "' and labtestid='" & lbt & "'"
  With rst2
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
  .MoveFirst
  If IsNumeric(.fields("cnt")) Then
  arrLbt(numLbt, 3) = CStr(.fields("cnt"))
  End If
  End If
  .Close
  End With

.MoveNext
Loop
End If
.Close
End With
'Place Lab Test on pages
For num = 1 To numLbt
fnd = False
lnLbt = CInt(arrLbt(num, 3)) + 2
If numPg = 0 Then
numPg = numPg + 1
arrPg(numPg) = lnLbt
arrLbt(num, 4) = CStr(numPg)
fnd = True
Else
For num2 = 1 To numPg
If (lnPerPg >= (arrPg(num2) + lnLbt)) Then
arrPg(num2) = arrPg(num2) + lnLbt
arrLbt(num, 4) = CStr(num2)
fnd = True
Exit For
End If
Next
End If
If Not fnd Then
numPg = numPg + 1
arrPg(numPg) = lnLbt
arrLbt(num, 4) = CStr(numPg)
fnd = True
End If
Next
'Display Pages
response.write "<tr><td width=""100%"" bgcolor=""white"" align=""right"">"
response.write "<p align=""center"">"
response.write "</td></tr>"
For num = 1 To numPg
DisplayHead2
response.write "<tr><td valign=""top"" align=""center"" height=""81"" width=""100%""><table width=""100%"" border=""0"" cellpadding=""0"" cellSpacing=""0"">"
For num2 = 1 To numLbt
If CInt(arrLbt(num2, 4)) = num Then
tct = arrLbt(num2, 1)
lbt = arrLbt(num2, 2)

'Block C/S sensitivity when no isolate
blkCmp = False
blkCnt = 0

sql = "select column1,column2,column3,column4,column5,column6,testcomponentid from labresults "
sql = sql & " where LABREQUESTid='" & rec & "' and labtestid='" & lbt & "' order by comppos"
str = ""
With rst
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
.MoveFirst
sTy = "font-size:10pt; border-collapse:collapse"
str = str & "<tr><td>"
str = str & "<table border=""0"" style=""" & sTy & """ valign=""top"" width=""100%"" cellpadding=""2"" cellSpacing=""0"">"
str = str & "<tr>"
str = str & "<td height=""10"" colspan=""7"" align=""center""></td>"
str = str & "</tr>"
str = str & "<tr>"
ot2 = getValidateDetails(lbt, rec)
sTy = "font-weight:bold; font-size:10pt; border-bottom-style: solid; border-bottom-width: 1px"
str = str & "<td colspan=""7"" align=""left"" style=""" & sTy & """>" & GetComboName("labtest", lbt) & " -> [" & GetComboName("TestCategory", tct) & "]  Validated By: " & ot2 & " </td>"
str = str & "</tr>"
cPdPos = 0
Do While Not .EOF
cmp = .fields("testcomponentid")
out = GetComboName("Testcomponent", cmp)
If Len(out) < lth Then
out = Replace(out, " ", " ")
End If
sTy = ""
pos = InStr(1, UCase(out), "PARAMETER")
If (pos > 0) Or (UCase(cmp) = "L0516") Then
  cPdPos = 0
  sTy = "style=""font-weight:bold; font-size:10pt; border-bottom-style: solid; border-bottom-width: 1px"""
Else
  cPdPos = cPdPos + 1
  If cPdPos = 1 Then
    sTy = "style=""padding-top: 3px"""
  End If
End If
'Block C/S sensitivity when no isolate
blkCmp = False
If blkCnt > 0 Then
  blkCnt = blkCnt - 1
  blkCmp = True
End If
If (UCase(cmp) = "L0402") Or (UCase(cmp) = "L0410") Then
  blkVl = Trim(.fields("column1"))
  If UCase(blkVl) = "" Then
    blkCmp = True
    blkCnt = 6
  ElseIf (UCase(blkVl) = "12") Or (UCase(blkVl) = "71") Or (UCase(blkVl) = "72") Or (UCase(blkVl) = "75") Or (UCase(blkVl) = "77") Or (UCase(blkVl) = "80") Then
    blkCnt = 6
  ElseIf (UCase(blkVl) = "T006") Then
    blkCnt = 6
  End If
End If
If Not blkCmp Then
  str = str & "<tr>"
  str = str & "<td valign=""top""" & sTy & ">" & out & "</td>"
  For colm = 1 To colCnt
  tdColSp = "1"
  tdAlign = ""
  tdAttr = " valign=""top"""
  out = GetCompFldVal("TestComponentID", cmp, "Column" & CStr(colm), .fields("column" & CStr(colm)))
  If Len(tdAlign) > 0 Then
  tdAttr = tdAttr & " align=""" & tdAlign & """ "
  Else
  If Len(Trim(out)) >= 3 Then
  If UCase(Left(Trim(out), 3)) = "<B>" Then
  tdAttr = tdAttr & " align=""left"" "
  Else
  tdAttr = tdAttr & " align=""left"" " 'Frank 2024-05-13 changed from center to left
  End If
  Else
  tdAttr = tdAttr & " align=""left"" " 'Frank 2024-05-13 changed from center to left
  End If
  End If
  If Len(out) < lth Then
  out = Replace(out, " ", " ")
  End If
  If Len(Trim(out)) = 0 Then
  out = " "
  End If
  If CInt(tdColSp) = 1 Then
  str = str & "<td " & sTy & " " & tdAttr & ">" & out & "</td>"
  Else
  str = str & "<td " & sTy & " " & tdAttr & " colspan=""" & tdColSp & """>" & out & "</td>"
  colm = colm + CInt(tdColSp) - 1
  End If
  Next
  str = str & "</tr>"
End If 'blkCmp
.MoveNext
Loop
str = str & "</table>"
str = str & "</td></tr>"
response.write str
End If
.Close
End With
End If
Next


'Fill Blank empty space
For pdCnt = arrPg(num) To lnPerPg
 response.write "<tr>"
 response.write "<td colspan=""7"" align=""center"">&nbsp;</td>"
 response.write "</tr>"
Next
If num = numPg Then
  response.write "<tr>"
  response.write "<td colspan=""7"" align=""right""><br>Electronically Signed By:  " & GetComboName("LabTech", lstLbTech) & "  <br> </td>"
  response.write "</tr>"
End If
response.write "<tr>"
response.write "<td colspan=""7"" align=""center""><br><u>Page " & CStr(num) & " of " & CStr(numPg) & "</u></td>"
response.write "</tr>"
response.write "</table></td></tr>"
If num = numPg Then
DisplayFoot
End If
Next
End If
Set rst = Nothing
Set rst2 = Nothing
End Sub

Dim reqTest, lstLbTech, lstDate, cliDia
reqTest = ""
lstLbTech = ""
lstDate = CDate("1 Jan 2000")
cliDia = ""
DisplayLabResults
DisplayLabResults2

Function getValidateDetails(lbtID, labrqtID)
Set rst = server.CreateObject("ADODB.Recordset")
sql = "SELECT labtechid FROM Investigation2 WHERE labtestid = '" & lbtID & "' AND labrequestid = '" & labrqtID & "'"
With rst
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then

ot = GetComboName("Labtech", rst.fields("labtechid"))

End If
End With
getValidateDetails = ot
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
'Empty
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
