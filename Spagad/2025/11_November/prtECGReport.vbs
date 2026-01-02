'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim emrDat, VisitationID


emrDat = Trim(Request.querystring("EMRDataID"))
VisitationID = Trim(Request.querystring("VisitationID"))
If emrDat = "" Then
    emrDat = "IM058"
End If

'Response.Write "START 12<br/>"

response.write PrintForm(VisitationID, emrDat)

Function PrintForm(VisitationID, emrDat)
    Dim str, sql, rst, reqID, sql2, cmpList
    
'    If UCase(uname) <> "M0303" Then
'        Exit Function
'    End If
    
    cmpList = "IM05801||Column2||Date/Time~~IM05801||Column4||Doctor~~IM05802||Column2||Clinical Indication" _
            & "~~IM05803||Column2||Heart Rate~~IM05803||Column4||Rhythm" _
            & "~~IM05803||Column6||P Wave~~IM05804||Column2||PR Interval~~IM05804||Column4||QRS duration" _
            & "~~IM05804||Column6||QRS Morphology~~IM05805||Column2||QT Interval~~IM05805||Column4||QTc (240 - 440ms Males; 350 - 450ms Females)" _
            & "~~IM05805||Column6||ST segment~~IM05806||Column2||Axis~~IM05807||Column2||Others" _
            & "~~IM05808||Column2||Impression/Dx"
            
    sql = PreparePrintOutSQL(emrDat, VisitationID, cmpList)
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    
    'Response.Write sql
    If rst.recordCount > 0 Then
        'DisplayHeader
        Do While Not rst.EOF
            str = GenerateGenericPrintOut(rst)
        Loop
        rst.Close
        Set rst = Nothing
    End If
    PrintForm = str
End Function

Function GetComponentColValue(emrRequestID, componentID, columnName)
    Dim sql, rst
    
    GetComponentColValue = ""
    
    sql = sql & "SELECT TOP 1 * FROM EMRResults "
    sql = sql & "WHERE 1=1"
    sql = sql & "   AND EMRRequestID='" & emrRequestID & "' "
    sql = sql & "   AND EMRComponentID='" & componentID & "'"
    
    Set rst = CreateObject("ADODB.RecordSet")
    
    
    rst.open sql, conn, 3, 4
    
    If rst.recordCount > 0 Then
        rst.MoveFirst
        GetComponentColValue = Replace(rst.fields(columnName), vbCrLf, "<br/>")
        rst.Close
        Set rst = Nothing
    End If
    
End Function

Function GetMultiSelectValue(tbl, value)
    Dim res
    res = Split(value, "||")
    GetMultiSelectValue = "<ul>"
    For Each val In res
        If val <> "" Then
            GetMultiSelectValue = GetMultiSelectValue & "<li>" & GetComboName(tbl, val) & "</li>"
        End If
    Next
    GetMultiSelectValue = GetMultiSelectValue & "</ul>"
End Function


Function GetHeader()
Dim str
str = ""
  str = str & "<table border=""0px, 1px, 0px, 1px""cellspacing=""0"" cellpadding=""0"" width=""100%"">"
  str = str & "<tr><td>"
  str = str & "<img src=""images/logo.jpg"" height=""110"" width=""110"" style=""filter: gray; -webkit-filter: grayscale(1); filter: grayscale(1);"">"
  str = str & "</td>"
  str = str & "<td>"
  
    str = str & "<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">"
    str = str & "<tr>"
    str = str & "<td align=""center"" style=""font-weight:bold"" colspan=""6""></td>"
    str = str & "</tr>"
    str = str & "<tr>"
    str = str & "<td align=""right"" style=""font-weight: bold; font-size: large;"" colspan=""6"">INTERNATIONAL MARITIME HOSPITAL (GH) LTD</td>"
    str = str & "</tr>"
    str = str & "<tr>"
    str = str & "<td align=""right"" style=""colspan=""6"">P.O. Box CO 4297, Community One, Tema, GHANA</td>"
    str = str & "</tr>"
'    str = str & "<tr>"
'    str = str & "<td align=""right"" style=""font-size: 10pt"" colspan=""6"">COMMUNITY 3, TEMA</td>"
'    str = str & "</tr>"
    
    str = str & "<tr>"
    str = str & "<td align=""right"" style=""colspan=""6"">Telephone:&nbsp;+233-0303-220120 -49</td>"
    str = str & "</tr>"
    str = str & "<tr >"
    str = str & "<td align=""right"" style=""colspan=""6""> Email:&nbsp;info@imah.com.gh&nbsp;&nbsp;&nbsp;&nbsp;Website:&nbsp;imah.com.gh</td>"
    str = str & "<tr/>"
'    str = str & "<tr >"
'    str = str & "<td align=""right"" style=""colspan=""6"">&nbsp;&nbsp;&nbsp;&nbsp;Website:&nbsp;imah.com.gh</td>"
'    str = str & "<tr/>"
    
    
   
    str = str & "</table>"
    
  str = str & "</td>"
  str = str & "</tr>"
  
    'str = str & "<tr>"
    'str = str & "<td align=""center"" colspan=""7""><hr color=""#999999"" size=""1""></td>"
    'str = str & "</tr>"
  str = str & "</table>"
  
  GetHeader = str
End Function


Function PreparePrintOutSQL(emrDataID, VisitationID, cmpColList)
    'cmpColList = EMRComponentID||Column1||Alias~~EMRComponentID||Column2||Alias2
    
    Dim sql, cnt, insql
    cnt = 0
    
    sql = sql & " SELECT v.PatientAge, e.EMRRequestID, v.VisitationID, e.PatientID"
    sql = sql & " , v.GenderID, p.ResidencePhone, p.Occupation, p.CountryID, p.ResidenceAddress "
    sql = sql & " ,p.PatientModeID "
    
    For Each cmpPr In Split(cmpColList, "~~")
        cmp = Split(cmpPr, "||")
        If UBound(cmp) < 2 Then
            PreparePrintOutSQL = ""
            Exit Function
        End If
        
        insql = " , a" & cnt & "." & cmp(1) & " AS [" & cmp(2) & "]"
        
        'MsgBox insql
        
        sql = sql & insql
        cnt = cnt + 1
    Next
    
    sql = sql & ""
    sql = sql & " From EMRRequestItems AS e"
    sql = sql & "   INNER JOIN Visitation AS V ON e.VisitationID=v.VisitationID "
    sql = sql & "   INNER JOIN Patient AS p ON p.PatientID=v.PatientID   "
    
    cnt = 0
    For Each cmpPr In Split(cmpColList, "~~")
        cmp = Split(cmpPr, "||")
        If UBound(cmp) < 2 Then
            PreparePrintOutSQL = ""
            Exit Function
        End If
        sql = sql & "   INNER JOIN EMRResults AS a" & cnt & " ON a" & cnt & ".EMRRequestID=e.EMRRequestID "
        sql = sql & "       AND a" & cnt & ".EMRDataID='" & emrDataID & "' "
        sql = sql & "       AND a" & cnt & ".EMRComponentID='" & cmp(0) & "' "
        
        cnt = cnt + 1
    Next
    
    sql = sql & " WHERE 1=1"
    sql = sql & "   AND e.VisitationID='" & VisitationID & "' "
    sql = sql & "   AND e.EMRDataID='" & emrDataID & "'"
    sq = sql & "    ORDER BY e.EMRDate DESC "
    
    PreparePrintOutSQL = sql
    
End Function

Sub WriteFormData(rst)
    response.write GenerateGenericPrintOut(rst)
End Sub


Function GetConsultingDoctor(VisitationID)
  Dim sql, rst
  rst = CreateObject("ADODB.Recordset")

  sql = "SELECT Staff.StaffName as Doctor FROM EMRRequestItems "
  sql = sql & "JOIN SystemUser ON EMRRequestItems.SystemUserID = SystemUser.SystemUserID "
  sql = sql & "JOIN Staff ON SystemUser.StaffID = Staff.StaffID "
  sql = sql & "WHERE EMRDataid = 'Im058' AND visitationid = '" & VisitationID & "' "
  
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      GetConsultingDoctor = .fields("Doctor")
    End If
    .Close
  End With
  Set rst = Nothing
End Function



Function getEMRResult(EMRRequestID, emrDataID, CompID, column)
    Dim sql, rst,emrValue
    Set rst = server.CreateObject("ADODB.Recordset")
    emrValue = ""
    sql = "SELECT * FROM emrresults WHERE emrrequestid ='" & EMRRequestID & "'"
    sql = sql & " AND emrdataid ='" & emrDataID & "' AND emrcomponentid='" & CompID & "'"
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            If Not IsNull(.fields(column)) Then
                emrValue = Trim(.fields(column))
            End If
            If IsNull(.fields(column)) Then
                emrValue = "Null"
            End If
        End If
        .Close
    End With
    getEMRResult = emrValue
    Set rst = Nothing
End Function
 
Function GenerateGenericPrintOut(rst)
    Dim str, exList
    
    exList = Array("PatientID", "GenderID", "VisitationID", "ResidencePhone", "ResidenceAddress", "PatientAge" _
                , "EMRRequestID", "Occupation", "CountryID", "PatientModeID", "Doctor")
     reqID = rst.fields("EMRRequestID")
     
     str = str & GetHeader
     str = str & GetStyles
'    str = str &  "<div style='height: 50mm;'></div>"
     str = str & "<table class='formtable' cellspacing='0'>"
     str = str & "   <tbody>"
     str = str & "       <tr>"
     str = str & "           <td colspan='11' style='text-align: center; background-color:white; color: black;'> " & UCase(GetComboName("EMRData", emrDat)) & " </td>"
     str = str & "       </tr>"
     str = str & "       <tr>"
     str = str & "           <td colspan='11' style='background-color: transparent;'></td>"
     str = str & "       </tr>"
     str = str & "       <tr>"
     str = str & "           <td colspan='11'> PATIENT INFORMATION </td>"
     str = str & "       </tr>"
     str = str & "       <tr>"
     str = str & "           <td colspan='11' style='background-color: transparent;'></td>"
     str = str & "       </tr>"
     str = str & "       <tr>"
     str = str & "           <td colspan='2'> Paitent Name </td>"
     str = str & "           <td colspan='6' style='white-space: nowrap;'>" & GetComboName("Patient", rst.fields("PatientID")) & " </td>"
     str = str & "           <td colspan='1' style='text-align: right;'> Sex </td>"
     str = str & "           <td colspan='2' style='font-weight: unset;'> " & GetComboName("Gender", rst.fields("GenderID")) & " </td>"
     str = str & "       </tr>"
     str = str & "       <tr>"
     str = str & "           <td colspan='2'> Age </td>"
     str = str & "           <td colspan='1' style='font-weight: unset;'>" & GetComboNameFld("Visitation", rst.fields("VisitationID"), "PatientAge") & "</td>"
     str = str & "           <td colspan='8'></td>"
     str = str & "       </tr>"
'     str = str &  "       <tr>"
'     str = str &  "           <td colspan='2'> Nationality </td>"
'     str = str &  "           <td colspan='3'>" & rst.fields("CountryID") & "</td>"
'     str = str &  "           <td colspan='1'></td>"
'     str = str & "           <td colspan='2'> Occupation </td>"
'     str = str & "           <td colspan='3'>" & rst.fields("Occupation") & "</td>"
'     str = str & "       </tr>"
     str = str & "       <tr>"
     str = str & "           <td colspan='2'> Patient Telephone </td>"
     str = str & "           <td colspan='6'>" & rst.fields("ResidencePhone") & "</td>"
     str = str & "           <td colspan='1' style='text-align: right;'> Address </td>"
     str = str & "           <td colspan='2' style='font-weight: unset;'>" & rst.fields("ResidenceAddress") & ", " & GetComboName("PatientMode", rst.fields("PatientModeID")) & "</td>"
     str = str & "       </tr>"
     str = str & "       <tr>"
     str = str & "           <td colspan='11' style='background-color: transparent;'></td>"
     str = str & "       </tr>"
     
     str = str & "       <tr>"
     str = str & "           <td colspan='11'> MEDICAL INFORMATION </td>"
     str = str & "       </tr>"
     str = str & "       <tr>"
     str = str & "           <td colspan='11' style='background-color: transparent;'></td>"
     str = str & "       </tr>"
     
     Dim newRow, colCount
     colCount = 1
     newRow = True
     
     For Each field In rst.fields
         If Not IsExempted(field.name, exList) Then
            str = str & "<tr><td colspan='3'>" & field.name & "</td>"
            str = str & "<td colspan='8'>" & Replace(rst.fields(field.name), vbCrLf, "<br/>") & "</td></tr>"
         End If
'     str = str & "       <tr>"
'     str = str & "           <td colspan='2'> Clinical Indication </td>"
'     str = str & "           <td colspan='3'>" & rst.fields("Clinical Indication") & "</td>"
'     str = str & "       </tr>"
     Next
     
     str = str & "       <tr>"
     str = str & "           <td colspan='11' style='background-color: transparent;'></td>"
     str = str & "       </tr>"
     str = str & "       <tr>"
     str = str & "           <td colspan='11' style='background-color: transparent;'></td>"
     str = str & "       </tr>"
     str = str & "       <tr>"
     str = str & "           <td></td>"
     str = str & "           <td colspan='1' style='text-align: right;'> Signed </td>"
     str = str & "           <td colspan='8'><hr/></td>"
     str = str & "           <td></td>"
     str = str & "       </tr>"
     str = str & "       <tr>"
     str = str & "           <td></td>"
     str = str & "           <td colspan='9' style='text-align: center;'> Doctor </td>"
     str = str & "           <td></td>"
     str = str & "       </tr>"
'        str = str & "       <tr>"
'        str = str & "           <td colspan='11' style='text-align: center;'> " & UCase(GetComboName("EMRData", emrDat)) & " </td>"
'        str = str & "       </tr>"
     str = str & "   </tbody>"
     str = str & " </table>"
    
     rst.MoveNext
      
      
     GenerateGenericPrintOut = str
End Function

Function IsExempted(fieldName, exList)
    IsExempted = False
    For Each itm In exList
        If UCase(itm) = UCase(fieldName) Then
            IsExempted = True
        End If
    Next
End Function

Function GetStyles()
    Dim str
    str = str & "<style>"
    str = str & ""
    str = str & "   table.formtable{"
    str = str & "       /*size: A4;*/"
    str = str & "       -webkit-print-color-adjust: exact;"
    str = str & "       page-break-after: always; "
    '            str = str & "       font-size: 11pt !important;"
    str = str & "   }"
    str = str & "   td{"
    str = str & "       "
    str = str & "   }"
    str = str & "   .fomtable tr td{"
    str = str & "       text-align: justify;"
    str = str & "       height: 8mm; padding-top: 5px; padding-bottom: 5px; padding-left: 5px;"
    str = str & "   }"
    str = str & "   .formtable tr td[colspan='11']{"
    str = str & "       /*background-color: lime;*/"
    'str = str & "       background-color: #7fd409;"
    'str = str & "       background-color: #313f99;"
    str = str & "       background-color: #959595;"
    str = str & "       padding-left: 10mm;"
    str = str & "       padding-top: 2mm;"
    str = str & "       padding-bottom: 2mm;"
    str = str & "       color: white;"
    str = str & "       font-weight: bold;"
    str = str & "   }"
    str = str & "   .formtable tr td[colspan='1']{"
    str = str & "       min-width: 20mm; padding-top: 5px; padding-bottom: 5px; padding-left: 5px;"
    str = str & "       font-weight: bold; vertical-align: top;"
    str = str & "   }"
    str = str & "   .formtable tr td[colspan='3']{"
    str = str & "       min-width: 60mm; padding-top: 5px; padding-bottom: 5px; padding-left: 10px; font-weight: bold;"
    str = str & "   }"
    str = str & "   .formtable tr td[colspan='2']{"
    str = str & "       min-width: 40mm; padding-top: 5px; padding-bottom: 5px; padding-left: 5px; "
    str = str & "       vertical-align: top; font-weight: bold; "
    str = str & "   }"
    str = str & "   .formtable tr td[colspan='8']{"
    str = str & "       min-width: 40mm; padding-top: 5px; padding-bottom: 5px; padding-left: 5px;"
    str = str & "   }"
    str = str & "   .formtable tr td[colspan='7']{"
    str = str & "       min-width: 40mm; padding-top: 5px; padding-bottom: 5px; padding-left: 5px;"
    str = str & "   }"
    str = str & "   hr{ color: #7fd409;} "
    str = str & "</style>"
    GetStyles = str
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
