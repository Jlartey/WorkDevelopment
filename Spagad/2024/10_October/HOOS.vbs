'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

AddCSSS
Dim sql, sql1, rst, rst1, datePeriod, cnt, vst, sponsor, gen, spclst, 
Dim spclstTpe, spclstGrp, vstTpe, specialist, whcls, WorkingMonthID
datePeriod = Split(Request.QueryString("PrintFilter"), "||")

cnt = 0
Set rst = CreateObject("ADODB.RecordSet")

sql = "SELECT DISTINCT EMRRequestID, PatientID FROM EMRRequestItems "
' sql = sql & " WHERE EMRDate1 BETWEEN '01 june 2024 00:00:00' AND '30 june 2024 23:59:59' "
sql = sql & " WHERE EMRDate1 BETWEEN '" & datePeriod(0) & "' AND '" & datePeriod(1) & "' "
sql = sql & " AND EMRDataID = 'RES006' "
With rst
    .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
                    dispOption
                    response.write "<table id='myTbl'>"
                    response.write "    <thead>"
                    response.write "        <tr>"
                    response.write "            <th colspan='8'>" & GetComboName("EMRData", "RES006") & " Scores Report Between " & FormatDate(datePeriod(0)) & " and " & FormatDate(datePeriod(1)) & "</th>"
                    response.write "        </tr>"
                    response.write "        <tr>"
                    response.write "            <th>#</th>"
                    response.write "            <th>Patient ID</th>"
                    response.write "            <th>Patient Name</th>"
                    response.write "            <th>Score</th>"
                    response.write "        </tr>"
                    response.write "    </thead>"
                    response.write "    <tbody>"
            .movefirst
                Do While Not .EOF
                  cnt = cnt + 1
                    PatientID = .fields("PatientID")
                    EMRRequestID = .fields("EMRRequestID")
                    ' PatientName = GetComboName("Patient", .fields("PatientID"))
                    response.write "     <tr>"
                    response.write "       <td>" & cnt & "</td>"
                    response.write "       <td>" & PatientID & "</td>"
                    response.write "       <td>" & GetComboName("Patient", PatientID) & "</td>"
                    response.write "       <td>" & getScore(EMRRequestID) & "</td>"
                    response.write "     </tr>"
                    .MoveNext
                Loop
                response.write "    </tbody>"
                response.write "</table>"
        Else
            response.write "<p style='margin: 50px;'>OOPS ... No records found. Please Try again.</p>"
        End If
    .Close
End With

Sub dispOption()
    Dim sql, rst, str
    Set rst = CreateObject("ADODB.RecordSet")
'    strTblID = strTable & "ID"
'    sql = "SELECT * FROM " & strTable & " "
    sql = "SELECT EMRDataID, EMRDataName FROM EMRData WHERE EMRDataID IN ('RES004', 'RES017', 'RES018', 'RES006') ORDER BY EMRDataName ASC"
    
    response.write "<style>"
    response.write ".cta-form{"
'    response.write "  width: 100%;"
    response.write "  padding: 6px;"
    response.write "  font-size: 15px;"
    response.write "  font-family: inherit;"
    response.write "  color: inherit;"
    response.write "  border: none;"
    response.write "  background-color: #f2f2f2;"
    response.write "  border-radius: 9px;"
    response.write "  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);"
    response.write "}"
    response.write "</style>"
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            response.write " <select name='researchForm' id='researchForm' class='cta-form'> "
            response.write "   <option value=''> Select Form </option> "
            .movefirst
                Do While Not .EOF
                    emrDataID = .fields("EMRDataID")
                    emrDataName = .fields("EMRDataName")
                    response.write "   <option value='" & emrDataID & "'> " & emrDataName & " </option> "
                    .MoveNext
                Loop
            response.write " </select> "
        End If
        .Close
    End With
End Sub

Function getScore(EMRRequestID)
  Dim sql, rst
  Set rst = CreateObject("ADODB.RecordSet")

  sql = "select sum(score) as userScore, sum(total) as totalScore from ( "
  sql = sql & " select sum(ABS(varpos-5)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 "
  sql = sql & " where cast(e.Column2 as varchar) = e2.EMRVar3BID and EMRRequestID =  '" & EMRRequestID & "' "
  sql = sql & " union all "
  sql = sql & " select sum(ABS(varpos-5)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 "
  sql = sql & " where cast(e.Column5 as varchar) = e2.EMRVar3BID and EMRRequestID =  '" & EMRRequestID & "' "
  sql = sql & " )as results "
'  response.write sql

  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        getScoreDivided = FormatNumber(.fields("userScore") / .fields("totalScore"))
        getScore = "" & .fields("userScore") & " / " & .fields("totalScore") & " = " & getScoreDivided & ""
'        getScore = "check your fields lol"
      End If
    .Close
  End With
End Function

Sub AddCSSS()
  response.write "<style>"
  ' response.write "  body {"
  ' response.write "    font-family: Arial, sans-serif;"
  ' response.write "    margin: 20px;"
  ' response.write "    padding: 0;"
  ' response.write "    background-color: #f4f4f9;"
  ' response.write "}  "
  response.write "  #myTbl {"
  response.write "    width: 55vw;"
  response.write "    border-collapse: collapse;"
  response.write "    margin-bottom: 20px;"
  response.write "    margin-top: 50px;"
  response.write "    text-transform: uppercase;"
  response.write "}  "
  response.write "  #myTbl th, #myTbl td {"
  response.write "    padding: 10px;"
  response.write "    border: 1px solid #dddddd;"
  response.write "    text-align: left;"
  response.write "}  "
  response.write "  #myTbl th {"
  response.write "    background-color: #f2f2f2;"
  response.write "}  "
  response.write "  #myTbl tbody tr:nth-child(odd) {"
  response.write "    background-color: #fafafa;"
  response.write "}"
  response.write "  #myTbl tbody tr:hover {"
  response.write "    background-color: #f2f2f2;"
  response.write "}"
  response.write "#SpecialistType, #submit, #WorkingMonth, #WorkingDay{"
  response.write "  padding: 5px;"
  response.write "  font-size: 15px;"
  response.write "  font-family: inherit;"
  response.write "  color: inherit;"
  response.write "  border: none;"
  response.write "  background-color: #f2f2f2;"
  response.write "  border-radius: 9px;"
  response.write "  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);"
'  response.write "  margin-bottom: 10px;"
  response.write "}"
  response.write "  .grid-2-temp {"
  response.write "    display: grid;"
  response.write "    grid-template-columns: 1fr 1fr 1fr;"
  response.write "    min-width: 80vw;"
  response.write "    gap: 10px;"
  response.write "}"
  response.write "  #submit {"
  response.write "    background-color: #b7b7b7;"
  response.write "}"
  response.write "  #submit:hover {"
  response.write "      background-color: #fff;"
  response.write "      color: #555;"
  response.write "      cursor: pointer;"
  response.write "  }"
  response.write "</style>"
End Sub

response.write vbCrLf & "<script>"
response.write vbCrLf & "   const researchForm = document.getElementById('researchForm');"
response.write vbCrLf & "   researchForm.addEventListener('change', updateURL);"
response.write vbCrLf & "   function updateURL(){"
response.write vbCrLf & "       const inpresearchForm = document.getElementById('researchForm');"
response.write vbCrLf & "       let currURL = new URL(window.location.href);"
response.write vbCrLf & "       let params = new URLSearchParams(currURL.search);"
response.write vbCrLf & "       params.set('researchForm', inpresearchForm.value);"
response.write vbCrLf & "       currURL.search = params.toString();"
response.write vbCrLf & "       window.location.href = currURL.toString();"
response.write vbCrLf & "       console.log(currURL.toString());"
response.write vbCrLf & "   };"
response.write vbCrLf & "</script>"






' dont know what this function is doing
Function getSQL(EMRdata, EMRRequestID)
    Dim sql
    
    If UCase(EMRdata) = "RES004" Then
    
    End If
'    If UCase(EMRdata) = "RES017" Then
'
'    End If
    If UCase(EMRdata) = "RES018" Or UCase(EMRdata) = "RES017" Then
        sql = "select sum(score) as userScore, sum(total) as totalScore from ( "
        sql = sql & " select sum(ABS(varpos-5)+1) as score, 5 * count(varpos) as total from EMRResults e, emrvar3b e2 "
        sql = sql & " where cast(e.Column3 as varchar) = e2.EMRVar3BID and EMRRequestID =  '" & EMRRequestID & "' "
        sql = sql & " union  "
        sql = sql & " select sum(ABS(varpos-5)+1) as score, 5 * count(varpos) as total from EMRResults e, emrvar3b e2 "
        sql = sql & " where cast(e.Column6 as varchar) = e2.EMRVar3BID and EMRRequestID =  '" & EMRRequestID & "' "
        sql = sql & " )as results "
    End If
    
    getSQL = sql
    
End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
