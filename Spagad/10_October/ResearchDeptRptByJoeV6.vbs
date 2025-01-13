'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
AddCSS
Dim sql, rst, datePeriod, cnt
'datePeriod = Split(Request.QueryString("PrintFilter"), "||")

Sub displaySRS()
  cnt = 0
  Set rst = CreateObject("ADODB.RecordSet")

  sql = "SELECT DISTINCT EMRRequestID, PatientID FROM EMRRequestItems "
  'sql = sql & " WHERE EMRDate1 BETWEEN '" & datePeriod(0) & "' AND '" & datePeriod(1) & "' "
  sql = sql & " WHERE EMRDate1 BETWEEN '2023-01-01 00:00:00.000' AND '2024-10-02 00:00:00.000' "
  sql = sql & " AND EMRDataID = 'RES018' "
  With rst
      .open qryPro.FltQry(sql), conn, 3, 4
          If .RecordCount > 0 Then
              response.write "<table id='myTbl'>"
              response.write "    <thead>"
              response.write "        <tr>"
              response.write "            <th colspan='8'>" & GetComboName("EMRData", "RES018") & " Scores Report Between January 2020 and September 2024</th>"
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
                  response.write "     <tr>"
                  response.write "       <td>" & cnt & "</td>"
                  response.write "       <td>" & PatientID & "</td>"
                  response.write "       <td>" & GetComboName("Patient", PatientID) & "</td>"
                  response.write "       <td>" & getSRS_Score(EMRRequestID) & "</td>"
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
End Sub

Sub displayHOOS()
  cnt = 0
  Set rst = CreateObject("ADODB.RecordSet")

  sql = "SELECT DISTINCT EMRRequestID, PatientID FROM EMRRequestItems "
  sql = sql & " WHERE EMRDate1 BETWEEN '2023-01-01 00:00:00.000' AND '2024-10-02 00:00:00.000' "
  sql = sql & " AND EMRDataID = 'RES006' "
  With rst
      .open qryPro.FltQry(sql), conn, 3, 4
          If .RecordCount > 0 Then
              response.write "<table id='myTbl'>"
              response.write "    <thead>"
              response.write "        <tr>"
              response.write "            <th colspan='8'>" & GetComboName("EMRData", "RES006") & " Scores Report Between January 2020 and September 2024</th>"
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
                  response.write "     <tr>"
                  response.write "       <td>" & cnt & "</td>"
                  response.write "       <td>" & PatientID & "</td>"
                  response.write "       <td>" & GetComboName("Patient", PatientID) & "</td>"
                  response.write "       <td>" & getHOOS_Score(EMRRequestID) & "</td>"
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
End Sub

Sub displayKOOS()
  cnt = 0
  Set rst = CreateObject("ADODB.RecordSet")

  sql = "SELECT DISTINCT EMRRequestID, PatientID FROM EMRRequestItems "
  sql = sql & " WHERE EMRDate1 BETWEEN '2023-01-01 00:00:00.000' AND '2024-10-02 00:00:00.000' "
  sql = sql & " AND EMRDataID = 'RES007' "
  With rst
      .open qryPro.FltQry(sql), conn, 3, 4
          If .RecordCount > 0 Then
              response.write "<table id='myTbl'>"
              response.write "    <thead>"
              response.write "        <tr>"
              response.write "            <th colspan='8'>" & GetComboName("EMRData", "RES007") & " Scores Report Between January 2020 and September 2024</th>"
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
                  response.write "     <tr>"
                  response.write "       <td>" & cnt & "</td>"
                  response.write "       <td>" & PatientID & "</td>"
                  response.write "       <td>" & GetComboName("Patient", PatientID) & "</td>"
                  response.write "       <td>" & getKOOS_Score(EMRRequestID) & "</td>"
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
End Sub

Sub dispOption()
    Dim sql, rst, str
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "SELECT EMRDataID, EMRDataName FROM EMRData WHERE EMRDataID IN ('RES018', 'RES006', 'RES007') ORDER BY EMRDataName ASC"
    
    response.write "<style>"
    response.write ".cta-form {"
    response.write "  padding: 6px;"
    response.write "  font-size: 15px;"
    response.write "  font-family: inherit;"
    response.write "  color: inherit;"
    response.write "  border: none;"
    response.write "  background-color: #f2f2f2;"
    response.write "  border-radius: 9px;"
    response.write "  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);"
    response.write "} "
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

Function getSRS_Score(EMRRequestID)
    Dim sql, rst, userScore, totalScore, getSRS_ScoreDivided
    Set rst = CreateObject("ADODB.RecordSet")

    sql = "select sum(score) as userScore, sum(total) as totalScore from ( "
    sql = sql & " select sum(ABS(varpos-5)+1) as score, count(varpos) as total from EMRResults e, emrvar3b e2 "
    sql = sql & " where cast(e.Column3 as varchar) = e2.EMRVar3BID and EMRRequestID =  '" & EMRRequestID & "' "
    sql = sql & " union all "
    sql = sql & " select sum(ABS(varpos-5)+1) as score, count(varpos) as total from EMRResults e, emrvar3b e2 "
    sql = sql & " where cast(e.Column6 as varchar) = e2.EMRVar3BID and EMRRequestID =  '" & EMRRequestID & "' "
    sql = sql & " ) as results "

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            userScore = .fields("userScore")
            totalScore = .fields("totalScore")
            
            ' Ensure userScore and totalScore are valid numbers
            If IsNull(userScore) Or IsNull(totalScore) Or totalScore = 0 Then
                getSRS_Score = "Invalid score data"
            Else
                getSRS_ScoreDivided = FormatNumber(userScore / totalScore, 2) ' Formatting with 2 decimal places
                getSRS_Score = "" & userScore & " / " & totalScore & " = " & getSRS_ScoreDivided & ""
            End If
        End If
        .Close
    End With
End Function

Function getHOOS_Score(EMRRequestID)
  Dim sql, rst, userScore, totalScore, getHOOS_ScoreDivided
  Set rst = CreateObject("ADODB.RecordSet")

  sql = "select sum(score) as userScore, sum(total) as totalScore from ( "
  sql = sql & " select sum(ABS(varpos-5)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 "
  sql = sql & " where cast(e.Column2 as varchar) = e2.EMRVar3BID and EMRRequestID =  '" & EMRRequestID & "' "
  sql = sql & " union all "
  sql = sql & " select sum(ABS(varpos-5)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 "
  sql = sql & " where cast(e.Column5 as varchar) = e2.EMRVar3BID and EMRRequestID =  '" & EMRRequestID & "' "
  sql = sql & " ) as results "

  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      ' Handle null values
      userScore = .fields("userScore")
      totalScore = .fields("totalScore")

      If IsNull(userScore) Then userScore = 0
      If IsNull(totalScore) Then totalScore = 0

      ' Check for division by zero
      If totalScore = 0 Then
        getHOOS_Score = "0"
      Else
        ' Multiply totalScore by 4 and calculate the division
        getHOOS_ScoreDivided = FormatNumber((userScore / (totalScore * 4)) * 100)
        getHOOS_Score = "" & getHOOS_ScoreDivided & ""
      End If
    End If
    .Close
  End With
End Function

Function getKOOS_Score(EMRRequestID)
  Dim sql, rst, userScore, totalScore, getKOOS_ScoreDivided
  Set rst = CreateObject("ADODB.RecordSet")

  sql = "select sum(score) as userScore, sum(total) as totalScore from ( "
  sql = sql & " select sum(ABS(varpos-5)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 "
  sql = sql & " where cast(e.Column2 as varchar) = e2.EMRVar3BID and EMRRequestID =  '" & EMRRequestID & "' "
  sql = sql & " union all "
  sql = sql & " select sum(ABS(varpos-5)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 "
  sql = sql & " where cast(e.Column4 as varchar) = e2.EMRVar3BID and EMRRequestID =  '" & EMRRequestID & "' "
  sql = sql & " union all "
  sql = sql & " select sum(ABS(varpos-5)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 "
  sql = sql & " where cast(e.Column6 as varchar) = e2.EMRVar3BID and EMRRequestID =  '" & EMRRequestID & "' "
  sql = sql & " ) as results "

  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      
      userScore = .fields("userScore")
      totalScore = .fields("totalScore")

      If IsNull(userScore) Then userScore = 0
      If IsNull(totalScore) Then totalScore = 0

      If totalScore = 0 Then
        getKOOS_Score = "0"
      Else
        getKOOS_ScoreDivided = FormatNumber((userScore / (totalScore * 4)) * 100)
        getKOOS_Score = "" & getKOOS_ScoreDivided & ""
      End If
    End If
    .Close
  End With
End Function

dispOption

response.write "<div id='srsReport' style='display:none;'>"
displaySRS
response.write "</div>"

response.write "<div id='hoosReport' style='display:none;'>"
displayHOOS
response.write "</div>"

response.write "<div id='koosReport' style='display: none;'>"
displayKOOS
response.write "</div>"

Sub AddCSS()
  response.write "<style>"
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
response.write vbCrLf & "   const srsReport = document.getElementById('srsReport');"
response.write vbCrLf & "   const hoosReport = document.getElementById('hoosReport');"
response.write vbCrLf & "   const koosReport = document.getElementById('koosReport');"

response.write vbCrLf & "   researchForm.addEventListener('change', function() {"
response.write vbCrLf & "       const selectedForm = researchForm.value;"
response.write vbCrLf & "       if (selectedForm === 'RES018') {"
response.write vbCrLf & "           srsReport.style.display = 'block';"
response.write vbCrLf & "           hoosReport.style.display = 'none';"
response.write vbCrLf & "           koosReport.style.display = 'none';"
response.write vbCrLf & "       } else if (selectedForm === 'RES006') {"
response.write vbCrLf & "           srsReport.style.display = 'none';"
response.write vbCrLf & "           hoosReport.style.display = 'block';"
response.write vbCrLf & "           koosReport.style.display = 'none';"
response.write vbCrLf & "       } else if (selectedForm === 'RES007') {"
response.write vbCrLf & "           srsReport.style.display = 'none';"
response.write vbCrLf & "           hoosReport.style.display = 'none';"
response.write vbCrLf & "           koosReport.style.display = 'block';"
response.write vbCrLf & "       }   else {"
response.write vbCrLf & "           srsReport.style.display = 'none';"
response.write vbCrLf & "           hoosReport.style.display = 'none';"
response.write vbCrLf & "           koosReport.style.display = 'none';"
response.write vbCrLf & "       }"
response.write vbCrLf & "   });"
response.write vbCrLf & "</script>"

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>



