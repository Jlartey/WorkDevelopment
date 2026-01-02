'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
addCSS
DatePicker
Dim sql, rst, datePeriod, cnt, periodStart, periodEnd
datePeriod = Split(Request.QueryString("DatePeriod"), "||")
 
form = Request.QueryString("ResearchForm")

If UCase(form) = "SRS" Then
    displaySRS
End If
If UCase(form) = "HOOS" Then
    displayHOOS
End If
If UCase(form) = "KOOS" Then
    displayKOOS
End If

Sub DatePicker()
       
    response.write "<style>"
    response.write "    #dateForm {"
    response.write "        max-width: 600px;"
    response.write "        margin: 20px auto;"
    response.write "        padding: 20px;"
    response.write "        border: 1px solid #ccc;"
    response.write "        border-radius: 8px;"
    response.write "        background-color: #f9f9f9;"
    response.write "        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);"
    response.write "    }"
    response.write "    .container {"
    response.write "        display: flex;"
    response.write "        justify-content: space-between;"
    response.write "        margin-bottom: 15px;"
    response.write "    }"
    response.write "    div {"
    response.write "        flex: 1;"
    response.write "        margin-right: 10px;"
    response.write "    }"
    response.write "    div:last-child {"
    response.write "        margin-right: 0;"
    response.write "    }"
    response.write "    label {"
    response.write "        display: block;"
    response.write "        margin-bottom: 5px;"
    response.write "        font-weight: bold;"
    response.write "    }"
    response.write "    .myinput[type='date'], select {"
    response.write "        width: 100%;"
    response.write "        padding: 8px;"
    response.write "        border: 1px solid #ccc;"
    response.write "        border-radius: 4px;"
    response.write "        box-sizing: border-box;"
    response.write "    }"
    response.write "    button {"
    response.write "        padding: 10px 15px;"
    response.write "        background-color: #28a745;"
    response.write "        color: white;"
    response.write "        border: none;"
    response.write "        border-radius: 4px;"
    response.write "        cursor: pointer;"
    response.write "        transition: background-color 0.3s;"
    response.write "        margin-top: 20px"
    response.write "    }"
    response.write "    button:hover {"
    response.write "        background-color: #218838;"
    response.write "    }"
    response.write "</style>"

    response.write "<form id='dateForm'>"
    response.write "    <div class='container' style='display: flex;'>"
    response.write "        <div>"
    response.write "            <label for='from'>From</label>"
    response.write "            <input type='date' name='from' id='from' class='myinput'>"
    response.write "        </div>"
    response.write "        <div>"
    response.write "            <label for='to'>To</label>"
    response.write "            <input type='date' name='to' id='to' class='myinput'>"
    response.write "        </div>"
    response.write "        <div>"
    response.write "            <label for='researchForm'>Select Form</label>"
    response.write "            <select name='researchForm' id='researchForm' class='cta-form' class='myinput'>"
    response.write "                <option value=''>Select Form</option>"
    response.write "                <option value='SRS'>SRS</option>"
    response.write "                <option value='HOOS'>HOOS</option>"
    response.write "                <option value='KOOS'>KOOS</option>"
    response.write "            </select>"
    response.write "        </div>"
    response.write "        <div>"
    response.write "            <button type='button' onclick='processPrint()'>Process Print</button>"
    response.write "        </div>"
    response.write "    </div>"
    response.write "</form>"
    response.write "<script>"
    response.write "    function processPrint() {"
    response.write "        const fromDate = document.getElementById('from').value;"
    response.write "        const toDate = document.getElementById('to').value;"
    response.write "        const researchForm = document.getElementById('researchForm').value;"
'    response.write "        if (!fromDate || !toDate || !researchForm) {"
'    response.write "            alert('Please select the form and date range.');"
'    response.write "            return;"
'    response.write "        }"

    response.write "        let url = window.location.href.split('?')[0];"
    response.write "        const params = new URLSearchParams(window.location.search);"

    response.write "        params.set('PrintLayoutName', 'ResearchDeptRpt');"
    response.write "        params.set('DatePeriod', fromDate + '||' + toDate);"
    response.write "        params.set('ResearchForm', researchForm);"

    response.write "        window.location.href = url + '?' + params.toString();"
    response.write "    }"
    response.write "</script>"
End Sub

Sub displaySRS()
  Dim url, vst, periodStart, periodEnd
  cnt = 0
  Set rst = CreateObject("ADODB.RecordSet")

  sql = "SELECT DISTINCT EMRRequestID, PatientID,visitationID FROM EMRRequestItems "
  sql = sql & " WHERE EMRDataID = 'RES018' "
  If IsArray(datePeriod) Then
      If Len(Trim(datePeriod(0))) > 0 And Len(Trim(datePeriod(1))) > 0 Then
          sql = sql & " AND EMRDate1 BETWEEN '" & datePeriod(0) & "' AND '" & datePeriod(1) & "' "
      End If
  End If
 
  'response.write sql
  With rst
      .open qryPro.FltQry(sql), conn, 3, 4
          If .RecordCount > 0 Then
              response.write "<table id='myTbl'>"
              response.write "    <thead>"
              response.write "        <tr>"
              response.write "            <th colspan='8' style='text-align: center'>" & GetComboName("EMRData", "RES018") & " Scores Report Between " & FormatDate(datePeriod(0)) & " and " & FormatDate(datePeriod(1)) & "</th>"
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
                vst = .fields("visitationID")
                
                url = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=EMRSelector&CompTableKeyID=EMRComponentID&VisitationID=" & vst & "&EMRDataID=RES018&InvestDays=&ModuleManagerName=Research&PositionForCtxTableName=Visitation&SectionType=EMR&WorkFlowNav=POP"
                
                response.write "     <tr onclick=""window.open('" & url & "', '_blank');"" style='cursor:pointer;'>"
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
  Dim url, vst
  cnt = 0
  Set rst = CreateObject("ADODB.RecordSet")

   sql = "SELECT DISTINCT EMRRequestID, PatientID,visitationID FROM EMRRequestItems "
   sql = sql & " WHERE EMRDataID = 'RES006' "
   If IsArray(datePeriod) Then
      If Len(Trim(datePeriod(0))) > 0 And Len(Trim(datePeriod(1))) > 0 Then
          sql = sql & " AND EMRDate1 BETWEEN '" & datePeriod(0) & "' AND '" & datePeriod(1) & "' "
      End If
   End If
  
  With rst
      .open qryPro.FltQry(sql), conn, 3, 4
          If .RecordCount > 0 Then
              response.write "<table id='myTbl'>"
              response.write "    <thead>"
              response.write "        <tr>"
              response.write "            <th colspan='8' style='text-align: center'>" & GetComboName("EMRData", "RES006") & " Scores Report Between " & FormatDate(datePeriod(0)) & " and " & FormatDate(datePeriod(1)) & "</th>"
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
                  vst = .fields("visitationID")
                  
                  url = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=EMRSelector&CompTableKeyID=EMRComponentID&VisitationID=" & vst & "&EMRDataID=RES006&InvestDays=&ModuleManagerName=Research&PositionForCtxTableName=Visitation&SectionType=EMR&WorkFlowNav=POP"
                  response.write "     <tr onclick=""window.open('" & url & "', '_blank');"" style='cursor:pointer;'>"
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
  Dim vst, url
  cnt = 0
  Set rst = CreateObject("ADODB.RecordSet")

   sql = "SELECT DISTINCT EMRRequestID, PatientID,visitationID FROM EMRRequestItems "
   sql = sql & " WHERE EMRDataID = 'RES007' "
   If IsArray(datePeriod) Then
      If Len(Trim(datePeriod(0))) > 0 And Len(Trim(datePeriod(1))) > 0 Then
          sql = sql & " AND EMRDate1 BETWEEN '" & datePeriod(0) & "' AND '" & datePeriod(1) & "' "
      End If
   End If
  
  With rst
      .open qryPro.FltQry(sql), conn, 3, 4
          If .RecordCount > 0 Then
              response.write "<table id='myTbl'>"
              response.write "    <thead>"
              response.write "        <tr>"
              response.write "            <th colspan='8' style='text-align: center'>" & GetComboName("EMRData", "RES007") & " Scores Report Between " & FormatDate(datePeriod(0)) & " and " & FormatDate(datePeriod(1)) & "</th>"
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
                  vst = .fields("visitationID")
                  
                  url = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=EMRSelector&CompTableKeyID=EMRComponentID&VisitationID=" & vst & "&EMRDataID=RES007&InvestDays=&ModuleManagerName=Research&PositionForCtxTableName=Visitation&SectionType=EMR&WorkFlowNav=POP"
                  response.write "     <tr onclick=""window.open('" & url & "', '_blank');"" style='cursor:pointer;'>"
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

Sub addCSS()
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

Sub ShowReport()
    Dim form
    form = Request.QueryString("ResearchForm")
    response.write form
    
    If form = "SRS" Then
        displaySRS
    ElseIf form = "HOOS" Then
        displayHOOS
    ElseIf form = "KOOS" Then
        displayKOOS
    End If
End Sub



'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
