'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim StartDate, endDate, dt, arr, labtestid
options
addCSS

dt = Request.QueryString("DatePeriod")
labtestid = Request.QueryString("LabtestID")
If Len(dt) > 0 Then
    arr = Split(dt, "||")
    StartDate = arr(0)
    endDate = arr(1)
End If
If Len(labtestid) > 0 And UCase(labtestid) = "86750" Then
    sqlCase = "SUM(CASE WHEN Combined.Column1 = '2' THEN 1 ELSE 0 END) AS Positive,"
    sqlCase = sqlCase & " SUM(CASE WHEN Combined.Column1 = '1' THEN 1 ELSE 0 END) AS Negative"
    testcomp = "L0698"
Else
    sqlCase = "SUM(CASE WHEN Combined.Column1 = 'T001' THEN 1 ELSE 0 END) AS Positive,"
    sqlCase = sqlCase & " SUM(CASE WHEN Combined.Column1 = 'T002' THEN 1 ELSE 0 END) AS Negative"
    testcomp = "865001"
End If

Sub options()
    response.Write " <style>"
    response.Write "    #dateForm {"
    response.Write "        max-width: 600px;"
    response.Write "        margin: 20px auto;"
    response.Write "        padding: 20px;"
    response.Write "        border: 1px solid #ccc;"
    response.Write "        border-radius: 8px;"
    response.Write "        background-color: #f9f9f9;"
    response.Write "        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);"
    response.Write "    }"
    response.Write "    .container {"
    response.Write "        display: flex;"
    response.Write "        justify-content: space-between;"
    response.Write "        margin-bottom: 15px;"
    response.Write "    }"
    response.Write "    div {"
    response.Write "        flex: 1;"
    response.Write "        margin-right: 10px;"
    response.Write "    }"
    response.Write "    div:last-child {"
    response.Write "        margin-right: 0;"
    response.Write "    }"
    response.Write "    label {"
    response.Write "        display: block;"
    response.Write "        margin-bottom: 5px;"
    response.Write "        font-weight: bold;"
    response.Write "    }"
    response.Write "    .myinput[type='date'], select {"
    response.Write "        width: 100%;"
    response.Write "        padding: 8px;"
    response.Write "        border: 1px solid #ccc;"
    response.Write "        border-radius: 4px;"
    response.Write "        box-sizing: border-box;"
    response.Write "    }"
    response.Write "    button {"
    response.Write "        padding: 10px 15px;"
    response.Write "        background-color: #28a745;"
    response.Write "        color: white;"
    response.Write "        border: none;"
    response.Write "        border-radius: 4px;"
    response.Write "        cursor: pointer;"
    response.Write "        transition: background-color 0.3s;"
    response.Write "        margin-top: 20px"
    response.Write "    }"
    response.Write "    button:hover {"
    response.Write "        background-color: #218838;"
    response.Write "    }"
    response.Write "</style>"
    response.Write "<form id='dateForm'>"
    response.Write "    <div class='container' style='display: flex;'>"
    response.Write "        <div>"
    response.Write "            <label for='from'>From</label>"
    response.Write "            <input type='date' name='from' id='from' class='myinput'>"
    response.Write "        </div>"
    response.Write "        <div>"
    response.Write "            <label for='to'>To</label>"
    response.Write "            <input type='date' name='to' id='to' class='myinput'>"
    response.Write "        </div>"
    response.Write "        <div>"
    response.Write "            <label for='malariaRpt'>Select Test</label>"
    response.Write "            <select name='malariaRpt' id='malariaRpt' class='cta-form' class='myinput'>"
    response.Write "                <option value=''>Select Test</option>"
    response.Write "                <option value='86750'>Malaria</option>"
    response.Write "                <option value='87390'>HIV</option>"
    response.Write "            </select>"
    response.Write "        </div>"
    response.Write "        <div>"
    response.Write "            <button type='button' onclick='processPrint()'>Process Print</button>"
    response.Write "        </div>"
    response.Write "    </div>"
    response.Write "</form>"
    response.Write "<script>"
    response.Write "    function processPrint() {"
    response.Write "        const LabtestID = document.getElementById('malariaRpt').value;"
    response.Write "        const fromDate = document.getElementById('from').value;"
    response.Write "        const toDate = document.getElementById('to').value;"
'    response.write "        const researchForm = document.getElementById('researchForm').value;"
    '    response.write "        if (!fromDate || !toDate || !researchForm) {"
    '    response.write "            alert('Please select the form and date range.');"
    '    response.write "            return;"
    '    response.write "        }"
    
    response.Write "        let url = window.location.href.split('?')[0];"
    response.Write "        const params = new URLSearchParams(window.location.search);"
    response.Write "        params.set('PrintLayoutName', 'MalariaRpt');"
    response.Write "        params.set('LabtestID', LabtestID);"
    response.Write "        params.set('DatePeriod', fromDate + '||' + toDate);"
    response.Write "        window.location.href = url + '?' + params.toString();"
    response.Write "    }"
    response.Write "</script>"
End Sub


generateReport

Sub generateReport()
  Dim sql, rst
  Set rst = CreateObject("ADODB.Recordset")
  hrf1 = "#"
  hrf2 = "#"
    ' 86750 malaria
  sql = "SELECT  "
sql = sql & "    p.GenderID, "
sql = sql & "    CASE  "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4 THEN '00-04' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9 THEN '05-09' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14 THEN '10-14' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19 THEN '15-19' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24 THEN '20-24' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29 THEN '25-29' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34 THEN '30-34' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39 THEN '35-39' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44 THEN '40-44' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49 THEN '45-49' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54 THEN '50-54' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59 THEN '55-59' "
sql = sql & "        ELSE '60+' "
sql = sql & "    END AS AgeRange, "
sql = sql & "    COUNT(*) AS TotalRecords, "
sql = sql & "    " & sqlCase & " "
'sql = sql & "    SUM(CASE WHEN Combined.Column1 = '1' THEN 1 ELSE 0 END) AS Negative "
sql = sql & "FROM ( "
sql = sql & "    SELECT DISTINCT  "
sql = sql & "        i.LabRequestID,  "
sql = sql & "        i.patientID,  "
sql = sql & "        CAST(lr.Column1 AS NVARCHAR(MAX)) AS Column1 "
sql = sql & "    FROM Investigation i "
sql = sql & "    JOIN LabResults lr ON i.LabRequestID = lr.LabRequestID "
sql = sql & "    WHERE i.LabTestID = '" & labtestid & "'  "
sql = sql & "      AND i.RequestStatusID = 'RRD002'  "
sql = sql & "      AND lr.LabTestID = '" & labtestid & "' "
sql = sql & "      AND lr.testcomponentid = '" & testcomp & "' "
sql = sql & "      AND i.requestdate BETWEEN '" & arr(0) & "' AND '" & arr(1) & "' "
sql = sql & "     "
sql = sql & "    UNION "
sql = sql & "     "
sql = sql & "    SELECT DISTINCT  "
sql = sql & "        i.LabRequestID,  "
sql = sql & "        i.patientID,  "
sql = sql & "        CAST(lr.Column1 AS NVARCHAR(MAX)) AS Column1 "
sql = sql & "    FROM Investigation2 i "
sql = sql & "    JOIN LabResults lr ON i.LabRequestID = lr.LabRequestID "
sql = sql & "    WHERE i.LabTestID = '" & labtestid & "'  "
sql = sql & "      AND i.RequestStatusID = 'RRD002'  "
sql = sql & "      AND lr.LabTestID = '" & labtestid & "' "
sql = sql & "      AND lr.testcomponentid = '" & testcomp & "' "
sql = sql & "      AND i.requestdate BETWEEN '" & arr(0) & "' AND '" & arr(1) & "' "
sql = sql & ") AS Combined "
sql = sql & "JOIN Patient p ON Combined.patientID = p.patientID "
sql = sql & "GROUP BY  "
sql = sql & "    p.GenderID, "
sql = sql & "    CASE  "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4 THEN '00-04' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9 THEN '05-09' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14 THEN '10-14' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19 THEN '15-19' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24 THEN '20-24' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29 THEN '25-29' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34 THEN '30-34' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39 THEN '35-39' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44 THEN '40-44' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49 THEN '45-49' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54 THEN '50-54' "
sql = sql & "        WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59 THEN '55-59' "
sql = sql & "        ELSE '60+' "
sql = sql & "    END "
sql = sql & "   ORDER BY AgeRange ASC  "


  cnt = 0
  PosCnt = 0
  NegCnt = 0

  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
        
            response.Write "<table class = 'anaesthesia' > "
            response.Write "    <thead> "
            response.Write "    <tr class = 'anaesthesia'>"
            response.Write "        <th colspan = '5'>Generated " & GetComboName("Labtest", labtestid) & " CASES Between " & FormatDate(StartDate) & " and " & FormatDate(endDate) & "</th>"
            response.Write "    </tr>"
            response.Write "    <tr class = 'tHead'> "
            response.Write "        <th>GENDER</th> "
            response.Write "        <th>AGE RANGE</th> "
            response.Write "        <th>POSITIVE</th> "
            response.Write "        <th>NEGATIVE</th> "
            response.Write "        <th>VIEW MORE</th> "
            response.Write "    </tr> "
            response.Write "    </thead><tbody> "

            .MoveFirst
            Do While Not .EOF
                PosCnt = PosCnt + .fields("Positive")
                NegCnt = NegCnt + .fields("Negative")
                response.Write "  <tr class = 'queryData'> "
                response.Write "      <td>" & GetComboName("Gender", .fields("GenderID")) & "</td> "
                response.Write "      <td>" & .fields("AgeRange") & "</td> "
                response.Write "      <td>" & .fields("Positive") & "</td> "
                response.Write "      <td>" & .fields("Negative") & "</td> "
                response.Write "      <td> <a href=" & hrf1 & " target=""_blank"">View More</a> </td> "
                response.Write "  </tr> "
                .MoveNext
            Loop
        End If
                response.Write "  <tr class = 'queryData'> "
                response.Write "      <td colspan='2'><b>TOTAL</b></td> "
                response.Write "      <td>" & PosCnt & "</td> "
                response.Write "      <td>" & NegCnt & "</td> "
                response.Write "      <td> - </td> "
                response.Write "  </tr> "
        response.Write "</tbody></table>"

    .Close
    Set rst = Nothing
  End With
End Sub

Sub addCSS()
  With response
    .Write " <style> "
    .Write "    .anaesthesia, .anaesthesia th, .anaesthesia td{ "
    .Write "        border: 1px solid silver; "
    .Write "        border-collapse: collapse; "
    .Write "        padding: 5px; "
    .Write "    } "
    .Write "    .anaesthesia{ "
    .Write "        width: 650px; "
    .Write "        margin: 0 auto; "
    .Write "        font-family: sans-serif; "
    .Write "        font-size: 13px; "
    .Write "        box-sizing: border-box; "
    .Write "    }"
    .Write "    .anaesthesia tr{page-break-inside:avoid; "
    .Write "        page-break-after:auto "
    .Write "    } "
    .Write "    .anaesthesia th, .anaesthesia td { "
    .Write "        border: 1px solid silver; "
    .Write "        text-align: center; "
    .Write "        padding: 5px; "
    .Write "        font-size:13px; "
    .Write "        margin: 0 auto; "
    .Write "    } "
    .Write "    .tHead{ "
    .Write "        position: sticky; top: 0; "
    .Write "    }  "
    .Write "    .queryData td{ "
    .Write "        font-size: 12; "
    .Write "    }  "
    .Write "    .anaesthesia th{ "
    .Write "        background-color: blanchedalmond; "
    .Write "        text-align: center; "
    .Write "        font-weight: bold;"
    .Write "        font-size: 14px;color:#000;"
    .Write "   } "
    .Write "    .text-align td:nth-child(2), .text-align th:nth-child(2) { "
    .Write "        text-align: left; "
    .Write "   } "
    .Write " </style> "
  End With
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
