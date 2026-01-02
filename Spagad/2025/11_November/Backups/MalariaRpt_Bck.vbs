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
    response.write " <style>"
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
    response.write "            <label for='malariaRpt'>Select Test</label>"
    response.write "            <select name='malariaRpt' id='malariaRpt' class='cta-form' class='myinput'>"
    response.write "                <option value=''>Select Test</option>"
    response.write "                <option value='86750'>Malaria</option>"
    response.write "                <option value='87390'>HIV</option>"
    response.write "            </select>"
    response.write "        </div>"
    response.write "        <div>"
    response.write "            <button type='button' onclick='processPrint()'>Process Print</button>"
    response.write "        </div>"
    response.write "    </div>"
    response.write "</form>"
    response.write "<script>"
    response.write "    function processPrint() {"
    response.write "        const LabtestID = document.getElementById('malariaRpt').value;"
    response.write "        const fromDate = document.getElementById('from').value;"
    response.write "        const toDate = document.getElementById('to').value;"
'    response.write "        const researchForm = document.getElementById('researchForm').value;"
    '    response.write "        if (!fromDate || !toDate || !researchForm) {"
    '    response.write "            alert('Please select the form and date range.');"
    '    response.write "            return;"
    '    response.write "        }"
    
    response.write "        let url = window.location.href.split('?')[0];"
    response.write "        const params = new URLSearchParams(window.location.search);"
    response.write "        params.set('PrintLayoutName', 'MalariaRpt');"
    response.write "        params.set('LabtestID', LabtestID);"
    response.write "        params.set('DatePeriod', fromDate + '||' + toDate);"
    response.write "        window.location.href = url + '?' + params.toString();"
    response.write "    }"
    response.write "</script>"
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

  response.write sql
  cnt = 0
  PosCnt = 0
  NegCnt = 0

  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
        
            response.write "<table class = 'anaesthesia' > "
            response.write "    <thead> "
            response.write "    <tr class = 'anaesthesia'>"
            response.write "        <th colspan = '5'>Generated " & GetComboName("Labtest", labtestid) & " CASES Between " & FormatDate(StartDate) & " and " & FormatDate(endDate) & "</th>"
            response.write "    </tr>"
            response.write "    <tr class = 'tHead'> "
            response.write "        <th>GENDER</th> "
            response.write "        <th>AGE RANGE</th> "
            response.write "        <th>POSITIVE</th> "
            response.write "        <th>NEGATIVE</th> "
            response.write "        <th>VIEW MORE</th> "
            response.write "    </tr> "
            response.write "    </thead><tbody> "

            .MoveFirst
            Do While Not .EOF
                PosCnt = PosCnt + .fields("Positive")
                NegCnt = NegCnt + .fields("Negative")
                response.write "  <tr class = 'queryData'> "
                response.write "      <td>" & GetComboName("Gender", .fields("GenderID")) & "</td> "
                response.write "      <td>" & .fields("AgeRange") & "</td> "
                response.write "      <td>" & .fields("Positive") & "</td> "
                response.write "      <td>" & .fields("Negative") & "</td> "
                response.write "      <td> <a href=" & hrf1 & " target=""_blank"">View More</a> </td> "
                response.write "  </tr> "
                .MoveNext
            Loop
        End If
                response.write "  <tr class = 'queryData'> "
                response.write "      <td colspan='2'><b>TOTAL</b></td> "
                response.write "      <td>" & PosCnt & "</td> "
                response.write "      <td>" & NegCnt & "</td> "
                response.write "      <td> - </td> "
                response.write "  </tr> "
        response.write "</tbody></table>"

    .Close
    Set rst = Nothing
  End With
End Sub

Sub addCSS()
  With response
    .write " <style> "
    .write "    .anaesthesia, .anaesthesia th, .anaesthesia td{ "
    .write "        border: 1px solid silver; "
    .write "        border-collapse: collapse; "
    .write "        padding: 5px; "
    .write "    } "
    .write "    .anaesthesia{ "
    .write "        width: 650px; "
    .write "        margin: 0 auto; "
    .write "        font-family: sans-serif; "
    .write "        font-size: 13px; "
    .write "        box-sizing: border-box; "
    .write "    }"
    .write "    .anaesthesia tr{page-break-inside:avoid; "
    .write "        page-break-after:auto "
    .write "    } "
    .write "    .anaesthesia th, .anaesthesia td { "
    .write "        border: 1px solid silver; "
    .write "        text-align: center; "
    .write "        padding: 5px; "
    .write "        font-size:13px; "
    .write "        margin: 0 auto; "
    .write "    } "
    .write "    .tHead{ "
    .write "        position: sticky; top: 0; "
    .write "    }  "
    .write "    .queryData td{ "
    .write "        font-size: 12; "
    .write "    }  "
    .write "    .anaesthesia th{ "
    .write "        background-color: blanchedalmond; "
    .write "        text-align: center; "
    .write "        font-weight: bold;"
    .write "        font-size: 14px;color:#000;"
    .write "   } "
    .write "    .text-align td:nth-child(2), .text-align th:nth-child(2) { "
    .write "        text-align: left; "
    .write "   } "
    .write " </style> "
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
