'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim genderID, ageRange, periodStart, periodEnd
genderID = Request.QueryString("genderID")
ageRange = Request.QueryString("ageRange")
periodStart = Request.QueryString("periodStart")
periodEnd = Request.QueryString("periodEnd")

tableStyles
dispMalariaRptDetails


Sub dispMalariaRptDetails()
    Dim count, sql, rst, ageRange
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT "
    sql = sql & " p.PatientID,"
    sql = sql & " p.PatientName,"
    sql = sql & " CASE"
    sql = sql & "     WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4 THEN '00-04'"
    sql = sql & "     WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9 THEN '05-09'"
    sql = sql & "     WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14 THEN '10-14'"
    sql = sql & "     WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19 THEN '15-19'"
    sql = sql & "     WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24 THEN '20-24'"
    sql = sql & "     WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29 THEN '25-29'"
    sql = sql & "     WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34 THEN '30-34'"
    sql = sql & "     WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39 THEN '35-39'
    sql = sql & "     WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44 THEN '40-44'
    sql = sql & "     WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49 THEN '45-49'
    sql = sql & "     WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54 THEN '50-54'
    sql = sql & "     WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59 THEN '55-59'
    sql = sql & "     Else '60+'
    sql = sql & " END AS AgeRange,
    sql = sql & " COUNT(*) AS TotalRecords,
    sql = sql & " SUM(CASE WHEN Combined.Column1 = '2' THEN 1 ELSE 0 END) AS Positive,
    sql = sql & " SUM(CASE WHEN Combined.Column1 = '1' THEN 1 ELSE 0 END) AS Negative
    sql = sql & " FROM (
    sql = sql & " 
    sql = sql & " SELECT DISTINCT
    sql = sql & "     i.LabRequestID,
    sql = sql & "     i.patientID,
    sql = sql & "     CAST(lr.Column1 AS NVARCHAR(MAX)) AS Column1
    sql = sql & " FROM Investigation i
    sql = sql & " JOIN LabResults lr ON i.LabRequestID = lr.LabRequestID
    sql = sql & " WHERE i.LabTestID = '86750'
    sql = sql & "     AND i.RequestStatusID = 'RRD002'
    sql = sql & "     AND lr.LabTestID = '86750'
    sql = sql & "     AND lr.testcomponentid = 'L0698'
    sql = sql & "     AND i.requestdate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'
    sql = sql & " Union
    sql = sql & " -- Second part: From Investigation2 table
    sql = sql & " SELECT DISTINCT
    sql = sql & "     i.LabRequestID,
    sql = sql & "     i.patientID,
    sql = sql & "     CAST(lr.Column1 AS NVARCHAR(MAX)) AS Column1
    sql = sql & " FROM Investigation2 i
    sql = sql & " JOIN LabResults lr ON i.LabRequestID = lr.LabRequestID
    sql = sql & " WHERE i.LabTestID = '86750'
    sql = sql & "    AND i.RequestStatusID = 'RRD002'
    sql = sql & "    AND lr.LabTestID = '86750'
    sql = sql & "    AND lr.testcomponentid = 'L0698'
    sql = sql & "    AND i.requestdate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'
    sql = sql & " ) AS Combined
    sql = sql & " JOIN Patient p ON Combined.patientID = p.patientID
    sql = sql & " WHERE p.GenderID = '" & genderID & "'
    sql = sql & "     AND CASE
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4 THEN '00-04'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9 THEN '05-09'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14 THEN '10-14'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19 THEN '15-19'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24 THEN '20-24'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29 THEN '25-29'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34 THEN '30-34'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39 THEN '35-39'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44 THEN '40-44'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49 THEN '45-49'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54 THEN '50-54'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59 THEN '55-59'
    sql = sql & "         Else '60+'
    sql = sql & "     END = '" & ageRange & "'
    sql = sql & " Group By
    sql = sql & "     p.PatientID,
    sql = sql & "     p.PatientName,
    sql = sql & "     CASE
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4 THEN '00-04'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9 THEN '05-09'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14 THEN '10-14'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19 THEN '15-19'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24 THEN '20-24'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29 THEN '25-29'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34 THEN '30-34'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39 THEN '35-39'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44 THEN '40-44'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49 THEN '45-49'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54 THEN '50-54'
    sql = sql & "         WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59 THEN '55-59'
    sql = sql & "         Else '60+'
    sql = sql & "     End
    sql = sql & " ORDER BY p.PatientName ASC;
    
    response.write sql

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
        If .RecordCount > 0 Then
            response.write "<h1 style='text-align: center; color: #000;'>TEST RESULTS FOR " & GetComboName("Gender", genderID) & " WITHIN THE AGE RANGE OF " & ageRange& "</h1>"
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Patient ID</th>"
            response.write "<th class='myth'>Name</th>"
            response.write "<th class='myth'>Age Range</th>"
            response.write "<th class='myth'>Total Records</th>"
            response.write "<th class='myth'>Positive</th>"
            response.write "<th class='myth'>Negative</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                visitationID = .fields("VisitationID")
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("PatientID") & "</td>"
                response.write "<td class='mytd'>" & .fields("PatientName") & "</td>"
                response.write "<td class='mytd'>" & .fields("AgeRange") & "</td>"
                response.write "<td class='mytd'>" & .fields("TotalRecords") & "</td>"
                response.write "<td class='mytd'>" & .fields("Positive") & "</td>"
                response.write "<td class='mytd'>" & .fields("Negative") & "</td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop

            response.write "</table>"
            
        Else
            response.write "<h1>No records found</h1>"
        End If
        
        .Close
    End With
    
    Set rst = Nothing
End Sub


Sub tableStyles()
    response.write "<style>"
        response.write ".mytable {"
        response.write "    width: 65vw;"
        response.write "    border-collapse: collapse;"
        response.write "    margin: 20px 0;"
        response.write "    font-size: 16px;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "}"
        response.write ".mytable, .myth, .mytd {"
        response.write "    border: 1px solid #dddddd;"
        response.write "}"
        response.write ".myth, .mytd {"
        response.write "    padding: 12px;"
        response.write "    text-align: left;"
        response.write "}"
        response.write ".myth {"
        response.write "    background-color: #f2f2f2;"
        response.write "    color: #333;"
        response.write "    font-weight: bold;"
        response.write "}"
        response.write ".mytr:nth-child(even) {"
        response.write "    background-color: #f9f9f9;"
        response.write "}"
        response.write ".mytr:hover {"
        response.write "    background-color: #f1f1f1;"
        response.write "}"
        response.write ".myth {"
        response.write "    text-transform: uppercase;"
        response.write "}"
        response.write "h1 {"
        response.write "    font-size: 18px;"
        response.write "    color: #555;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "    margin: 20px 0;"
        response.write "}"
response.write "</style>"

End Sub



'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>



