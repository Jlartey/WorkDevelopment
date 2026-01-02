<%
' MalariaRptDetails.asp
' Detailed view for Malaria Report

Dim StartDate, endDate, dt, arr, labtestid, genderFilter, ageFilter, sqlCase, testcomp, caseExpr
Dim rst, sql, ageWhere, PosCnt, NegCnt, TotalRecCnt

Set rst = CreateObject("ADODB.Recordset")

dt = Request.QueryString("DatePeriod")
labtestid = Request.QueryString("LabtestID")
genderFilter = Request.QueryString("gender")
ageFilter = Request.QueryString("agerange")

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

' Define the common CASE expression for AgeRange
caseExpr = "CASE " & _
           "WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4 THEN '00-04' " & _
           "WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9 THEN '05-09' " & _
           "WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14 THEN '10-14' " & _
           "WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19 THEN '15-19' " & _
           "WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24 THEN '20-24' " & _
           "WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29 THEN '25-29' " & _
           "WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34 THEN '30-34' " & _
           "WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39 THEN '35-39' " & _
           "WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44 THEN '40-44' " & _
           "WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49 THEN '45-49' " & _
           "WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54 THEN '50-54' " & _
           "WHEN DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59 THEN '55-59' " & _
           "ELSE '60+' " & _
           "END"

' Common subquery part
Dim subQuery
subQuery = "( " & _
           "SELECT DISTINCT " & _
           "i.LabRequestID, " & _
           "i.patientID, " & _
           "CAST(lr.Column1 AS NVARCHAR(MAX)) AS Column1 " & _
           "FROM Investigation i " & _
           "JOIN LabResults lr ON i.LabRequestID = lr.LabRequestID " & _
           "WHERE i.LabTestID = '" & labtestid & "' " & _
           "AND i.RequestStatusID = 'RRD002' " & _
           "AND lr.LabTestID = '" & labtestid & "' " & _
           "AND lr.testcomponentid = '" & testcomp & "' " & _
           "AND i.requestdate BETWEEN '" & arr(0) & "' AND '" & arr(1) & "' " & _
           "UNION " & _
           "SELECT DISTINCT " & _
           "i.LabRequestID, " & _
           "i.patientID, " & _
           "CAST(lr.Column1 AS NVARCHAR(MAX)) AS Column1 " & _
           "FROM Investigation2 i " & _
           "JOIN LabResults lr ON i.LabRequestID = lr.LabRequestID " & _
           "WHERE i.LabTestID = '" & labtestid & "' " & _
           "AND i.RequestStatusID = 'RRD002' " & _
           "AND lr.LabTestID = '" & labtestid & "' " & _
           "AND lr.testcomponentid = '" & testcomp & "' " & _
           "AND i.requestdate BETWEEN '" & arr(0) & "' AND '" & arr(1) & "' " & _
           ") AS Combined "

' Determine ageWhere based on ageFilter
Select Case ageFilter
  Case "00-04"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 0 AND 4"
  Case "05-09"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 5 AND 9"
  Case "10-14"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 10 AND 14"
  Case "15-19"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 15 AND 19"
  Case "20-24"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 20 AND 24"
  Case "25-29"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 25 AND 29"
  Case "30-34"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 30 AND 34"
  Case "35-39"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 35 AND 39"
  Case "40-44"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 40 AND 44"
  Case "45-49"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 45 AND 49"
  Case "50-54"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 50 AND 54"
  Case "55-59"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) BETWEEN 55 AND 59"
  Case "60+"
    ageWhere = "DATEDIFF(YEAR, p.birthdate, GETDATE()) >= 60"
  Case Else
    Response.Write "Invalid age range filter."
    Response.End
End Select

sql = "SELECT " & _
      "p.PatientName, " & _
      caseExpr & " AS AgeRange, " & _
      "COUNT(*) AS TotalRecords, " & _
      sqlCase & " " & _
      "FROM " & subQuery & " " & _
      "JOIN Patient p ON Combined.patientID = p.patientID " & _
      "WHERE p.GenderID = '" & genderFilter & "' " & _
      "AND " & ageWhere & " " & _
      "GROUP BY p.PatientName, " & caseExpr & " " & _
      "ORDER BY p.PatientName ASC"

addCSS

PosCnt = 0
NegCnt = 0
TotalRecCnt = 0

With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    Response.Write "<p><a href='MalariaRpt.asp?DatePeriod=" & dt & "&LabtestID=" & labtestid & "'>Back to Summary</a></p>"
    Response.Write "<table class='anaesthesia'>"
    Response.Write "    <thead>"
    Response.Write "    <tr class='anaesthesia'>"
    Response.Write "        <th colspan='5'>Detailed " & GetComboName("Labtest", labtestid) & " CASES for " & GetComboName("Gender", genderFilter) & ", Age Range " & ageFilter & " Between " & FormatDate(StartDate) & " and " & FormatDate(endDate) & "</th>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tHead'>"
    Response.Write "        <th>Patient Name</th>"
    Response.Write "        <th>Age Range</th>"
    Response.Write "        <th>Total Records</th>"
    Response.Write "        <th>Positive</th>"
    Response.Write "        <th>Negative</th>"
    Response.Write "    </tr>"
    Response.Write "    </thead><tbody>"
    
    .MoveFirst
    Do While Not .EOF
      PosCnt = PosCnt + .fields("Positive").Value
      NegCnt = NegCnt + .fields("Negative").Value
      TotalRecCnt = TotalRecCnt + .fields("TotalRecords").Value
      Response.Write "  <tr class='queryData'>"
      Response.Write "      <td>" & .fields("PatientName").Value & "</td>"
      Response.Write "      <td>" & .fields("AgeRange").Value & "</td>"
      Response.Write "      <td>" & .fields("TotalRecords").Value & "</td>"
      Response.Write "      <td>" & .fields("Positive").Value & "</td>"
      Response.Write "      <td>" & .fields("Negative").Value & "</td>"
      Response.Write "  </tr>"
      .MoveNext
    Loop
  Else
    Response.Write "<p>No records found for the selected criteria.</p>"
  End If
  Response.Write "  <tr class='queryData'>"
  Response.Write "      <td colspan='2'><b>TOTAL</b></td>"
  Response.Write "      <td>" & TotalRecCnt & "</td>"
  Response.Write "      <td>" & PosCnt & "</td>"
  Response.Write "      <td>" & NegCnt & "</td>"
  Response.Write "  </tr>"
  Response.Write "</tbody></table>"
  
  .Close
  Set rst = Nothing
End With

Sub addCSS()
  With Response
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

%>