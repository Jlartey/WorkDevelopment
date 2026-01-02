'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
addCSS
generateReport

Sub generateReport()
    Dim rst, sql, labtestID, arPeriod, periodStart, periodEnd

    arPeriod = getDatePeriodFromDelim(Trim(Request.QueryString("PrintFilter")))
    periodStart = FormatDate(arPeriod(0))
    periodEnd = FormatDate(arPeriod(1))
    Set rst = CreateObject("ADODB.RecordSet")

    sql = "SELECT LabTestID "
    sql = sql & " FROM LabTest "
    sql = sql & " WHERE LabTestID LIKE '%pink%' OR labtestID IN ('PO_MG01', 'PO_US020', 'US020', 'PO_US_MG01', 'MG001')"

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .recordCount > 0 Then
            response.write "<table class = 'anaesthesia' > "
            response.write "    <thead> "
            response.write "    <tr class = 'anaesthesia'>"
            response.write "        <th colspan = '3'>Generated BREAST MAMMOGRAM REPORT Between " & periodStart & " and " & periodEnd & "</th>"
            response.write "    </tr>"
            response.write "    </thead><tbody> "
            .MoveFirst

            Do While Not .EOF
                labtestID = .fields("LabTestID")

                getTotalTest labtestID, periodStart, periodEnd
                getAgeGroup labtestID, periodStart, periodEnd

                .MoveNext
            Loop
                response.write "</tbody></table>"
        End If
        .Close
    End With
    Set rst = Nothing
End Sub

Function getTotalTest(labtestID, periodStart, periodEnd)
    Dim sql, rst, cnt, testID, hrf
    Set rst = server.CreateObject("ADODB.Recordset")
    cnt = 0

    hrf = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PinkOctoberDetail&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd & "&printfilter2=" & labtestID
    
    sql = "WITH biradsData AS ("
    sql = sql & " SELECT CONCAT(FLOOR(visitation.PatientAge / 10) * 10, ' - ', FLOOR(visitation.PatientAge / 10) * 10 + 9) AS Age"
    sql = sql & " , Visitation.PatientAge, Visitation.VisitationName"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column2) AS birads, LabResults.CompPos"
    sql = sql & " ,Visitation.VisitDate, TestVar3B.TestVar3BName"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column1) AS breast"
    sql = sql & " FROM Visitation JOIN LabRequest ON LabRequest.VisitationID=Visitation.VisitationID "
    sql = sql & " JOIN LabResults ON LabResults.LabRequestID=LabRequest.LabRequestID "
    sql = sql & " JOIN TestVar3B ON TestVar3B.TestVar3BID=CONVERT(VARCHAR(255), LabResults.column2)"
    sql = sql & " WHERE Visitation.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' AND"
    sql = sql & " LabResults.labtestid = '" & labtestID & "' AND "
    sql = sql & " LabResults.CompPos IN ('06','07','08')  "
    sql = sql & " AND CONVERT(VARCHAR(255), labresults.Column2) LIKE '%T0080%'"
    sql = sql & " GROUP BY FLOOR(visitation.PatientAge / 10) * 10, Visitation.PatientAge"
    sql = sql & " , Visitation.VisitationName"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column2)"
    sql = sql & " ,LabResults.CompPos, Visitation.VisitDate"
    sql = sql & " , TestVar3B.TestVar3BName"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column1)"
    sql = sql & " ) "
    sql = sql & " SELECT COUNT(visitationName) AS PatientCount"
    sql = sql & " FROM biradsData"
    sql = sql & " WHERE breast <> ''"

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .recordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                response.write "  <tr class = 'queryDataFoot'> "
                response.write "      <td><b>" & GetComboName("Labtest", labtestID) & "</b></td> "
                response.write "      <td><b>" & .fields("PatientCount") & "</b></td> "
                response.write "      <td><a href='" & hrf & "' target='_blank'>View More</a></td> "
                response.write "  </tr> "
            .MoveNext
            Loop
        End If
    End With
    Set rst = Nothing
    
End Function

Sub getAgeGroup(labtestID, periodStart, periodEnd)
    Dim sql, rst, cnt
    Set rst = server.CreateObject("ADODB.Recordset")
    cnt = 0

    sql = "WITH biradsData AS ("
    sql = sql & " SELECT CONCAT(FLOOR(visitation.PatientAge / 10) * 10, ' - ', FLOOR(visitation.PatientAge / 10) * 10 + 9) AS Age"
    sql = sql & " , Visitation.PatientAge, Visitation.VisitationName"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column2) AS birads, LabResults.CompPos"
    sql = sql & " ,Visitation.VisitDate, TestVar3B.TestVar3BName"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column1) AS breast"
    sql = sql & " FROM Visitation JOIN LabRequest ON LabRequest.VisitationID=Visitation.VisitationID "
    sql = sql & " JOIN LabResults ON LabResults.LabRequestID=LabRequest.LabRequestID "
    sql = sql & " JOIN TestVar3B ON TestVar3B.TestVar3BID=CONVERT(VARCHAR(255), LabResults.column2)"
    sql = sql & " WHERE Visitation.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' AND"
    sql = sql & " LabResults.labtestid = '" & labtestID & "' AND "
    sql = sql & " LabResults.CompPos IN ('06','07','08')  "
    sql = sql & " AND CONVERT(VARCHAR(255), labresults.Column2) LIKE '%T0080%'"
    sql = sql & " GROUP BY FLOOR(visitation.PatientAge / 10) * 10, Visitation.PatientAge"
    sql = sql & " , Visitation.VisitationName"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column2)"
    sql = sql & " ,LabResults.CompPos, Visitation.VisitDate"
    sql = sql & " , TestVar3B.TestVar3BName"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column1)"
    sql = sql & " ) "
    sql = sql & "SELECT age, COUNT(visitationname) AS PatientCount"
    sql = sql & " FROM biradsData "
    sql = sql & " WHERE breast <> ''"
    sql = sql & " GROUP BY age"
    sql = sql & " ORDER BY age"

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .recordCount > 0 Then
            response.write "    <tr class = 'queryDataFoot'><td colspan = '3'><b>Age Grouping for " & GetComboName("Labtest", labtestID) & "</b></td></tr>"
            response.write "    <tr class = 'tHead'> "
            response.write "        <td><b>Age Group</b></td> "
            response.write "        <td colspan = '2'><b>Patient Count</b></td> "
            response.write "    </tr> "
            .MoveFirst
            Do While Not .EOF
                cnt = cnt + .fields("PatientCount")
                response.write "  <tr class = 'queryData'> "
                response.write "      <td>" & .fields("Age") & "</td> "
                response.write "      <td colspan = '2'>" & .fields("PatientCount") & "</td> "
                response.write "  </tr> "
            .MoveNext
            Loop
            response.write "  <tr class = 'queryData'> "
            response.write "      <td><b>Sub Total</b></td> "
            response.write "      <td colspan = '2'><b>" & cnt & "</b></td> "
            response.write "  </tr> "
        End If
        testvar3BID labtestID, periodStart, periodEnd
    End With
    Set rst = Nothing
    
End Sub

Sub testvar3BID(labtestID, periodStart, periodEnd)
    Dim sql, rst, cnt, testID, breastPos
    Set rst = server.CreateObject("ADODB.Recordset")
    cnt = 0

    
    sql = "SELECT * FROM TestVar3B WHERE testvar3BID LIKE '%T0120%'"

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .recordCount > 0 Then
            .MoveFirst
            response.write "  <tr class = 'queryData'> "
            response.write "      <td colspan = '3'><b>Assessment</b></td> "
            response.write "  </tr> "
            response.write "    <tr class = 'tHead'> "
            response.write "        <td><b>Age Group</b></td> "
            response.write "        <td><b>Birad count</b></td> "
            response.write "        <td><b>Patient count</b></td> "
            response.write "    </tr> "
            Do While Not .EOF
                breastPos = .fields("TestVar3BID")
                If breastPos = "T012001" Then
                    response.write "  <tr class = 'queryData'> "
                    response.write "      <td colspan = '3'><b>Right Breast</b></td> "
                    response.write "  </tr> "
                End If
                If breastPos = "T012002" Then
                    response.write "  <tr class = 'queryData'> "
                    response.write "      <td colspan = '3'><b>Left Breast</b></td> "
                    response.write "  </tr> "
                End If
                If breastPos = "T012003" Then
                    response.write "  <tr class = 'queryData'> "
                    response.write "      <td colspan = '3'><b>Both Breast</b></td> "
                    response.write "  </tr> "
                End If
                BiradCount labtestID, periodStart, periodEnd, breastPos
            .MoveNext
            Loop
        End If
    End With
    Set rst = Nothing
    
End Sub

Sub BiradCount(labtestID, periodStart, periodEnd, breastPos)
    Dim sql, rst, cnt, testID
    Set rst = server.CreateObject("ADODB.Recordset")
    cnt = 0
    
    sql = "WITH BiradsData AS ("
    sql = sql & " SELECT CONCAT(FLOOR(visitation.PatientAge / 10) * 10, ' - ', FLOOR(visitation.PatientAge / 10) * 10 + 9) AS Age"
    sql = sql & " , Visitation.PatientAge, Visitation.VisitationName"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column2) AS birads, LabResults.CompPos"
    sql = sql & " ,Visitation.VisitDate, TestVar3B.TestVar3BName"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column1) AS breast"
    sql = sql & " FROM Visitation JOIN LabRequest ON LabRequest.VisitationID=Visitation.VisitationID "
    sql = sql & " JOIN LabResults ON LabResults.LabRequestID=LabRequest.LabRequestID "
    sql = sql & " JOIN TestVar3B ON TestVar3B.TestVar3BID=CONVERT(VARCHAR(255), LabResults.column2)"
    sql = sql & " WHERE Visitation.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' AND"
    sql = sql & " LabResults.labtestid = '" & labtestID & "' AND "
    sql = sql & " LabResults.CompPos IN ('06','07','08')  "
    sql = sql & " AND CONVERT(VARCHAR(255), labresults.Column2) LIKE '%T0080%'"
    sql = sql & " GROUP BY FLOOR(visitation.PatientAge / 10) * 10, Visitation.PatientAge"
    sql = sql & " , Visitation.VisitationName"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column2)"
    sql = sql & " ,LabResults.CompPos, Visitation.VisitDate"
    sql = sql & " , TestVar3B.TestVar3BName"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column1)"
    sql = sql & " ) "
    sql = sql & " SELECT age, COUNT(visitationname) AS PatientCount, Birads "
    sql = sql & " FROM BiradsData "
    If breastPos = "T012001" Then
        sql = sql & " WHERE breast IN ('T012001', '<b>Right Breast</b>') "
    End If
    If breastPos = "T012002" Then
        sql = sql & " WHERE breast IN ('T012002', '<b>Left Breast</b>') "
    End If
    If breastPos = "T012003" Then
        sql = sql & " WHERE breast IN ('T012003', '<b>Both Breast</b>') "
    End If
    sql = sql & " AND breast <> '' "
    sql = sql & " GROUP BY age, Birads"
    sql = sql & " ORDER BY age"

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .recordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                response.write "  <tr class = 'tHead'> "
                response.write "      <td><b>" & .fields("Age") & "</b></td> "
                response.write "      <td>" & GetComboName("TestVar3B", .fields("Birads")) & "</td> "
                response.write "      <td>" & .fields("PatientCount") & "</td> "
                response.write "  </tr> "
            .MoveNext
            Loop
        End If
    End With
    Set rst = Nothing
    
End Sub

Function getDatePeriodFromDelim(strDelimPeriod)
        
    Dim arPeriod, periodStart, periodEnd

    Dim arOut(1)

    arPeriod = Split(strDelimPeriod, "||")

    If UBound(arPeriod) >= 0 Then
        periodStart = arPeriod(0)
    End If
    If UBound(arPeriod) >= 1 Then
        periodEnd = arPeriod(1)
    End If

    periodStart = makeDatePeriod(Trim(periodStart), periodEnd, "0:00:00")
    periodEnd = makeDatePeriod(Trim(periodEnd), periodStart, "23:59:59")

    arOut(0) = periodStart
    arOut(1) = periodEnd

    getDatePeriodFromDelim = arOut

End Function

Function makeDatePeriod(strDateStart, defaultDate, strTime)

    If IsDate(strDateStart) Then
        makeDatePeriod = FormatDate(strDateStart) & " " & Trim(strTime)
    Else

        If IsDate(defaultDate) Then
            makeDatePeriod = FormatDate(defaultDate) & " " & Trim(strTime)
        Else
            makeDatePeriod = FormatDate(Now()) & " " & Trim(strTime)
        End If
    End If

End Function

Sub addCSS()
  With response
    .write " <style> "
    .write "    .anaesthesia, .anaesthesia th, .anaesthesia td{ "
    .write "        border: 1px solid silver; "
    .write "        border-collapse: collapse; "
    .write "        padding: 5px; "
    .write "    } "
    .write "    .anaesthesia{ "
    .write "        width: 80vw; "
    .write "        margin: 0 auto; "
    .write "        font-family: sans-serif; "
    .write "        font-size: 15px; "
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
    .write "    .queryData td{ "
    .write "        font-size: 12; "
    .write "    }  "
    .write "    .queryDataFoot td{ "
    .write "        font-size: 12; "
    .write "        background-color: blanchedalmond; "
    .write "    }  "
    .write "    .anaesthesia th{ "
    .write "        background-color: blanchedalmond; "
    .write "        text-align: center; "
    .write "        font-weight: bold;"
    .write "        font-size: 16px;"
    .write "        color:#000;"
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
