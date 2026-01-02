'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

addCSS
generateReport

Sub generateReport()

    Dim rst, sql, breastPos, arPeriod, periodStart, periodEnd, labtestID

    arPeriod = getDatePeriodFromDelim(Trim(Request.QueryString("PrintFilter")))
    periodStart = Request.QueryString("PrintFilter")
    periodEnd = Request.QueryString("PrintFilter1")
    labtestID = Request.QueryString("PrintFilter2")
    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT  Visitation.VisitationID, Visitation.PatientAge AS Age"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column2) AS Birads"
    sql = sql & " , Visitation.VisitDate"
    sql = sql & " , CONVERT(VARCHAR(255), LabResults.column1) AS Breast"
    sql = sql & " FROM Visitation JOIN LabRequest ON LabRequest.VisitationID=Visitation.VisitationID "
    sql = sql & " JOIN LabResults ON LabResults.LabRequestID=LabRequest.LabRequestID "
    sql = sql & " JOIN TestVar3B ON TestVar3B.TestVar3BID=CONVERT(VARCHAR(255), LabResults.column2)"
    sql = sql & " WHERE Visitation.VisitDate  BETWEEN '" & periodStart & "' AND '" & periodEnd & "' AND "
    sql = sql & " LabResults.labtestid = '" & labtestID & "' AND LabResults.CompPos IN ('06','07','08')  "
    sql = sql & " AND CONVERT(VARCHAR(255), labresults.Column2) LIKE '%T0080%'"
    sql = sql & " AND CONVERT(VARCHAR(255), LabResults.column1) <> ''"
    sql = sql & " ORDER BY Age"

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        response.write "<table class='anaesthesia'> "
        response.write "    <thead> "
        response.write "        <tr class='anaesthesia'> "
        response.write "            <td colspan='5'> "
        response.write "                <b>Detailed Report for " & GetComboName("Labtest", labtestID) & " Between " & periodStart & " and " & periodEnd & "</b> "
        response.write "            </td></tr> "
        response.write "        <tr class='h_names'> "
        response.write "            <th>VisitID</th> "
        response.write "            <th>Age</th> "
        response.write "            <th>Birad</th> "
        response.write "            <th>Breast Position</th> "
        response.write "            <th>Visit Date</th> "
        response.write "        </tr> "
        response.write "    </thead><tbody>"

        If .recordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                response.write "     <tr class = 'queryData'> "
                response.write "         <td>" & .fields("VisitationID") & "</td> "
                response.write "         <td>" & .fields("Age") & "</td> "
                response.write "         <td>" & GetComboName("TestVar3B", .fields("Birads")) & "</td> "
                response.write "         <td>" & GetComboName("TestVar3B", .fields("Breast")) & "</td> "
                response.write "         <td>" & .fields("VisitDate") & "</td> "
                response.write "     </tr>"
                response.flush
                .MoveNext
            Loop
            response.write "    </tbody></table>"
        End If
        .Close
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
