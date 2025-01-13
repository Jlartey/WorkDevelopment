Styling
DiagStatus
Sub DiagStatus()
    Dim rst, sql

    Set rst = CreateObject("ADODB.RecordSet")
    sql = "WITH EMRResultsCTE "
    sql = sql & "AS( "
    sql = sql & "select left (convert(varchar(20),column2),8) EMRVar2BID , emrdate, VisitationID  from EmrResults "
    sql = sql & "join emrrequest on emrrequest.emrrequestid = emrresults.emrrequestid "
    sql = sql & "where EMRDataID = 'TH060' "
    sql = sql & "and EMRComponentID = 'TH06008' "
    sql = sql & "and EMRdate BETWEEN '2018-01-01' and '2018-02-28' "
    sql = sql & "), "
    sql = sql & "EMRVar2BCTE "
    sql = sql & "AS "
    sql = sql & "( "
    sql = sql & "select * from emrvar2b "
    sql = sql & "where emrvar2AID = 'TH065' "
    sql = sql & "), "
    sql = sql & "DiagnosisCTE "
    sql = sql & "AS "
    sql = sql & "( "
    sql = sql & "SELECT EMRResultsCTE.EMRVar2BID,EMRDate, VisitationID, EMRVar2BName,EMRVar2AID FROM EMRResultsCTE "
    sql = sql & "JOIN EMRVar2BCTE on EMRResultsCTE.EMRVar2BID = EMRVar2BCTE.EMRVar2BID "
    sql = sql & "), "
    sql = sql & "DiagnosisDiseaseCTE "
    sql = sql & "AS "
    sql = sql & "( "
    sql = sql & "Select DiagnosisCTE.VisitationID, DiseaseName, EMRVar2BName [Diagnosis_Status] ,EMRDate "
    sql = sql & "from DiagnosisCTE "
    sql = sql & "JOIN Diagnosis on DiagnosisCTE.VisitationID = Diagnosis.VisitationID "
    sql = sql & "JOIN Disease on Disease.DiseaseID = Diagnosis.DiseaseID "
    sql = sql & ") "
    sql = sql & "SELECT [Diagnosis_Status], COUNT(*) [Total_Diagnosis] from DiagnosisDiseaseCTE "
    sql = sql & "GROUP BY [Diagnosis_Status]"

    
    response.write "<h2> Diagnosis </h2>"
    
    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
      
            .MoveFirst
            Dim cnt
            cnt = 0
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr>"
                response.write "<th class='myth'> No. </th>"
                response.write "<th class='myth'> Diagnosis Status </th>"
                response.write "<th class='myth'> Total Diagnosis </th>"
                response.write "</tr>"
            Do While Not .EOF
                cnt = cnt + 1
                response.write "<tr class='clickableRow' onclick='showDetails(" & cnt & ")'>"
                response.write "<td class='mytd1'>" & cnt & "</td>"
                response.write "<td class='mytd1'>" & .fields("Diagnosis_Status") & "</td>"
                response.write "<td class='mytd2'>" & .fields("Total_Diagnosis") & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
            response.write "</table><br><br>"
            
            response.write "<div id='ProvDiag' class='hide diag' >"
                DiffDiag
            response.write "</div>"

            response.write "<div id='FinDiag' class='hide diag'>"
                FinDiag
            response.write "</div>"

            response.write "<div id='DiffDiag' class='hide diag'>"
                ProvDiag
            response.write "</div>"
            
            response.write "<script>"
            response.write "function showDetails(rowNumber) {"
            response.write "    var provDiagDiv = document.getElementById('ProvDiag');"
            response.write "    var finDiagDiv = document.getElementById('FinDiag');"
            response.write "    var diffDiagDiv = document.getElementById('DiffDiag');"
            response.write "    provDiagDiv.style.display = 'none';"
            response.write "    finDiagDiv.style.display = 'none';"
            response.write "    diffDiagDiv.style.display = 'none';"

            response.write "    if (rowNumber === 1) {"
            response.write "        provDiagDiv.style.display = 'block';"
            response.write "    } else if (rowNumber === 2) {"
            response.write "        finDiagDiv.style.display = 'block';"
            response.write "    } else if (rowNumber === 3) {"
            response.write "        diffDiagDiv.style.display = 'block';"
            response.write "    }"
            response.write "}"
            response.write "</script>"
        End If
        .Close
    End With
    
End Sub

Sub DiffDiag()
    Dim rst, sql
    
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "WITH EMRResultsCTE "
    sql = sql & "AS("
    sql = sql & "select left (convert(varchar(20), column2), 8) EMRVar2BID, emrdate, VisitationID from EmrResults "
    sql = sql & "join emrrequest on emrrequest.emrrequestid = emrresults.emrrequestid "
    sql = sql & "where EMRDataID = 'TH060' "
    sql = sql & "and EMRComponentID = 'TH06008' "
    sql = sql & "and EMRdate BETWEEN '2018-01-01' and '2018-02-28' "
    sql = sql & "), "
    sql = sql & "EMRVar2BCTE "
    sql = sql & "AS ("
    sql = sql & "select * from emrvar2b "
    sql = sql & "where emrvar2AID = 'TH065' "
    sql = sql & "), "
    sql = sql & "DiagnosisCTE "
    sql = sql & "AS ("
    sql = sql & "SELECT EMRResultsCTE.EMRVar2BID, EMRDate, VisitationID, EMRVar2BName, EMRVar2AID FROM EMRResultsCTE "
    sql = sql & "JOIN EMRVar2BCTE on EMRResultsCTE.EMRVar2BID = EMRVar2BCTE.EMRVar2BID "
    sql = sql & "), "
    sql = sql & "DiagnosisDiseaseCTE "
    sql = sql & "AS ("
    sql = sql & "Select DiagnosisCTE.VisitationID, DiseaseName, EMRVar2BName AS [Diagnosis_Status], EMRDate, EMRVar2BID "
    sql = sql & "from DiagnosisCTE "
    sql = sql & "JOIN Diagnosis on DiagnosisCTE.VisitationID = Diagnosis.VisitationID "
    sql = sql & "JOIN Disease on Disease.DiseaseID = Diagnosis.DiseaseID "
    sql = sql & ") "
    sql = sql & "SELECT VisitationID, DiagnosisDiseaseCTE.DiseaseName, [Diagnosis_Status], CONVERT(VARCHAR(30), DiagnosisDiseaseCTE.EMRDate, 103) [Diagnosis_Date] "
    sql = sql & "FROM DiagnosisDiseaseCTE "
    sql = sql & "WHERE DiagnosisDiseaseCTE.EMRVar2BID = 'TH065001' "
    
    response.write "<h2> Differential Diagnosis </h2>"
    
    With rst
        .open sql, conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr>"
                response.write "<th class='myth'> No. </th>"
                response.write "<th class='myth'> Visitation ID </th>"
                response.write "<th class='myth'> Disease Name </th>"
                response.write "<th class='myth'> Diagnosis Status </th>"
                response.write "<th class='myth'> Diagnosis Date </th>"
                response.write "</tr>"
            Do While Not .EOF
                cnt = cnt + 1
                response.write "<tr>"
                response.write "<td class='mytd1'>" & cnt & "</td>"
                response.write "<td class='mytd1'>" & .fields("VisitationID") & "</td>"
                response.write "<td class='mytd3'>" & .fields("DiseaseName") & "</td>"
                response.write "<td class='mytd3'>" & .fields("Diagnosis_Status") & "</td>"
                response.write "<td class='mytd2'>" & .fields("Diagnosis_Date") & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
            response.write "</table>"
        End If
    End With

End Sub

Sub ProvDiag()
    Dim rst, sql, cnt
    cnt = 0
    
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "WITH EMRResultsCTE "
    sql = sql & "AS( "
    sql = sql & "select left (convert(varchar(20), column2), 8) EMRVar2BID, emrdate, VisitationID from EmrResults "
    sql = sql & "join emrrequest on emrrequest.emrrequestid = emrresults.emrrequestid "
    sql = sql & "where EMRDataID = 'TH060' "
    sql = sql & "and EMRComponentID = 'TH06008' "
    sql = sql & "and EMRdate BETWEEN '2018-01-01' and '2018-02-28' "
    sql = sql & "), "
    sql = sql & "EMRVar2BCTE "
    sql = sql & "AS ("
    sql = sql & "select * from emrvar2b "
    sql = sql & "where emrvar2AID = 'TH065' "
    sql = sql & "), "
    sql = sql & "DiagnosisCTE "
    sql = sql & "AS ("
    sql = sql & "SELECT EMRResultsCTE.EMRVar2BID, EMRDate, VisitationID, EMRVar2BName, EMRVar2AID FROM EMRResultsCTE "
    sql = sql & "JOIN EMRVar2BCTE on EMRResultsCTE.EMRVar2BID = EMRVar2BCTE.EMRVar2BID "
    sql = sql & "), "
    sql = sql & "DiagnosisDiseaseCTE "
    sql = sql & "AS ("
    sql = sql & "Select DiagnosisCTE.VisitationID, DiseaseName, EMRVar2BName AS [Diagnosis_Status], EMRDate, EMRVar2BID "
    sql = sql & "from DiagnosisCTE "
    sql = sql & "JOIN Diagnosis on DiagnosisCTE.VisitationID = Diagnosis.VisitationID "
    sql = sql & "JOIN Disease on Disease.DiseaseID = Diagnosis.DiseaseID "
    sql = sql & ") "
    sql = sql & "SELECT VisitationID, DiagnosisDiseaseCTE.DiseaseName, [Diagnosis_Status], CONVERT(VARCHAR(30), DiagnosisDiseaseCTE.EMRDate, 103) [Diagnosis_Date] "
    sql = sql & "FROM DiagnosisDiseaseCTE "
    sql = sql & "WHERE DiagnosisDiseaseCTE.EMRVar2BID = 'TH065002' "
    
    response.write "<h2> Provisional Diagnosis </h2>"
    
    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then

            .MoveFirst
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr>"
                response.write "<th class='myth'> No. </th>"
                response.write "<th class='myth'> Visitation ID </th>"
                response.write "<th class='myth'> Disease Name </th>"
                response.write "<th class='myth'> Diagnosis Status </th>"
                response.write "<th class='myth'> Diagnosis Date </th>"
                response.write "</tr>"
            Do While Not .EOF
                cnt = cnt + 1
                response.write "<tr>"
                response.write "<td class='mytd1'>" & cnt & "</td>"
                response.write "<td class='mytd1'>" & .fields("VisitationID") & "</td>"
                response.write "<td class='mytd3'>" & .fields("DiseaseName") & "</td>"
                response.write "<td class='mytd3'>" & .fields("Diagnosis_Status") & "</td>"
                response.write "<td class='mytd2'>" & .fields("Diagnosis_Date") & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
            response.write "</table>"
        End If
    End With

End Sub

Sub FinDiag()
    Dim rst, sql, cnt
    cnt = 0
    
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "WITH EMRResultsCTE "
    sql = sql & "AS( "
    sql = sql & "select left (convert(varchar(20), column2), 8) EMRVar2BID, emrdate, VisitationID from EmrResults "
    sql = sql & "join emrrequest on emrrequest.emrrequestid = emrresults.emrrequestid "
    sql = sql & "where EMRDataID = 'TH060' "
    sql = sql & "and EMRComponentID = 'TH06008' "
    sql = sql & "and EMRdate BETWEEN '2018-01-01' and '2018-02-28' "
    sql = sql & "), "
    sql = sql & "EMRVar2BCTE "
    sql = sql & "AS ("
    sql = sql & "select * from emrvar2b "
    sql = sql & "where emrvar2AID = 'TH065' "
    sql = sql & "), "
    sql = sql & "DiagnosisCTE "
    sql = sql & "AS ("
    sql = sql & "SELECT EMRResultsCTE.EMRVar2BID, EMRDate, VisitationID, EMRVar2BName, EMRVar2AID FROM EMRResultsCTE "
    sql = sql & "JOIN EMRVar2BCTE on EMRResultsCTE.EMRVar2BID = EMRVar2BCTE.EMRVar2BID "
    sql = sql & "), "
    sql = sql & "DiagnosisDiseaseCTE "
    sql = sql & "AS ("
    sql = sql & "Select DiagnosisCTE.VisitationID, DiseaseName, EMRVar2BName AS [Diagnosis_Status], EMRDate, EMRVar2BID "
    sql = sql & "from DiagnosisCTE "
    sql = sql & "JOIN Diagnosis on DiagnosisCTE.VisitationID = Diagnosis.VisitationID "
    sql = sql & "JOIN Disease on Disease.DiseaseID = Diagnosis.DiseaseID "
    sql = sql & ") "
    sql = sql & "SELECT VisitationID, DiagnosisDiseaseCTE.DiseaseName, [Diagnosis_Status], CONVERT(VARCHAR(30), DiagnosisDiseaseCTE.EMRDate, 103) [Diagnosis_Date] "
    sql = sql & "FROM DiagnosisDiseaseCTE "
    sql = sql & "WHERE DiagnosisDiseaseCTE.EMRVar2BID = 'TH065003' "
    
    response.write "<h2> Final Diagnosis </h2>"
    
    With rst
        .open sql, conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr>"
                response.write "<th class='myth'> No. </th>"
                response.write "<th class='myth'> Visitation ID </th>"
                response.write "<th class='myth'> Disease Name </th>"
                response.write "<th class='myth'> Diagnosis Status </th>"
                response.write "<th class='myth'> Diagnosis Date </th>"
                response.write "</tr>"
            Do While Not .EOF
                cnt = cnt + 1
                response.write "<tr>"
                response.write "<td class='mytd1'>" & cnt & "</td>"
                response.write "<td class='mytd1'>" & .fields("VisitationID") & "</td>"
                response.write "<td class='mytd3'>" & .fields("DiseaseName") & "</td>"
                response.write "<td class='mytd3'>" & .fields("Diagnosis_Status") & "</td>"
                response.write "<td class='mytd2'>" & .fields("Diagnosis_Date") & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
            response.write "</table>"
        End If
    End With

End Sub

Sub Styling()
    response.write " <style>"
        response.write " table {"
        response.write "     width: 65vw;"
        response.write "     font-family: 'Poppins', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
        response.write "     border-collapse: collapse;"
        response.write "     margin-top: 10px;"
        response.write " }"
        
        response.write " .myth {"
        response.write "     border: 1px solid #020929;"
        response.write "     padding: 8px;"
        response.write "     text-align: center"
        response.write " }"
        
        response.write " .mytd1 {"
        response.write "     border: 1px solid #020929;"
        response.write "     padding: 8px;"
        response.write "     text-align: left"
        response.write " }"
        
        response.write " .mytd2 {"
        response.write "     border: 1px solid #020929;"
        response.write "     padding: 8px;"
        response.write "     text-align: right"
        response.write " }"
        
        response.write "  tr:nth-child(even) {"
        response.write "    background-color: #edeef0; "
        response.write " } "
        
        
        response.write " .myth {"
        response.write "     background-color: #627df5;"
        response.write "     color: black;"
        response.write " }"
        
        response.write "  .mytd3 {"
        response.write "     border: 1px solid #020929;"
        response.write "     padding: 8px;"
        response.write "     text-align: left"
        response.write "  }"
        
        response.write " .hide {"
        response.write "    display: none;"
        response.write "  }"
        
        response.write " .clickableRow {"
        response.write "    cursor: pointer;"
        response.write "  }"

    response.write " </style>"
End Sub