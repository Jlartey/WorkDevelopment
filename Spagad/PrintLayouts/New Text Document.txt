'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Styling
getEMRResult

Sub getEMRResult()
    Dim sql, count
    
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "WITH EMRResultsCTE AS ("
    sql = sql & " SELECT TOP 100 LEFT(CONVERT(VARCHAR(20), column2), 8) AS EMRVar2BID, emrdate, VisitationID"
    sql = sql & " FROM EmrResults"
    sql = sql & " JOIN emrrequest ON emrrequest.emrrequestid = emrresults.emrrequestid"
    sql = sql & " WHERE EMRDataID = 'TH060'"
    sql = sql & " AND EMRComponentID = 'TH06008'"
    sql = sql & " AND EMRdate BETWEEN '2018-01-01' AND '2024-02-28'),"
    sql = sql & " EMRVar2BCTE AS ("
    sql = sql & " SELECT * FROM emrvar2b"
    sql = sql & " WHERE emrvar2AID = 'TH065'),"
    sql = sql & " DiagnosisCTE AS ("
    sql = sql & " SELECT EMRResultsCTE.EMRVar2BID, EMRDate, VisitationID, EMRVar2BName, EMRVar2AID"
    sql = sql & " FROM EMRResultsCTE"
    sql = sql & " JOIN EMRVar2BCTE ON EMRResultsCTE.EMRVar2BID = EMRVar2BCTE.EMRVar2BID),"
    sql = sql & " DiagnosisDiseaseCTE AS ("
    sql = sql & " SELECT DiagnosisCTE.VisitationID, DiseaseName, EMRVar2BName AS [Diagnosis_Status], EMRDate"
    sql = sql & " FROM DiagnosisCTE"
    sql = sql & " JOIN Diagnosis ON DiagnosisCTE.VisitationID = Diagnosis.VisitationID"
    sql = sql & " JOIN Disease ON Disease.DiseaseID = Diagnosis.DiseaseID)"
    sql = sql & " SELECT [Diagnosis_Status], COUNT(*) AS [Total_Diagnosis] FROM DiagnosisDiseaseCTE"
    sql = sql & " GROUP BY [Diagnosis_Status]"
    
    'response.write sql
    
     With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
           
            .MoveFirst
            
            response.write "<table cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr>"
                response.write "<th class='myth'> No.</th>"
                response.write "<th class='myth'> DIAGNOSIS STATUS </th>"
                response.write "<th class='myth'>  TOTAL DIAGNOSIS  </th>"
            response.write "</tr>"
            
            Do While Not .EOF
                count = count + 1
                response.write "<tr>"
                    response.write "<td class='mytd' align= 'center'>" & count & "</td>"
                    response.write "<td class='mytd' align= 'left'>" & .fields("Diagnosis_Status") & "</td>"
                    response.write "<td class='mytd' align= 'center'>" & .fields("Total_Diagnosis") & "</td>"
                    
                response.write "</tr>"
                .MoveNext
            Loop
            response.write "</table>"
        End If
        .Close
    End With
     Set rst = Nothing
End Sub


Sub Styling()
    response.write " <style>"
        response.write " table {"
        response.write "     width: 65vw;"
        response.write "     font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
        response.write "     border-collapse: collapse;"
        response.write " }"
        
        response.write " .myth, .mytd {"
        response.write "     border: 1px solid #ddd;"
        response.write "     padding: 8px;"
        response.write " }"
        
        response.write " .mytd {"
        response.write "     text-align: 1px solid #ddd;"
        response.write "     padding: 8px;"
        response.write " }"
        
        response.write "  tr:nth-child(even) {"
        response.write "    background-color: #f9f9f9;"
        response.write " } "
        
        response.write " .myth {"
        response.write "     background-color: #f2f2f2;"
        response.write "     color: black;"
        response.write "     text-align: center; "
        response.write " }"
        
    response.write " </style>"

End Sub











'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
