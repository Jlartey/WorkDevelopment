'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
dateP = Trim(Request.QueryString("PrintFilter0"))
minF = Trim(Request.QueryString("PrintFilter1"))
specialistid = Trim(Request.QueryString("PrintFilter2"))
arr = Split(dateP, "||")
startDate = arr(0)
endDate = arr(1)
displayStyle
displayReport startDate, endDate, minF, specialistid

Sub displayReport(st, en, mn, specialistid)
    Dim rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    sql = getQuery(st, en, mn, specialistid)
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        response.write "<table class='table-report'><thead><tr style='border:1px solid gray'><th>SPECIALIST NAME</th><th>DISEASE NAME</th>"
        response.write "<th>NEW CASES</th><th>OLD CASES</th></tr></thead><tbody>"
        rst.MoveFirst
        Do While Not rst.EOF
            str = "<tr><td>" & rst.fields("specialistname") & "</td><td>" & rst.fields("diseasename") & "</td>"
            str = str & "<td>" & rst.fields("new_frequency") & "</td><td>" & rst.fields("old_frequency") & "</td></tr>"
            response.write str
            rst.MoveNext
            response.flush
        Loop
    End If
    response.write "</tbody></table>"
    Set rst = Nothing
End Sub

Function getQuery(st, en, mn, specialistid)
    sql = "WITH myTable"
    sql = sql & "    AS"
    sql = sql & "    ("
    sql = sql & "    select  Specialist.SpecialistName,diseasename,"
    sql = sql & "     diagnosis.patientid+'-'+diagnosis.diseaseid+'-'+Specialist.SpecialistName [IDS],"
    sql = sql & "     lag(diagnosis.patientid+'-'+diagnosis.diseaseid+'-'+Specialist.SpecialistName)"
    sql = sql & "     OVER (PARTITION BY diagnosis.patientid+'-'+diagnosis.diseaseid+'-'+Specialist.SpecialistName"
  
    sql = sql & "     ORDER BY  ConsultReviewDate) [lags]"

    sql = sql & "   ,ConsultReviewDate"
    sql = sql & "   from disease join diagnosis"
    sql = sql & "   on disease.DiseaseID=Diagnosis.DiseaseID"
    sql = sql & "   join Visitation on visitation.VisitationID=diagnosis.VisitationID"
    sql = sql & "   join specialist on Visitation.SpecialistID=Specialist.SpecialistID"
    sql = sql & "   where ConsultReviewDate between '" & st & "' and '" & en & "'"
    sql = sql & "   and Specialist.SpecialistName NOT LIKE '%-%'"
    If specialistid <> "" Then
        sql = sql & "   and Specialist.specialistid ='" & specialistid & "'"
    End If
    sql = sql & "   group by diseasename,Specialist.SpecialistName,"
    sql = sql & "   diagnosis.patientid+'-'+diagnosis.diseaseid,"
    sql = sql & "   ConsultReviewDate"
    sql = sql & "   ),"
    sql = sql & "   mytable2 AS "
    sql = sql & "   (select specialistname,diseasename,ConsultReviewDate,"
    sql = sql & "   CASE"
    sql = sql & "   WHEN lags IS NUll THEN 'New Case'"
    sql = sql & "   END [new],"
    sql = sql & "   CASE"
    sql = sql & "   WHEN lags IS NOT NUll THEN 'Old Case'"
    sql = sql & "   END [Old]"
    sql = sql & "   From myTable"
    sql = sql & "   )"
    sql = sql & "   select specialistname,diseasename,"
    sql = sql & "   COUNT(specialistname+diseasename+new) [new_frequency],"
    sql = sql & "   COUNT(specialistname+diseasename+old) [old_frequency]"
    sql = sql & "   From mytable2"
 
    sql = sql & "   GROUP BY specialistname,diseasename"
    sql = sql & "   HAVING COUNT(specialistname+diseasename+new) >= " & mn
    sql = sql & "   AND  COUNT(specialistname+diseasename+old) >= " & mn
    sql = sql & "   ORDER BY specialistname, new_frequency DESC, [old_frequency] DESC"
    getQuery = sql
End Function
Sub displayStyle()
    css = " <style>    .table-report{"
css = css & "position: relative;"
css = css & "border-collapse: collapse;"
css = css & "font-family:sans-serif;"
css = css & "font-size:small;"
css = css & "}"
css = css & ".table-report td,.table-report th{"
css = css & "border: 1px solid gray;"
css = css & "padding: 5px 10px 5px 10px;"
css = css & "}"
css = css & ".table-report thead{"
css = css & "position: sticky;"
css = css & "top: 0;"
css = css & "background-color:whitesmoke;"
css = css & "}"
css = css & ".table-report tbody>:nth-child(even){"
css = css & "background-color: white;"
css = css & "}"
css = css & "@media print{"
css = css & ".table-report thead{"
css = css & "position:unset;"
css = css & "}"
css = css & "}"
css = css & ".table-report tbody>:nth-child(odd){"
css = css & "background-color: whitesmoke;"
css = css & "} </style>"
response.write css
End Sub
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
