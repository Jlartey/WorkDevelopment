AddCSSS
Dim sql, sql1, rst, rst1, datePeriod, cnt, vst, sponsor, gen, spclst, 
Dim spclstTpe, spclstGrp, vstTpe, specialist, whcls, WorkingMonthID
datePeriod = Split(Request.QueryString("PrintFilter"), "||")

Sub SRS22
  cnt = 0
Set rst = CreateObject("ADODB.RecordSet")

sql = "SELECT DISTINCT EMRRequestID, PatientID FROM EMRRequestItems "
' sql = sql & " WHERE EMRDate1 BETWEEN '01 june 2024 00:00:00' AND '30 june 2024 23:59:59' "
sql = sql & " WHERE EMRDate1 BETWEEN '" & datePeriod(0) & "' AND '" & datePeriod(1) & "' "
sql = sql & " AND EMRDataID = 'RES018' "
With rst
    .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
                    dispOption
                    response.write "<table id='myTbl'>"
                    response.write "    <thead>"
                    response.write "        <tr>"
                    response.write "            <th colspan='8'>" & GetComboName("EMRData", "RES018") & " Scores Report Between " & FormatDate(datePeriod(0)) & " and " & FormatDate(datePeriod(1)) & "</th>"
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
                    ' PatientName = GetComboName("Patient", .fields("PatientID"))
                    response.write "     <tr>"
                    response.write "       <td>" & cnt & "</td>"
                    response.write "       <td>" & PatientID & "</td>"
                    response.write "       <td>" & GetComboName("Patient", PatientID) & "</td>"
                    response.write "       <td>" & getScore(EMRRequestID) & "</td>"
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