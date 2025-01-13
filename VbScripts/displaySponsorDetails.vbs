' Code that I commented out
'Sub displayNameOfPrincipal()
'    Dim principalName, sponsorID
'    sponsorID = request.querystring("sponsorID")
'
'    sql = "SELECT InsuredPatientID FROM InsuredPatient WHERE SponsorID = '" & sponsorID & "'"
'
'    Set rst = Server.CreateObject("ADODB.Recordset")
'    With rst
'        .open sql, conn, 3, 4
'        If Not .EOF Then
'        .MoveFirst
'        Do While Not .EOF
'            principalName = .fields("InsuredPrincid")
'
'            response.write "<tr>"
'                response.write "  <td>" & GetComboName("InsuredPatient", patientid) & "</td>"
'            response.write "</tr>"
'        Loop
'        End If
'        .Close
'    End With
'End Sub

Sub displayData()
    Dim insuredPatientID, insuredPrincipalID, insuranceNo, sponsorID, sql
    Dim visitNo, admissionDate, dischargeDate, wardID
    Dim conn, rst

    sponsorID = Request.QueryString("sponsorID")

    sql = "SELECT InsuredPatient.InsuredPatientID, InsuredPatient.InsuredPrincipalID, InsuredPatient.insuranceNo, "
    sql = sql & "Admission.VisitationID, CONVERT(VARCHAR(150), Admission.AdmissionDate, 103) AS AdmissionDate, "
    sql = sql & "CONVERT(VARCHAR(150), Admission.DischargeDate, 103) AS DischargeDate, Admission.WardID "
    sql = sql & "FROM InsuredPatient INNER JOIN Admission "
    sql = sql & "ON InsuredPatient.PatientID = Admission.PatientID "
    sql = sql & "WHERE InsuredPatient.SponsorID = '" & sponsorID & "' "
    sql = sql & "AND Admission.WorkingMonthID = 'MTH202207'"

    Set conn = Server.CreateObject("ADODB.Connection")
    ' Assume conn is opened properly here (provide the necessary connection string and open the connection)
    conn.Open "your_connection_string"

    Set rst = Server.CreateObject("ADODB.Recordset")
    With rst
        .Open sql, conn, 3, 4

        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                insuredPatientID = .Fields("InsuredPatientID")
                insuredPrincipalID = .Fields("InsuredPrincipalID")
                insuranceNo = .Fields("insuranceNo")
                visitNo = .Fields("VisitationID")
                admissionDate = .Fields("AdmissionDate")
                dischargeDate = .Fields("DischargeDate")
                wardID = .Fields("WardID")

                Response.Write "<tr>"
                Response.Write "<td>" & GetComboName("InsuredPatient", insuredPatientID) & "</td>"
                Response.Write "<td>" & GetComboName("InsuredPrincipal", insuredPrincipalID) & "</td>"
                Response.Write "<td>" & insuranceNo & "</td>"
                Response.Write "<td>" & visitNo & "</td>"
                Response.Write "<td>" & admissionDate & "</td>"
                Response.Write "<td>" & dischargeDate & "</td>"
                Response.Write "<td>" & GetComboName("Ward", wardID) & "</td>"
                Response.Write "</tr>"
                .MoveNext
            Loop
        Else
            Response.Write "No records found"
        End If
        .Close
    End With

    conn.Close
    Set rst = Nothing
    Set conn = Nothing
End Sub
