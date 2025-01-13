Function getInvestigationsCost(visitNo)
    Dim investigationCost, sql
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "select labtestid,sum(qty) as qty,avg(unitcost) as unt,sum(finalamt) as tot from ("
    sql = sql & "select labtestid, qty, unitcost, finalamt from investigation where visitationid='" & visitNo & "'"
    sql = sql & "union all "
    sql = sql & "select labtestid, qty, unitcost, finalamt from investigation2 where visitationid='" & visitNo & "') as t"
    sql = sql & " group by labtestid"
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
           .MoveFirst
           investigationCost = .fields("tot")
        End If
        .Close
    End With
    Set rst = Nothing
    getInvestigationsCost = investigationCost
End Function

Sub displayData()
    Dim insuredPatientID, insuredPrincipalID, insuranceNo, sponsorID, sql
    Dim visitNo, admissionDate, dischargeDate, wardID, admissionID
    
    sponsorID = request.querystring("sponsorID")

    sql = "SELECT InsuredPatient.InsuredPatientID, InsuredPatient.InsuredPrincipalID, InsuredPatient.insuranceNo, "
    
    sql = sql & "Admission.AdmissionID, Admission.VisitationID, CONVERT (VARCHAR(150), Admission.AdmissionDate,103) AS AdmissionDate, "
    
    sql = sql & "CONVERT (VARCHAR(150), Admission.DischargeDate, 103) AS DischargeDate, Admission.WardID "
    sql = sql & "From InsuredPatient INNER JOIN Admission "
    sql = sql & "ON InsuredPatient.PatientID = Admission.PatientID "
    sql = sql & "WHERE InsuredPatient.SponsorID = '" & sponsorID & "' "
    sql = sql & "AND Admission.WorkingMonthID = 'MTH202207'"
    
    Set rst = Server.CreateObject("ADODB.Recordset")
    With rst
      .open sql, conn, 3, 4

      If Not .EOF Then
      .MoveFirst
      Do While Not .EOF
        insuredPatientID = .fields("InsuredPatientID")
        insuredPrincipalID = .fields("InsuredPrincipalID")
        insuranceNo = .fields("InsuranceNo")
        visitNo = .fields("VisitationID")
        admissionDate = .fields("AdmissionDate")
        dischargeDate = .fields("DischargeDate")
        wardID = .fields("WardID")
        admissionID = .fields("AdmissionID")
        
        Response.Write "<tr>"
          Response.Write "<td class='mytd'>" & GetComboName("InsuredPatient", insuredPatientID) & "</td>"
          Response.Write "<td class='mytd' align='center'>" & GetComboName("InsuredPrincipal", insuredPrincipalID) & "</td>"
          Response.Write "<td class='mytd' align='center'>" & insuranceNo & "</td>"
          Response.Write "<td class='mytd'>" & visitNo & "</td>"
          Response.Write "<td class='mytd'>" & admissionDate & "</td>"
          Response.Write "<td class='mytd'>" & dischargeDate & "</td>"
          Response.Write "<td class='mytd'>" & GetComboName("Ward", wardID) & "</td>"
          Response.Write "<td class='mytd' align='center'>" & GetVisitCost(visitNo) & "</td>"
          Response.Write "<td class='mytd' align='center'>" & getInvestigationsCost(visitNo) & "</td>"
          Response.Write "<td class='mytd' align='center'>" & getMedItemsCost(visitNo) & "</td>"
          Response.Write "<td class='mytd' align='center'>" & getConsumablesCost(visitNo) & "</td>"
          Response.Write "<td class='mytd'><a href='wpgPrtPrintLayoutAll.asp?PrintLayoutName=Admission1&PositionForTableName=Admission&AdmissionID=" & admissionID & "&WorkFlowNav=POP' target='_blank'>View Bill</a></td>"
        Response.Write "</tr>"
        .MoveNext
      Loop
      Else
        Response.Write "No records found"
      End If
      .Close
    End With
End Sub


Response.Write " <script language='JavaScript'>"
Response.Write "   const fromDate = document.getElementById('From');"
Response.Write "   const toDate = document.getElementById('To');"
Response.Write "   document.getElementById('process').onclick = function() {"
Response.Write "   console.log(fromDate, toDate) "
Response.Write "   };"
Response.Write "</script>"