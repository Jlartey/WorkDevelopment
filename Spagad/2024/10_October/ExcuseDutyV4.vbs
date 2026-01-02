Sub populateTable()
    Dim count, sql, rst, PatientName, admissionDate, dischargeDate, operation, duty, medOfficer
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT DISTINCT ER.EMRRequestID FROM EMRRequest ER "
    sql = sql & "JOIN EMRResults ERT ON ERT.EMRRequestID = ER.EMRRequestID "
    sql = sql & "WHERE ERT.EMRDataID = 'AC002' AND ER.VisitationID = '" & visitationID & "'"

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
        If .RecordCount > 0 Then
            PatientName = getEMRResult(.fields("EMRRequestID"), "AC002", "AC00201", "column2")
            admissionDate = getEMRResult(.fields("EMRRequestID"), "AC002", "AC00202", "column2")
            dischargeDate = getEMRResult(.fields("EMRRequestID"), "AC002", "AC00202", "column5")
            operation = getEMRResult(.fields("EMRRequestID"), "AC002", "AC00203", "column2")
            duty = getEMRResult(.fields("EMRRequestID"), "AC002", "AC00203", "column5")
            medOfficer = getEMRResult(.fields("EMRRequestID"), "AC002", "AC00204", "column6")

            response.write "<div class='container'>"
            response.write "  <table class='excuse-duty-table'>"
            response.write "    <thead>"
            response.write "      <tr>"
            response.write "        <th colspan='2' class='table-heading'>Excuse Duty</th>"
            response.write "      </tr>"
            response.write "    </thead>"
            response.write "    <tbody>"

            response.write "      <tr>"
            response.write "        <td class='label-cell'>Name</td>"
            response.write "        <td class='value-cell'>" & PatientName & "</td>"
            response.write "      </tr>"

            response.write "      <tr>"
            response.write "        <td class='label-cell'>Date of Admission</td>"
            response.write "        <td class='value-cell'>" & FormatDate(admissionDate) & "</td>"
            response.write "      </tr>"

            response.write "      <tr>"
            response.write "        <td class='label-cell'>Date of Discharge</td>"
            response.write "        <td class='value-cell'>" & FormatDate(dischargeDate) & "</td>"
            response.write "      </tr>"

            response.write "      <tr>"
            response.write "        <td class='label-cell'>Operation</td>"
            response.write "        <td class='value-cell'>" & operation & "</td>"
            response.write "      </tr>"

            response.write "      <tr>"
            response.write "        <td class='label-cell'>Excuse Duty</td>"
            response.write "        <td class='value-cell'>" & duty & "</td>"
            response.write "      </tr>"

            response.write "      <tr>"
            response.write "        <td class='label-cell'>Medical Officer</td>"
            response.write "        <td class='value-cell'>" & medOfficer & "</td>"
            response.write "      </tr>"

            response.write "    </tbody>"
            response.write "  </table>"
            response.write "</div>"

        Else
            response.write "<h3 class='container'>No records found for this visit</h3>"
        End If
        
        .Close
    End With
    
    Set rst = Nothing
End Sub

http://192.168.2.14/hms/wpgPrtPrintLayoutAll.asp?PrintLayoutName=ExcuseDuty&PositionForTableName=Visitation&VisitationID=V1240930079