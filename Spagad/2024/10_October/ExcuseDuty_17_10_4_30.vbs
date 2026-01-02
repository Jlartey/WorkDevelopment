'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
'response.write "Hello Joe"

Dim visitationID
visitationID = Request.queryString("VisitationID")

tableStyles
populateTable

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

'            response.write "<div class='container'>"
            response.write "    <header class='frm-header' style='margin-bottom: 16px; position: relative'> "
            response.write "        <img src='images/letterhead5.jpg' alt='Clinic Logo' class='clinic-logo'> "
            response.write "        <div style=""text-align: right; font-size:16px; translate: -10px; margin-top: 20px; position: absolute; top: -15px; right: 10px; font-weight: bold; font-family: Arial, Helvetica, sans-serif; ""> ACL/ED/" & Mid(visitationID, 2, 10) & "</div>"
            response.write "    </header> "
            
            response.write "  <table class='report'>"
            response.write "    <tbody>"
            response.write "      <tr>"
            response.write "        <td colspan='2' class='header'>Excuse Duty</td>"
            response.write "      </tr>"
            
            response.write "      <tr>"
            response.write "        <td>Name</td>"
            response.write "        <td>" & PatientName & "</td>"
            response.write "      </tr>"

            response.write "      <tr>"
            response.write "        <td>Date of Admission</td>"
            response.write "        <td>" & FormatDate(admissionDate) & "</td>"
            response.write "      </tr>"

            response.write "      <tr>"
            response.write "        <td >Date of Discharge</td>"
            response.write "        <td >" & FormatDate(dischargeDate) & "</td>"
            response.write "      </tr>"

            response.write "      <tr>"
            response.write "        <td>Operation</td>"
            response.write "        <td>" & operation & "</td>"
            response.write "      </tr>"

            response.write "      <tr>"
            response.write "        <td>Excuse Duty</td>"
            response.write "        <td>" & duty & "</td>"
            response.write "      </tr>"

            response.write "      <tr>"
            response.write "        <td>Medical Officer</td>"
            response.write "        <td>" & medOfficer & "</td>"
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

Sub tableStyles()
    response.write "    <style>"
    response.write "        .report {"
    response.write "            border-collapse: separate;"
    response.write "            width: 100%;"
    response.write "            border-spacing: 0px;"
    response.write "            font-family: Arial, Helvetica, sans-serif; "
    response.write "            font-size: 1rem; "
    response.write "        }"
    response.write "        .report >tbody>tr>td {"
'    Response.Write "            font-size:13px;"
    response.write "            text-align: left;"
    response.write "            padding: 8px 10px;"
    response.write "            border-left: 1px solid silver;"
    response.write "            border-bottom: 1px solid silver;"
    response.write "        }"
    response.write "        .report>tbody>tr:first-child td {"
    response.write "            border-top: 1px solid silver;"
    response.write "        }"
    response.write "        .report>tbody>tr>td:last-child {"
    response.write "            border-right: 1px solid silver;"
    response.write "        }"
    response.write "        .report>tbody>tr>td.header {"
    response.write "            background-color: #f0f0f0;"
    response.write "            font-weight: bold;"
    response.write "            text-align: center;"
    response.write "            text-transform: uppercase;"
    response.write "        }"
    response.write ""
    response.write "        * {"
    response.write "            -webkit-print-color-adjust: exact;"
    response.write "            print-color-adjust: exact;"
    response.write "            color-adjust: exact !important;"
    response.write "        }"
    response.write "    </style>"
End Sub

Function getEMRResult(EMRRequestID, emrDataID, CompID, column)

    Dim sql, rst
    Set rst = server.CreateObject("ADODB.Recordset")
    getEMRResult = ""

    sql = "SELECT * FROM emrresults WHERE emrrequestid ='" & EMRRequestID & "'"
    sql = sql & " AND emrdataid ='" & emrDataID & "' AND emrcomponentid='" & CompID & "'"
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            If Not IsNull(.fields(column)) Then
                getEMRResult = Trim(.fields(column))
            End If
            If IsNull(.fields(column)) Then
                getEMRResult = "Null"
            End If
        End If
    End With
    Set rst = Nothing

End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
