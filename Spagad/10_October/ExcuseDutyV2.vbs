'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
'response.write "Hello Joe"

Dim visitationID
visitationID = Request.querystring("VisitationID")

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
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then

          PatientName = getEMRResult(.fields("EMRRequestID"), "AC002", "AC00201", "column2")
          admissionDate = getEMRResult(.fields("EMRRequestID"), "AC002", "AC00202", "column2")
          dischargeDate = getEMRResult(.fields("EMRRequestID"), "AC002", "AC00202", "column5")
          operation = getEMRResult(.fields("EMRRequestID"), "AC002", "AC00203", "column2")
          duty = getEMRResult(.fields("EMRRequestID"), "AC002", "AC00203", "column5")
          medOfficer = getEMRResult(.fields("EMRRequestID"), "AC002", "AC00204", "column6")

          response.write "  <div class='container'> "
          response.write "    <header class='frm-header'> "
          response.write "        <img src='images/letterhead5.jpg' alt='Clinic Logo' class='clinic-logo'> "
        '   response.write "        <h1>Airport Clinic Limited</h1> "
        '   response.write "        <p>Private Mail Bag</p>  "
        '   response.write "        <p>Kotoka International Airport, Accra, Ghana</p> "
        '   response.write "        <p>Tel: (+233-302) 764987 / 0505 527641</p> "
          response.write "    </header> "

          response.write "    <h2>Excuse Duty Form</h2> "

          response.write "    <div class='frm-detail'> "
          response.write "<div style='text-align: left;'>"
          response.write "        <label class='frm-label' for='name'>Name:</label> "
          response.write "        & PatientName & "> "
          response.write "</div>"

          response.write "        <label class='frm-label' for='admission'>Date of Admission:</label> "
          response.write "        <input type='text' id='admission' name='admission' value='" & FormatDate(admissionDate) & "'> "

          response.write "        <label class='frm-label' for='discharge'>Date of Discharge:</label> "
          response.write "        <input type='text' id='discharge' name='discharge' value='" & FormatDate(dischargeDate) & "'> "

          response.write "        <label class='frm-label' for='operation'>Operation:</label> "
          response.write "        <input type='text' id='operation' name='operation' value='" & operation & "'> "

          response.write "        <label class='frm-label' for='excuse'>Excuse Duty:</label> "
          response.write "        <input type='text' id='excuse' name='excuse' value='" & duty & "'> "

          response.write "        <div class='signature'> "
          response.write "            <p>Medical Officer:</p> "
          response.write "            <input type='text' id='medicalOfficer' name='medicalOfficer' value='" & medOfficer & "'> "
          response.write "            <label for='date'>Date:</label> "
          response.write "            <input type='text' id='date' name='date' value='" & FormatDate(Now()) & "'> "
          response.write "        </div> "
          response.write "    </div> "
        response.write "  </div> "
        Else
            response.write "<h3 class='container'>No records found for this visit</h3>"
        End If
        
        .Close
    End With
    
    Set rst = Nothing
End Sub


Function getEMRResult(EMRRequestID, emrDataID, CompID, column)

    Dim sql, rst
    Set rst = Server.CreateObject("ADODB.Recordset")
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

Sub tableStyles()
  response.write "  <style> "
  response.write "  * { "
  response.write "      margin: 0; "
  response.write "      padding: 0; "
  response.write "      box-sizing: border-box; "
  response.write "  } "
  response.write "  body { "
  response.write "      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; "
  response.write "      background-color: #f4f4f9; "
  response.write "      padding: 40px; "
  response.write "  } "
  response.write "  .container { "
  response.write "      max-width: 600px; "
  response.write "      background-color: #fff; "
  response.write "      padding: 30px; "
  response.write "      border-radius: 10px; "
  response.write "      box-shadow: 0 0 15px rgba(0, 0, 0, 0.1); "
  response.write "      margin: 0 auto; "
  response.write "      border-top: 5px solid #007bff; "
  response.write "  } "
  response.write "  .frm-header { "
  response.write "      text-align: center; "
'   response.write "      margin-bottom: 30px; "
  response.write "  } "
  response.write "  .clinic-logo { "
  response.write "      width: 540px; "
  response.write "      margin-bottom: 10px; "
  response.write "  } "
  response.write "  .frm-header h1 { "
  response.write "      font-size: 1.8rem; "
  response.write "      margin-bottom: 5px; "
  response.write "      color: #333; "
  response.write "  } "
  response.write "  .frm-header p, .signature p { "
  response.write "      font-size: 0.95rem; "
  response.write "      color: #555; "
  response.write "      margin-bottom: 5px; "
  response.write "  } "
  response.write "  .container h2 { "
  response.write "      text-align: center; "
  response.write "      margin-bottom: 30px; "
  response.write "      font-size: 1.5rem; "
  response.write "      color: #333; "
  response.write "  } "
  response.write "  .frm-detail { "
  response.write "      display: flex; "
  response.write "      flex-direction: column; "
  response.write "  } "
  response.write "  .frm-detail label { "
  response.write "      font-size: 1rem; "
  response.write "      color: #333; "
  response.write "      margin-bottom: 5px; "
  response.write "      text-align: left; "
  response.write "  } "
  response.write "  .frm-detail input[type='text'] { "
  response.write "      padding: 10px 5px; "
  response.write "      margin-bottom: 30px; "
  response.write "      border: none; "
  response.write "      border-bottom: 2px solid #007bff; "
  response.write "      font-size: 1rem; "
  response.write "      color: #333; "
  response.write "      background-color: transparent; "
  response.write "      outline: none; "
  response.write "      transition: border-bottom 0.3s ease; "
  response.write "  } "
  response.write "  .frm-detail input[type='text']:focus { "
  response.write "      border-bottom: 2px solid #0056b3; "
  response.write "  } "
  response.write "  .signature { "
  response.write "      display: flex; "
  response.write "      justify-content: space-between; "
  response.write "      align-items: center; "
  response.write "  } "
  response.write "  .signature p { "
  response.write "      font-size: 1rem; "
  response.write "      color: #333; "
  response.write "  } "
  response.write "  .signature input[type='text'] { "
  response.write "      width: calc(100% - 160px); "
  response.write "  } "
  response.write "  .signature input[type='date'] { "
  response.write "      width: 150px; "
  response.write "  } "
  response.write "  </style> "
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>


