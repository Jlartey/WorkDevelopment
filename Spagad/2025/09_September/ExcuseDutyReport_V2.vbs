'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim VisitationID, patient, age, Gender, dateOfAdmission, dateOfDischarge, excuseFrom, excuseTo, isReviewRequired, reviewDate, comments
Dim practitioner, signature, stamp, emrRequestID

VisitationID = Request.QueryString("VisitationID")
Dim sql, rst

    Set rst = CreateObject("ADODB.Recordset")

    sql = "WITH PatientDetails AS ( "
    sql = sql & "SELECT p.PatientName, p.BirthDate, p.genderID, v.VisitationID "
    sql = sql & "FROM Visitation v "
    sql = sql & "JOIN Patient p "
    sql = sql & "ON v.PatientID = p.PatientID "
    sql = sql & "WHERE v.VisitationID = '" & VisitationID & "' "
    sql = sql & "), "
    sql = sql & "EmrRecords AS ( "
    sql = sql & "SELECT TOP 1 EMRRequestID, VisitationID "
    sql = sql & "From EMRRequestItems "
    sql = sql & "WHERE EMRDataID = 'IM081' "
    sql = sql & "AND VisitationID = '" & VisitationID & "' "
    sql = sql & ") "
    sql = sql & "SELECT "
    sql = sql & "pd.PatientName,  "
    sql = sql & "DATEDIFF(YEAR, pd.BirthDate, GETDATE()) - "
        sql = sql & "CASE "
            sql = sql & "WHEN DateAdd(Year, DateDiff(Year, pd.BirthDate, getDate()), pd.BirthDate) > getDate() "
            sql = sql & "THEN 1 "
            sql = sql & "ELSE 0 "
        sql = sql & "END AS Age, "
    sql = sql & "pd.genderID, "
    sql = sql & "emr.emrRequestID "
    sql = sql & "FROM PatientDetails pd "
    sql = sql & "JOIN EmrRecords emr "
    sql = sql & "ON pd.VisitationID = emr.VisitationID; "
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
        If .recordCount > 0 Then
            emrRequestID = .fields("EMRRequestID")
            patient = .fields("PatientName")
            age = .fields("Age")
            Gender = GetComboName("Gender", .fields("genderID"))
            dateOfAdmission = getEMRResult(emrRequestID, "IM081", "IM081.1 ", "Column2")
            dateOfDischarge = getEMRResult(emrRequestID, "IM081", "IM081.1 ", "Column5")
            excuseFrom = getEMRResult(emrRequestID, "IM081", "IM081.2", "Column2")
            excuseTo = getEMRResult(emrRequestID, "IM081", "IM081.2", "Column5")
            isReviewRequired = GetComboName("EMRVar3B", getEMRResult(emrRequestID, "IM081", "IM081.3", "Column2"))
            reviewDate = getEMRResult(emrRequestID, "IM081", "IM081.3", "Column5")
            comments = getEMRResult(emrRequestID, "IM081", "IM081.4", "Column2")
            practitioner = getEMRResult(emrRequestID, "IM081", "IM081.5", "Column2")
            signature = getEMRResult(emrRequestID, "IM081", "IM081.6", "Column2")
            stamp = getEMRResult(emrRequestID, "IM081", "IM081.7", "Column2")
        End If
        
        .Close
    End With
    
    Set rst = Nothing

    Styles
    Report

Sub Styles()
  response.write "  <style>"
  response.write "    .container {"
  response.write "      font-family: Arial, sans-serif;"
  response.write "      line-height: 1.5;"
  response.write "      width: 650px;"
  response.write "      margin: 0 auto;"
  response.write "    }"
  response.write "    .hospital-name {"
  response.write "      text-align: center;"
  response.write "    }"
  response.write "    .stamp {"
  response.write "      width: 250px;"
  response.write "      height: 70px;"
  response.write "      border: 3px solid black;"
  response.write "    }"
  response.write ""
  response.write "    .label2 {"
  response.write "      width: 160px;"
  response.write "      display: inline-block;"
  response.write "      font-weight: bold;"
  response.write "    }"
  response.write ""
  response.write "    .mytextarea {"
  response.write "      font-size: 0.8rem;"
  response.write "      letter-spacing: 1px;"
  response.write "      padding: 10px;"
  response.write "      width: 95%;"
  response.write "      line-height: 1.5;"
  response.write "      border-radius: 5px;"
  response.write "      border: 1px solid #cccccc;"
  response.write "      box-shadow: 1px 1px 1px #999999;"
  response.write "    }"
  
response.write "  </style>"
End Sub

Sub Report()
  response.write "  <div class='container'>"
  response.write "    <div style=""display: flex; margin-top: 20px"">"
  response.write "      <div style=""margin-top: 20px"">"
  response.write "        <img"
  response.write "          src=""images/banner3.bmp"""
  response.write "          alt=""IMAH Logo"""
  response.write "          width=""100"""
  response.write "          height=""100"""
  response.write "        />"
  response.write "      </div>"
  response.write ""
  response.write "      <div class=""hospital-name"">"
  response.write "        <h3 style=""margin-top: 30px"">INTERNATIONAL MARITIME HOSPITAL (IMaH)</h3>"
  response.write "        <h3 style=""font-weight: 700; margin-top: -10px"">"
  response.write "          MEDICAL EXCUSE DUTY FORM"
  response.write "        </h3>"
  response.write "      </div>"
  response.write "    </div>"
  response.write ""
  response.write "    <div>"
  response.write "      <!-- <div></div> -->"
  response.write "      <label class=""label2 style"">Name </label>"
  response.write "      <span style=""margin-left: -5px"">: " & patient & "</span> <br />"
  response.write "      <div style=""display: flex"">"
  response.write "        <label class=""label2""> Age &nbsp;</label>"
  response.write "        <span class=""value"" style=""margin-right: 187px"">: " & age & "</span>"
  response.write "        <label style=""font-weight: bold"">Gender</label>"
  response.write "        <span>&nbsp;: " & Gender & "</span>"
  response.write "      </div>"
  response.write ""
  response.write "      <div></div>"
  response.write "      <div>"
  response.write "        <label class=""label2"">Date of Admission</label"
  response.write "        ><span>: " & dateOfAdmission & "</span>"
  response.write "      </div>"
  response.write "      <label class=""label2"">Date of Discharge</label"
  response.write "      ><span>: " & dateOfDischarge & "</span>"
  response.write ""
  response.write "      <div style=""display: flex"">"
  response.write "        <div style=""margin-right: auto"">"
  response.write "          <label for class=""label2"">Excuse Duty From</label>"
  response.write "          <span style=""margin-left: -5px"">: " & excuseFrom & "</span>"
  response.write "        </div>"
  response.write "        <div>"
  response.write "          <label style=""font-weight: bold"">To</label>"
  response.write "          <span>: " & excuseTo & "</span>"
  response.write "        </div>"
  response.write "      </div>"
  response.write ""
  response.write "      <div style=""display: flex; justify-content: space-between"">"
  response.write "        <div>"
  response.write "          <label class=""label2"">Review Required</label>"
  response.write "          <span style=""margin-left: -5px"">: " & isReviewRequired & "</span>"
  response.write "        </div>"
  response.write ""
  response.write "        <div>"
  response.write "          <label style=""font-weight: bold"">Review Date</label>"
  response.write "          <span>: " & reviewDate & "</span>"
  response.write "        </div>"
  response.write "      </div>"
  response.write ""
  response.write "      <div>"
  response.write "        <label class=""comments-label"" for=""comments"" style=""font-weight: bold"""
  response.write "          >Comments:</label"
  response.write "        ><br />"
  response.write "        <textarea class=""mytextarea"" name=""comments"" id=""comments"" rows=""4"" cols=""33"">"
  response.write "  " & comments & "</textarea"
  response.write "        >"
  response.write "      </div>"
  response.write ""
  response.write "      <div style=""margin-top: 10px"">"
  response.write "        <label class=""label2"">Name Of Practitioner</label>"
  response.write "        <span>: " & practitioner & " </span>"
  response.write "      </div>"
  response.write ""
  response.write "      <div"
  response.write "        style=""display: flex; justify-content: space-between; margin-top: 20px"""
  response.write "      >"
  response.write "        <div>"
  response.write "          <label style=""font-weight: bold"">Practitioner's Signature:</label>"
  response.write "          <div style=""margin-top: 40px""></div>"
  response.write "          <span>.......................................</span>"
  response.write "        </div>"
  response.write ""
  response.write "        <div style=""display: flex"">"
  response.write "          <p style=""margin-right: 10px; margin-top: 25px"">STAMP:</p>"
  response.write "          <div class=""stamp""></div>"
  response.write "        </div>"
  response.write "      </div>"
  response.write "    </div>"
  response.write "  </div>"
End Sub

Function getEMRResult(emrRequestID, emrDataID, CompID, column)
    Dim sql, rst, emrValue
    Set rst = server.CreateObject("ADODB.Recordset")
    emrValue = ""
    sql = "SELECT * FROM emrresults WHERE emrrequestid ='" & emrRequestID & "'"
    sql = sql & " AND emrdataid ='" & emrDataID & "' AND emrcomponentid='" & CompID & "'"
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .recordCount > 0 Then
            .MoveFirst
            If Not IsNull(.fields(column)) Then
                emrValue = Trim(.fields(column))
            End If
            If IsNull(.fields(column)) Then
                emrValue = "Null"
            End If
        End If
        .Close
    End With
    getEMRResult = emrValue
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

