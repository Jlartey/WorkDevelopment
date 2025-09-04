'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>



Dim dob, gender, name, weight, tel, visitationID

visitationID = Trim(Request.QueryString("VisitationID"))

Styles
DrugReactionReport


Sub DrugReactionReport()
    Dim sql, rst

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT convert(VARCHAR(20), Visitation.BirthDate, 106)DATEOFBIRTH, "
    sql = sql & "Gender.GenderName, "
    sql = sql & "Patient.PatientName, "
    sql = sql & "Weight, "
    sql = sql & "Patient.ResidencePhone "
    sql = sql & "From Visitation "
    sql = sql & "Join gender "
    sql = sql & "ON Gender.GenderID = Visitation.GenderID "
    sql = sql & "Join Patient "
    sql = sql & "ON Patient.PatientID = Visitation.PatientID "
    sql = sql & "WHERE VisitationID = '" & visitationID & "'"

    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
        If .RecordCount > 0 Then
            dob = .fields("DATEOFBIRTH")
            gender = .fields("GenderName")
            name = .fields("PatientName")
            weight = .fields("Weight")
            tel = .fields("ResidencePhone")
        End If
        
        .Close
    End With
    
    Set rst = Nothing
    
    Set rst1 = CreateObject("ADODB.Recordset")
    
    Dim sql2, rst2, isPregnant, pregnancyAge, description, reactionStart, reactionEnd, outcome
    Dim untowardMedCon, adverseRecSpec, seriousNess, brandName, genericName, batch, expiryDate
    Dim manufacturer, reasons, dosage, noOfDays, route, dateStarted, dateEnded, isSubsided, isPrescribed
    Dim source, isReused, repName, repPro, repAddress, repTel, repEmail
   
    sql2 = "SELECT TOP 1 EMRRequestID "
    sql2 = sql2 & "From EMRRequestItems WHERE EMRDataID = 'ADV001' "
    sql2 = sql2 & "AND VisitationID = '" & visitationID & "'"

    With rst1
        .open qryPro.FltQry(sql2), conn, 3, 4
        
        If .RecordCount > 0 Then
           emrRequestID = .fields("EMRRequestID")
           isPregnant = GetComboName("EMRVar3B", getEMRResult(emrRequestID, "ADV001", "ADV001.0", "Column2"))
           pregnancyAge = getEMRResult(emrRequestID, "ADV001", "ADV001.0", "Column5")
           description = getEMRResult(emrRequestID, "ADV001", "ADV001.2", "Column1")
           reactionStart = getEMRResult(emrRequestID, "ADV001", "ADV001.3", "Column2")
           reactionEnd = getEMRResult(emrRequestID, "ADV001", "ADV001.3", "Column5")
           outcome = GetComboName("EMRVar3B", getEMRResult(emrRequestID, "ADV001", "ADV001.5", "Column1"))
           untowardMedCon = GetComboName("EMRVar3B", getEMRResult(emrRequestID, "ADV001", "ADV001.6", "Column2"))
           adverseRecSpec = getEMRResult(emrRequestID, "ADV001", "ADV001.6", "Column5")
           seriousNess = GetComboName("EMRVar3B", getEMRResult(emrRequestID, "ADV001", "ADV001.7", "Column2"))
           brandName = getEMRResult(emrRequestID, "ADV001", "ADV001.9", "Column2")
           genericName = getEMRResult(emrRequestID, "ADV001", "ADV001.9", "Column5")
           batch = getEMRResult(emrRequestID, "ADV001", "ADV001.10", "Column2")
           expiryDate = getEMRResult(emrRequestID, "ADV001", "ADV001.10", "Column5")
           manufacturer = getEMRResult(emrRequestID, "ADV001", "ADV001.11", "Column2")
           reasons = getEMRResult(emrRequestID, "ADV001", "ADV001.11", "Column5")
           dosage = getEMRResult(emrRequestID, "ADV001", "ADV001.12", "Column2")
           noOfDays = getEMRResult(emrRequestID, "ADV001", "ADV001.12", "Column5")
           route = getEMRResult(emrRequestID, "ADV001", "ADV001.13", "Column2")
           dateStarted = getEMRResult(emrRequestID, "ADV001", "ADV001.14", "Column2")
           dateEnded = getEMRResult(emrRequestID, "ADV001", "ADV001.14", "Column5")
           isSubsided = GetComboName("EMRVar3B", getEMRResult(emrRequestID, "ADV001", "ADV001.15", "Column2"))
           isPrescribed = GetComboName("EMRVar3B", getEMRResult(emrRequestID, "ADV001", "ADV001.15", "Column5"))
           source = getEMRResult(emrRequestID, "ADV001", "ADV001.16", "Column2")
           isReused = GetComboName("EMRVar3B", getEMRResult(emrRequestID, "ADV001", "ADV001.16", "Column5"))
           isReappear = GetComboName("EMRVar3B", getEMRResult(emrRequestID, "ADV001", "ADV001.17", "Column2"))
           repName = getEMRResult(emrRequestID, "ADV001", "ADV001.23", "Column2")
           repPro = getEMRResult(emrRequestID, "ADV001", "ADV001.23", "Column5")
           repAddress = getEMRResult(emrRequestID, "ADV001", "ADV001.24", "Column2")
           repTel = getEMRResult(emrRequestID, "ADV001", "ADV001.25", "Column2")
           repEmail = getEMRResult(emrRequestID, "ADV001", "ADV001.25", "Column5")
           repDate = getEMRResult(emrRequestID, "ADV001", "ADV001.26", "Column2")
        End If
        
        .Close
    End With
    
    Set rst2 = Nothing
   
    Response.Write "  <div class=""container"">"
    Response.Write "    <div class=""form-header"">"
    Response.Write "      FDA/ADR <br />"
    Response.Write "      FOOD AND DRUGS AUTHORITY<br />"
    Response.Write "      In Strict Confidence"
    Response.Write "    </div>"
    Response.Write "    <h1>"
    Response.Write "      ADVERSE REACTION REPORTING FORM <br />"
    Response.Write "      (Please complete all sections as much as possible)"
    Response.Write "    </h1>"
    Response.Write ""
    Response.Write "    <div class=""section-title"">(A) PATIENT DETAILS</div>"
    Response.Write "    <div class=""section1"">"
    Response.Write "      <div class=""group1"">"
    Response.Write "        <p>"
    Response.Write "          <label for="""">Age/Date Of Birth</label> <span>:</span>"
    Response.Write "          <span class=""answer"">" & dob & "</span>"
    Response.Write "        </p>"
    Response.Write "        <p>"
    Response.Write "          <label for="""">Gender</label> <span>:</span>"
    Response.Write "          <span class=""answer"">" & gender & "</span>"
    Response.Write "        </p>"
    Response.Write "        <p>"
    Response.Write "          <label for="""">If Female, Pregnant</label> <span>:</span>"
    Response.Write "          <span class=""answer"">" & isPregnant & "</span>"
    Response.Write "        </p>"
    Response.Write "        <p>"
    Response.Write "          <label for="""">Name/Folder Number</label> <span>:</span>"
    Response.Write "          <span class=""answer"">" & name & "</span>"
    Response.Write "        </p>"
    Response.Write "        <p>"
    Response.Write "          <label for="""">Hospital/Treatment Centre</label> <span>:</span>"
    Response.Write "          <span class=""answer"">FOCOS Orthopaedic Hospital</span>"
    Response.Write "        </p>"
    Response.Write "      </div>"
    Response.Write ""
    Response.Write "      <div class=""group2"">"
    Response.Write "        <p>"
    Response.Write "          <label for="""">Wt(kg)</label> <span>:</span>"
    Response.Write "          <span class=""answer"">" & weight & "</span>"
    Response.Write "        </p>"
    Response.Write "        <p>"
    Response.Write "          <label for="""">Age of Pregnancy</label> <span>:</span>"
    Response.Write "          <span class=""answer"">" & pregnancyAge & "</span>"
    Response.Write "        </p>"
    Response.Write "        <p>"
    Response.Write "          <label for="""">Telephone No</label> <span>:</span>"
    Response.Write "          <span class=""answer"">" & tel & "</span>"
    Response.Write "        </p>"
    Response.Write "      </div>"
    Response.Write "    </div>"
    Response.Write "    <div class=""section-title"">"
    Response.Write "      (B) DETAILS OF ADVERSE REACTION AND ANY TREATMENT GIVEN<br />"
    Response.Write "      (Attach a separate sheet and all relevant laboratory tests/data where"
    Response.Write "      necessary)"
    Response.Write "    </div>"
    Response.Write "  </div>"
    Response.Write "  <textarea name=""textarea"" rows=""5"" cols=""15"">" & description & "</textarea>"
    Response.Write "  <div class=""dates"" style=""margin-bottom: 15px"">"
    Response.Write "    <label>Date reaction started(dd/mm/yyy):</label> <span>" & reactionStart & "</span>"
    Response.Write "    <label>Date reaction stopped(dd/mm/yyy):</label> <span>" & reactionEnd & "</span>"
    Response.Write "  </div>"
    
    Response.Write " <div class='section-c'>"
    Response.Write "  <div class=""section-title"">(C) OUTCOME OF ADVERSE REACTION</div>"
    Response.Write "  <p><label>Outcome:</label> <span>" & outcome & "</span></p>"
    Response.Write "  <p>"
    Response.Write "    Did the adverse result in any untoward medical condition?  <span>" & untowardMedCon & "</span>"
    Response.Write "  </p>"
    Response.Write "  <p>If yes, Specify: <span>" & adverseRecSpec & "</span></p>"
    Response.Write "  <h4>SERIOUSNESS</h4>"
    Response.Write "  <p><span>" & seriousNess & "</span></p>"
    Response.Write "</div>"
    
    Response.Write "  <div class=""section-title"">"
    Response.Write "    (D) SUSPECTED PRODUCT(S) (Attach sample or product label if available)"
    Response.Write "  </div>"
    Response.Write "  <table class='mytable'>"
    Response.Write "    <tr>"
    Response.Write "      <th class='myth'>Brand name</th>"
    Response.Write "      <th class='myth'>Generic name</th>"
    Response.Write "      <th class='myth'>Batch name</th>"
    Response.Write "      <th class='myth'>Expiry date</th>"
    Response.Write "      <th class='myth'>Manufacturer</th>"
    Response.Write "    </tr>"
    Response.Write "    <tr>"
    Response.Write "      <td class='mytd'>" & brandName & "</td>"
    Response.Write "      <td class='mytd'>" & genericName & "</td>"
    Response.Write "      <td class='mytd'>" & batch & "</td>"
    Response.Write "      <td class='mytd'>" & expiryDate & "</td>"
    Response.Write "      <td class='mytd'>" & manufacturer & "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr>"
    Response.Write "      <th class='myth' colspan=""2"">Reasons for use (Indication):</th>"
    Response.Write "      <th class='myth'>Dosage regimen:</th>"
    Response.Write "      <th class='myth'>No. of days given</th>"
    Response.Write "      <th class='myth'>Route of administration:</th>"
    Response.Write "    </tr>"
    Response.Write "    <tr>"
    Response.Write "      <td colspan=""2"" class='mytd'>" & reasons & "</td>"
    Response.Write "      <td class='mytd'>" & dosage & "</td>"
    Response.Write "      <td class='mytd'>" & noOfDays & "</td>"
    Response.Write "      <td class='mytd'>" & route & "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr>"
    Response.Write "      <td class='mytd' colspan=""5"">"
    Response.Write "        Date started (dd/mm/yyyy): <span>" & dateStarted & "</span> &nbsp; &nbsp; &nbsp;"
    Response.Write "        &nbsp; Date stopped (dd/mm/yyyy): <span>" & dateEnded & "</span>"
    Response.Write "        <br />"
    Response.Write "        <em"
    Response.Write "          >Did the adverse reaction subside when the drug was"
    Response.Write "          stopped(de-challenge)? <span>" & isSubsided & "</span></em"
    Response.Write "        >"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr>"
    Response.Write "      <td class='mytd' colspan=""3"">Was the product prescribed? <span>" & isPrescribed & "</span></td>"
    Response.Write "      <td class='mytd' colspan=""2"">Source of drug: <span>" & source & "</span></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr>"
    Response.Write "      <td colspan=""5"" class='mytd'>"
    Response.Write "        Was the product re-used after detection of adverse reaction"
    Response.Write "        (re-challenge)? <span class=""answer"">" & isReused & "</span>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr>"
    Response.Write "      <td colspan=""5"" class='mytd'>"
    Response.Write "        Did adverse reaction re-appear upon re-use?"
    Response.Write "        <span class=""answer"">" & isReappear & "</span>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write ""
    Response.Write "  <div class=""section-title"">"
    Response.Write "    (E) CONCOMITANT DRUGS: INCLUDING COMPLEMENTARY MEDICINES, ADMINISTERED"
    Response.Write "    <br />"
    Response.Write "    AT THE SAME TIME AND/OR 3 MONTHS BEFORE (Attach a separate sheet when"
    Response.Write "    necessary)"
    Response.Write "  </div>"
    Response.Write ""
    Response.Write "  <table class='mytable'>"
    Response.Write "    <tr>"
    Response.Write "      <th class='myth'>Name of Drug</th>"
    Response.Write "      <th class='myth'>Daily dose</th>"
    Response.Write "      <th class='myth'>Date started</th>"
    Response.Write "      <th class='myth'>Date stopped/Ongoing</th>"
    Response.Write "      <th class='myth'>Reason(s) for use</th>"
    Response.Write "    </tr>"
    
    Response.Write "    <tr>"
    
    Dim sql3, rst3
    Set rst3 = CreateObject("ADODB.Recordset")
    
    sql3 = "SELECT DrugId, PrescInfo1, convert(VARCHAR(20), PrescriptionDate, 106)[Date Started] FROM "
    sql3 = sql3 & "Prescription WHERE VisitationID = '" & visitationID & "' "
    
    With rst3
        .open sql3, conn, 3, 4
        
        If .RecordCount > 0 Then

            Do While Not .EOF
                Response.Write "<tr class='mytr'>"
                Response.Write "<td class='mytd'>" & GetComboName("Drug", .fields("DrugID")) & "</td>"
                Response.Write "<td class='mytd'>" & .fields("PrescInfo1") & "</td>"
                Response.Write "<td class='mytd'>" & .fields("Date Started") & "</td>"
                Response.Write "<td class='mytd'></td>"
                Response.Write "<td class='mytd'> Doctor's Prescription</td>"
                Response.Write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop
        
        End If
        
        .Close
    End With
    
    Set rst3 = Nothing
    Response.Write "  </table>"
    Response.Write ""
    Response.Write "  <div class=""section-title"">(F) REPORTER DETAILS</div>"
    Response.Write ""
    Response.Write "  <div class=""section1"">"
    Response.Write "    <div class=""group1"">"
    Response.Write "      <p>"
    Response.Write "        <label for="""">Name of Reporter</label> <span>:</span>"
    Response.Write "        <span class=""answer"">" & repName & "</span>"
    Response.Write "      </p>"
    Response.Write "      <p>"
    Response.Write "        <label for="""">Institution's Address</label> <span>:</span>"
    Response.Write "        <span class=""answer"">" & repAddress & "</span>"
    Response.Write "      </p>"
    Response.Write "      <p>"
    Response.Write "        <label for="""">Signature</label> <span>:</span>"
    Response.Write "        <span class=""answer"">_____________</span>"
    Response.Write "      </p>"
    Response.Write "      <p>"
    Response.Write "        <label for="""">Date</label> <span>:</span>"
    Response.Write "        <span class=""answer"">" & repDate & "</span>"
    Response.Write "      </p>"
    Response.Write "    </div>"
    Response.Write ""
    Response.Write "    <div class=""group2"">"
    Response.Write "      <p>"
    Response.Write "        <label for="""">Profession</label> <span>:</span>"
    Response.Write "        <span class=""answer"">" & repPro & "</span>"
    Response.Write "      </p>"
    Response.Write "      <p>"
    Response.Write "        <label for="""">Tel.</label> <span>:</span>"
    Response.Write "        <span class=""answer"">" & repTel & "</span>"
    Response.Write "      </p>"
    Response.Write "      <p>"
    Response.Write "        <label for="""">Email</label> <span>:</span>"
    Response.Write "        <span class=""answer"">" & repEmail & "</span>"
    Response.Write "      </p>"
    Response.Write "    </div>"
    Response.Write "  </div>"
End Sub

Sub Styles()
    Response.Write "    <style>"
    Response.Write "      * {"
    Response.Write "        font-family: Arial, sans-serif;"
    Response.Write "      }"
    Response.Write "      .container {"
    Response.Write "        width: 65vw;"
    Response.Write "        margin: 0;"
'    response.write "        text-align: left;"
    Response.Write "      }"
    Response.Write ""
    Response.Write "      h1 {"
    Response.Write "        text-align: cemter;"
    Response.Write "        margin-bottom: 10px;"
    Response.Write "        font-weight: bold;"
    Response.Write "        font-size: 16px;"
    Response.Write "      }"
    Response.Write ""
    Response.Write "      .form-header {"
    Response.Write "        text-align: right;"
    Response.Write "        font-size: 12px;"
    Response.Write "      }"
    Response.Write "      .section-title {"
    Response.Write "        font-weight: bold;"
    Response.Write "        background-color: #f0f0f0;"
    Response.Write "        padding: 5px;"
    Response.Write "        margin-bottom: 10px;"
    Response.Write "        text-align: left;"
    Response.Write "      }"
    Response.Write "      .section1,"
    Response.Write "      .dates {"
    Response.Write "        display: flex;"
    Response.Write "        justify-content: space-between;"
    Response.Write "        text-align: left;"
    Response.Write "      }"
    Response.Write "      .section-c {"
    Response.Write "        text-align: left;"
    Response.Write "      }"
    Response.Write "      textarea {"
    Response.Write "        min-height: 100px;"
    Response.Write "        min-width: 800px;"
    Response.Write "        font-family: Arial, sans-serif;"
    Response.Write "      }"
    Response.Write "      .mytable {"
    Response.Write "        width: 100%;"
    Response.Write "        border-collapse: collapse;"
    Response.Write "        margin-bottom: 10px;"
    Response.Write "      }"
    Response.Write "      .myth,"
    Response.Write "      .mytd {"
    Response.Write "        border: 1px solid #000;"
    Response.Write "        padding: 8px;"
    Response.Write "        text-align: left;"
    Response.Write "        font-family: Arial, sans-serif;"
    Response.Write "      }"
    Response.Write "      .myth {"
    Response.Write "        background-color: #f0f0f0;"
    Response.Write "      }"
    Response.Write "      label, span.answer {"
    Response.Write "        font-family: Arial, sans-serif;"
    Response.Write "      }"
    Response.Write "    </style>"
End Sub

Function getEMRResult(emrRequestID, emrDataID, CompID, column)
    Dim sql, rst, emrValue
    Set rst = Server.CreateObject("ADODB.Recordset")
    emrValue = ""
    sql = "SELECT * FROM emrresults WHERE emrrequestid ='" & emrRequestID & "'"
    sql = sql & " AND emrdataid ='" & emrDataID & "' AND emrcomponentid='" & CompID & "'"
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
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
