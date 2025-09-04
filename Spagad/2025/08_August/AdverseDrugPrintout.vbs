'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Styles
DrugReactionReport

Dim dob, gender, name, weight, tel, visitationID

Sub Styles()
    response.write "    <style>"
    response.write "      * {"
    response.write "        font-family: Arial, sans-serif;"
    response.write "      }"
    response.write "      .container {"
    response.write "        width: 65vw;"
    response.write "        margin: 0;"
'    response.write "        text-align: left;"
    response.write "      }"
    response.write ""
    response.write "      h1 {"
    response.write "        text-align: cemter;"
    response.write "        margin-bottom: 10px;"
    response.write "        font-weight: bold;"
    response.write "        font-size: 16px;"
    response.write "      }"
    response.write ""
    response.write "      .form-header {"
    response.write "        text-align: right;"
    response.write "        font-size: 12px;"
    response.write "      }"
    response.write "      .section-title {"
    response.write "        font-weight: bold;"
    response.write "        background-color: #f0f0f0;"
    response.write "        padding: 5px;"
    response.write "        margin-bottom: 10px;"
    response.write "        text-align: left;"
    response.write "      }"
    response.write "      .section1,"
    response.write "      .dates {"
    response.write "        display: flex;"
    response.write "        justify-content: space-between;"
    response.write "        text-align: left;"
    response.write "      }"
    response.write "      .section-c {"
    response.write "        text-align: left;"
    response.write "      }"
    response.write "      textarea {"
    response.write "        min-height: 100px;"
    response.write "        min-width: 800px;"
    response.write "        font-family: Arial, sans-serif;"
    response.write "      }"
    response.write "      .mytable {"
    response.write "        width: 100%;"
    response.write "        border-collapse: collapse;"
    response.write "        margin-bottom: 10px;"
    response.write "      }"
    response.write "      .myth,"
    response.write "      .mytd {"
    response.write "        border: 1px solid #000;"
    response.write "        padding: 8px;"
    response.write "        text-align: left;"
    response.write "        font-family: Arial, sans-serif;"
    response.write "      }"
    response.write "      .myth {"
    response.write "        background-color: #f0f0f0;"
    response.write "      }"
    response.write "      label, span.answer {"
    response.write "        font-family: Arial, sans-serif;"
    response.write "      }"
    response.write "    </style>"
End Sub

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
    sql = sql & "WHERE VisitationID = 'V1220908001'"

    With rst
        .open sql, conn, 3, 4
        
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
   
    sql2 = "SELECT EMRRequestID "
    sql2 = sql2 & "From EMRRequestItems WHERE EMRDataID = 'ADV001' "
    sql2 = sql2 & "AND VisitationID = 'V1220908001'"

    With rst1
        .open sql2, conn, 3, 4
        
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
   
    response.write "  <div class=""container"">"
    response.write "    <div class=""form-header"">"
    response.write "      FDA/ADR <br />"
    response.write "      FOOD AND DRUGS AUTHORITY<br />"
    response.write "      In Strict Confidence"
    response.write "    </div>"
    response.write "    <h1>"
    response.write "      ADVERSE REACTION REPORTING FORM <br />"
    response.write "      (Please complete all sections as much as possible)"
    response.write "    </h1>"
    response.write ""
    response.write "    <div class=""section-title"">(A) PATIENT DETAILS</div>"
    response.write "    <div class=""section1"">"
    response.write "      <div class=""group1"">"
    response.write "        <p>"
    response.write "          <label for="""">Age/Date Of Birth</label> <span>:</span>"
    response.write "          <span class=""answer"">" & dob & "</span>"
    response.write "        </p>"
    response.write "        <p>"
    response.write "          <label for="""">Gender</label> <span>:</span>"
    response.write "          <span class=""answer"">" & gender & "</span>"
    response.write "        </p>"
    response.write "        <p>"
    response.write "          <label for="""">If Female, Pregnant</label> <span>:</span>"
    response.write "          <span class=""answer"">" & isPregnant & "</span>"
    response.write "        </p>"
    response.write "        <p>"
    response.write "          <label for="""">Name/Folder Number</label> <span>:</span>"
    response.write "          <span class=""answer"">" & name & "</span>"
    response.write "        </p>"
    response.write "        <p>"
    response.write "          <label for="""">Hospital/Treatment Centre</label> <span>:</span>"
    response.write "          <span class=""answer"">FOCOS Orthopaedic Hospital</span>"
    response.write "        </p>"
    response.write "      </div>"
    response.write ""
    response.write "      <div class=""group2"">"
    response.write "        <p>"
    response.write "          <label for="""">Wt(kg)</label> <span>:</span>"
    response.write "          <span class=""answer"">" & weight & "</span>"
    response.write "        </p>"
    response.write "        <p>"
    response.write "          <label for="""">Age of Pregnancy</label> <span>:</span>"
    response.write "          <span class=""answer"">" & pregnancyAge & "</span>"
    response.write "        </p>"
    response.write "        <p>"
    response.write "          <label for="""">Telephone No</label> <span>:</span>"
    response.write "          <span class=""answer"">" & tel & "</span>"
    response.write "        </p>"
    response.write "      </div>"
    response.write "    </div>"
    response.write "    <div class=""section-title"">"
    response.write "      (B) DETAILS OF ADVERSE REACTION AND ANY TREATMENT GIVEN<br />"
    response.write "      (Attach a separate sheet and all relevant laboratory tests/data where"
    response.write "      necessary)"
    response.write "    </div>"
    response.write "  </div>"
    response.write "  <textarea name=""textarea"" rows=""5"" cols=""15"">" & description & "</textarea>"
    response.write "  <div class=""dates"" style=""margin-bottom: 15px"">"
    response.write "    <label>Date reaction started(dd/mm/yyy):</label> <span>" & reactionStart & "</span>"
    response.write "    <label>Date reaction stopped(dd/mm/yyy):</label> <span>" & reactionEnd & "</span>"
    response.write "  </div>"
    
    response.write " <div class='section-c'>"
    response.write "  <div class=""section-title"">(C) OUTCOME OF ADVERSE REACTION</div>"
    response.write "  <p><label>Outcome:</label> <span>" & outcome & "</span></p>"
    response.write "  <p>"
    response.write "    Did the adverse result in any untoward medical condition?  <span>" & untowardMedCon & "</span>"
    response.write "  </p>"
    response.write "  <p>If yes, Specify: <span>" & adverseRecSpec & "</span></p>"
    response.write "  <h4>SERIOUSNESS</h4>"
    response.write "  <p><span>" & seriousNess & "</span></p>"
    response.write "</div>"
    
    response.write "  <div class=""section-title"">"
    response.write "    (D) SUSPECTED PRODUCT(S) (Attach sample or product label if available)"
    response.write "  </div>"
    response.write "  <table class='mytable'>"
    response.write "    <tr>"
    response.write "      <th class='myth'>Brand name</th>"
    response.write "      <th class='myth'>Generic name</th>"
    response.write "      <th class='myth'>Batch name</th>"
    response.write "      <th class='myth'>Expiry date</th>"
    response.write "      <th class='myth'>Manufacturer</th>"
    response.write "    </tr>"
    response.write "    <tr>"
    response.write "      <td class='mytd'>" & brandName & "</td>"
    response.write "      <td class='mytd'>" & genericName & "</td>"
    response.write "      <td class='mytd'>" & batch & "</td>"
    response.write "      <td class='mytd'>" & expiryDate & "</td>"
    response.write "      <td class='mytd'>" & manufacturer & "</td>"
    response.write "    </tr>"
    response.write "    <tr>"
    response.write "      <th class='myth' colspan=""2"">Reasons for use (Indication):</th>"
    response.write "      <th class='myth'>Dosage regimen:</th>"
    response.write "      <th class='myth'>No. of days given</th>"
    response.write "      <th class='myth'>Route of administration:</th>"
    response.write "    </tr>"
    response.write "    <tr>"
    response.write "      <td colspan=""2"" class='mytd'>" & reasons & "</td>"
    response.write "      <td class='mytd'>" & dosage & "</td>"
    response.write "      <td class='mytd'>" & noOfDays & "</td>"
    response.write "      <td class='mytd'>" & route & "</td>"
    response.write "    </tr>"
    response.write "    <tr>"
    response.write "      <td class='mytd' colspan=""5"">"
    response.write "        Date started (dd/mm/yyyy): <span>" & dateStarted & "</span> &nbsp; &nbsp; &nbsp;"
    response.write "        &nbsp; Date stopped (dd/mm/yyyy): <span>" & dateEnded & "</span>"
    response.write "        <br />"
    response.write "        <em"
    response.write "          >Did the adverse reaction subside when the drug was"
    response.write "          stopped(de-challenge)? <span>" & isSubsided & "</span></em"
    response.write "        >"
    response.write "      </td>"
    response.write "    </tr>"
    response.write "    <tr>"
    response.write "      <td class='mytd' colspan=""3"">Was the product prescribed? <span>" & isPrescribed & "</span></td>"
    response.write "      <td class='mytd' colspan=""2"">Source of drug: <span>" & source & "</span></td>"
    response.write "    </tr>"
    response.write "    <tr>"
    response.write "      <td colspan=""5"" class='mytd'>"
    response.write "        Was the product re-used after detection of adverse reaction"
    response.write "        (re-challenge)? <span class=""answer"">" & isReused & "</span>"
    response.write "      </td>"
    response.write "    </tr>"
    response.write "    <tr>"
    response.write "      <td colspan=""5"" class='mytd'>"
    response.write "        Did adverse reaction re-appear upon re-use?"
    response.write "        <span class=""answer"">" & isReappear & "</span>"
    response.write "      </td>"
    response.write "    </tr>"
    response.write "  </table>"
    response.write ""
    response.write "  <div class=""section-title"">"
    response.write "    (E) CONCOMITANT DRUGS: INCLUDING COMPLEMENTARY MEDICINES, ADMINISTERED"
    response.write "    <br />"
    response.write "    AT THE SAME TIME AND/OR 3 MONTHS BEFORE (Attach a separate sheet when"
    response.write "    necessary)"
    response.write "  </div>"
    response.write ""
    response.write "  <table class='mytable'>"
    response.write "    <tr>"
    response.write "      <th class='myth'>Name of Drug</th>"
    response.write "      <th class='myth'>Daily dose</th>"
    response.write "      <th class='myth'>Date started</th>"
    response.write "      <th class='myth'>Date stopped/Ongoing</th>"
    response.write "      <th class='myth'>Reason(s) for use</th>"
    response.write "    </tr>"
    response.write "    <tr>"
    response.write "      <td class='mytd'>Para</td>"
    response.write "      <td class='mytd'>Para</td>"
    response.write "      <td class='mytd'>Para</td>"
    response.write "      <td class='mytd'>Para</td>"
    response.write "      <td class='mytd'>Para</td>"
    response.write "    </tr>"
    response.write "  </table>"
    response.write ""
    response.write "  <div class=""section-title"">(F) REPORTER DETAILS</div>"
    response.write ""
    response.write "  <div class=""section1"">"
    response.write "    <div class=""group1"">"
    response.write "      <p>"
    response.write "        <label for="""">Name of Reporter</label> <span>:</span>"
    response.write "        <span class=""answer"">" & repName & "</span>"
    response.write "      </p>"
    response.write "      <p>"
    response.write "        <label for="""">Institution's Address</label> <span>:</span>"
    response.write "        <span class=""answer"">" & repAddress & "</span>"
    response.write "      </p>"
    response.write "      <p>"
    response.write "        <label for="""">Signature</label> <span>:</span>"
    response.write "        <span class=""answer"">_____________</span>"
    response.write "      </p>"
    response.write "      <p>"
    response.write "        <label for="""">Date</label> <span>:</span>"
    response.write "        <span class=""answer"">" & repDate & "</span>"
    response.write "      </p>"
    response.write "    </div>"
    response.write ""
    response.write "    <div class=""group2"">"
    response.write "      <p>"
    response.write "        <label for="""">Profession</label> <span>:</span>"
    response.write "        <span class=""answer"">" & repPro & "</span>"
    response.write "      </p>"
    response.write "      <p>"
    response.write "        <label for="""">Tel.</label> <span>:</span>"
    response.write "        <span class=""answer"">" & repTel & "</span>"
    response.write "      </p>"
    response.write "      <p>"
    response.write "        <label for="""">Email</label> <span>:</span>"
    response.write "        <span class=""answer"">" & repEmail & "</span>"
    response.write "      </p>"
    response.write "    </div>"
    response.write "  </div>"
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


