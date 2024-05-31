
'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim patientID, surname, gender, genderID, firstname, birthdate, insuranceNum, visitationID, determine, claimsCode
patientID = GetRecordField("PatientID")
visitationID = GetRecordField("VisitationID")

'response.write patientID
'response.write GetComboNameFld("patient", patientID, "SurName")

genderID = GetComboNameFld("Patient", patientID, "GenderID")
gender = GetComboName("Gender", genderID)
firstname = GetComboNameFld("Patient", patientID, "FirstName")
claimsCode = GetComboNameFld("Visitation", visitationID, "VisitInfo4")

specialistID = GetComboNameFld("Visitation", visitationID, "SpecialistID")
physician = GetComboNameFld("Specialist", specialistID, "SpecialistName")
'response.write VisitationID
'response.write GetComboNameFld("Visitation", "V1190104118", "PatientAge")

'determine = determineGender(visitationID)
'response.write determine

birthdate = GetComboNameFld("Patient", patientID, "BirthDate")
insuranceNum = GetComboNameFld("Visitation", visitationID, "InsuranceNo")
'Response.Write insuranceNum

CSS_Styles
NHISForm

Sub NHISForm()
  Response.Write " <div class='container' style='margin-top: 50px; margin-bottom: 50px'>"
  Response.Write "    <h2 class='heading'>NATIONAL HEALTH INSURANCE SCHEME</h2> "
  Response.Write "    <section id='form-info'> "
  Response.Write "      <div class='form-and-regulation' style='display: flex'> "
  Response.Write "        <img "
  Response.Write "          style='margin-right: 0.625rem' "
  Response.Write "          src='https://c8.alamy.com/comp/KYBMDA/national-health-insurance-scheme-ghana-nhis-logo-KYBMDA.jpg' '"
  Response.Write "          alt='National Health Insurance Logo' "
  Response.Write "          height='70' "
  Response.Write "         width='70' "
  Response.Write "        /> "
  Response.Write "        <div style='display: flex; flex-direction: column'> "
  Response.Write "          <p id='claim-form'>Claim Form</p> "
  Response.Write "          <p>(Regulation 62)</p> "
  Response.Write "        </div> "
  Response.Write "      </div> "

  Response.Write "      <div style='display: block'> "
  Response.Write "        <label for='form-no'>Form No.:</label> "
  Response.Write "        <input type='text' id='form-no' name='form-no' /><br /> "

  Response.Write "        <label for='hi-code' style='margin-right: 0.6785rem'>HI Code:</label> "
  Response.Write "        <input type='number' id='hi-code' name='hi-code' /> "
  Response.Write "      </div> "
  Response.Write "    </section> "

  Response.Write "    <section "
  Response.Write "      id='claims-info' "
  Response.Write "      style='display: flex; justify-content: space-between; align-items: center' '"
  Response.Write "    > "
  Response.Write "      <div style='display: flex; flex-direction: column'> "
  Response.Write "        <div style='display: block'> "
  Response.Write "          <label "
  Response.Write "            for='claim-code' "
  Response.Write "            class='claims-info-label' "
  Response.Write "            style='margin-right: 14px' "
  Response.Write "            >Claim Code:</label "
  Response.Write "          > "
  Response.Write "          <input type='text' id='claim-code' name='claim-code' /> "
  Response.Write "        </div> "

  Response.Write "        <div style='display: block'> "
  Response.Write "          <label for='scheme-code' class='claims-info-label' "
  Response.Write "            >Scheme Code:</label "
  Response.Write "          > "
  Response.Write "          <input type='text' id='scheme-code' name='scheme-code' /> "
  Response.Write "        </div> "

  Response.Write "        <div style='display: block'> "
  Response.Write "          <label "
  Response.Write "            for='referral-no' "
  Response.Write "            class='claims-info-label' "
  Response.Write "            style='margin-right: 14px' "
  Response.Write "            >Referral No:</label "
  Response.Write "          > "
  Response.Write "          <input type='text' id='referral-no' name='referral-no' /> "
  Response.Write "        </div> "
  Response.Write "      </div> "

  Response.Write "      <div> "
  Response.Write "        <label for='month-claim'>Month of Claim:</label> "
  Response.Write "        <input "
  Response.Write "          type='text' "
  Response.Write "          id='month-claim' "
  Response.Write "          name='month-claim' "
  Response.Write "          style='margin-right: 5px' "
  Response.Write "        /> "
  Response.Write "      </div> "
  Response.Write "      <div> "
  Response.Write "        <label for='date-claim'>Date of Claim: </label> "
  Response.Write "        <input type='text' id='date-claim' name='date-claim' /> "
  Response.Write "      </div> "
  Response.Write "    </section> "

  Response.Write "    <hr /> "

  Response.Write "    <h4>Client Information</h4> "
  Response.Write "    <div class='client-information'> "
  Response.Write "      <div class='surname-others-dob-hosp' style='display: flex; flex-direction: column' > "
  Response.Write "        <div style='display: flex'> "
  Response.Write "          <label for='surname' style='margin-right: 77.5px'>Surname:</label> "
  Response.Write "          <input type='text' id='surname' name='surname' value='" & GetComboNameFld("patient", patientID, "SurName") & "'/> "
  Response.Write "        </div> "

  Response.Write "        <div style='display: flex'> "
  Response.Write "          <label for='other-names' style='margin-right: 47px' >Other Names:</label> "
  Response.Write "          <input type='text' id='other-names' name='other-names' value= '" & firstname & "'/> "
  Response.Write "        </div> "

  Response.Write "        <div tyle='display: flex'> "
  Response.Write "          <label for='date-of-birth' style='margin-right: 43px'>Date of Birth:</label> "
  Response.Write "          <input type='text' id='date-of-birth' name='date-of-birth' value= '" & birthdate & "'/> "
  Response.Write "        </div> "

  Response.Write "        <div style='display: flex'> "
  Response.Write "          <label for='hospital-record' style='margin-right: 5px'>Hospital Record No:</label> "
  Response.Write "          <input type='text' id='hospital-record' name='hospital-record' value= '" & patientID & "'/> "
  Response.Write "        </div> "
  Response.Write "      </div> "

  Response.Write "      <div> "
  Response.Write "        <label for='age'>Age:</label> "
  Response.Write "        <input type='number' id='age' name='age' value= '" & GetComboNameFld("Patient", patientID, "Age") & "'/> "
  Response.Write "      </div> "

  Response.Write "      <div class='age-gender-member-claims'> "
  Response.Write "        <div> "
  Response.Write "          <label for='gender'>Gender</label> "
  Response.Write "          <input type='text' id='gender' name='gender' value= '" & gender & "'/> "
  Response.Write "        </div> "

  Response.Write "        <div> "
  Response.Write "          <label for='member-no'>Member No.:</label> "
  Response.Write "          <input type='number' id='member-no' name='member-no' value= '" & insuranceNum & "'/> "
  Response.Write "        </div> "

  Response.Write "        <div style='display: flex'> "
  Response.Write "          <label for='cliams-code'>Claims Check Code:</label> "
  Response.Write "          <input type='text' id='cliams-code' name='cliams-code' value= '" & claimsCode & "'/> "
  Response.Write "        </div> "
  Response.Write "      </div> "
  Response.Write "    </div> "

  Response.Write "    <hr /> "

  Response.Write "    <p><strong>Services Provided</strong><em>(to be filled by health care providers)</em></p> "

  Response.Write "    <section id='services-provided'> "
  Response.Write "      <div id='service-type'> "
  Response.Write "        <p><em>Type of Services</em></p> "
  Response.Write "(a)"
  Response.Write "        <input type='checkbox' id='outpatient' name='outpatient' /> "
  Response.Write "        <label for='outpatient'>Outpatient</label> "
  
  Response.Write "        <input type='checkbox' id='inpatient' name='inpatient' /> "
  Response.Write "        <label for='inpatient'>Inpatient</label> "
  
  Response.Write "        <input type='checkbox' id='pharmacy' name='pharmacy' /> "
  Response.Write "        <label for='pharmacy'>Pharmacy</label> "
  
  Response.Write "        <p style='margin-left: 100px; margin-bottom: -2px'>Investigation</p> "
  Response.Write "        <hr /> "
Response.Write "        (b) "
Response.Write "        <input "
Response.Write "          type='checkbox' "
Response.Write "          id='all-inclusive' "
Response.Write "          name='all-inclusive' "
Response.Write "          value='1' "
Response.Write "        /> "
Response.Write "        <label for='all-inclusive'>All Inclusive</label> "
Response.Write "        <input type='checkbox' id='unbundled' name='unbundled' /> "
Response.Write "        <label for='unbundled'>Unbundled</label> "
Response.Write "        <hr /> "
Response.Write "        <p><em>Outcome</em></p> "
Response.Write "        <input type='checkbox' id='discharged' name='discharged' /> "
Response.Write "        <label for='discharged'>Discharged</label> "
Response.Write "        <input type='checkbox' id='died' name='died' value='1' /> "
Response.Write "        <label for='died'>Died</label> "
Response.Write "        <input type='checkbox' id='transfer' name='transfer' /> "
Response.Write "        <label for='transfer'>Transfer</label> "
Response.Write "        <br /> "
Response.Write "        <input type='checkbox' id='absconded' name='absconded' /> "
Response.Write "        <label for='absconded' "
Response.Write "          >Absconded/Discharged against medical advice</label "
Response.Write "        > "
Response.Write "      </div> "
Response.Write "      <div id='provision-date'> "
Response.Write "        <h4>Date(s) of Services Provision</h4> "
Response.Write "        <label for='first-visit' style='margin-right: 40px' "
Response.Write "          >1st Visit/Admission</label "
Response.Write "        > "
Response.Write "        <input type='date' id='first-visit' name='first-visit' /><br /> "
Response.Write "        <label for='second-visit' style='margin-right: 35px' "
Response.Write "          >2nd Visit/Admission</label "
Response.Write "        > "
Response.Write "        <input type='date' id='second-visit' name='second-visit' /><br /> "
Response.Write "        <label for='third-visit' style='margin-right: 117px'>3rd Visit</label> "
Response.Write "        <input type='date' id='third-visit' name='third-visit' /><br /> "
Response.Write "        <label for='fourth-visit' style='margin-right: 38px' "
Response.Write "          >4th Visit/Admission</label "
Response.Write "        > "
Response.Write "        <input type='date' id='fourth-visit' name='fourth-visit' /><br /> "
Response.Write "        <label for='duration-spell'>Duration of Spell(days)</label> "
Response.Write "        <input "
Response.Write "          type='number' "
Response.Write "          id='duration-spell' "
Response.Write "          name='duration-spell' "
Response.Write "          style='width: 50px; margin-left: 77px' "
Response.Write "        /><br /> "
Response.Write "      </div> "
Response.Write "    </section> "
Response.Write "    <section id='attendance-type'> "
Response.Write "      <p><em>Type of Attendance</em></p> "
Response.Write "      <input type='checkbox' id='follow-up' name='follow-up' /> "
Response.Write "      <label for='follow-up'>Chronic Follow-up</label> "
Response.Write "      <input type='checkbox' id='emergency' name='emergency' /> "
Response.Write "      <label for='emergency'>Emergency</label> "
Response.Write "      <input type='checkbox' id='acute-episode' name='acute-episode' /> "
Response.Write "      <label for='acute-episode'>Acute Episode</label> "
Response.Write "      <input type='checkbox' id='anc' name='anc' /> "
Response.Write "      <label for='anc'>ANC</label> "
Response.Write "    </section> "
Response.Write "    <section id='physician-details'> "
Response.Write "      <label for='physician-name'>Physician/Clinician Name: </label> "
Response.Write "      <input "
Response.Write "        type='text' "
Response.Write "        id='physician-name' "
Response.Write "        name='physician-name' "
Response.Write "        value= '" & physician & "' "
Response.Write "      /><br /><br /> "
Response.Write "      <label for='physician-Id' style='margin-right: 26px' "
Response.Write "        >Physician/Clinician ID: "
Response.Write "      </label> "
Response.Write "      <input type='text' id='physician-Id' name='physician-Id' /><br /><br /> "
Response.Write "      <label for='speciality-code' style='margin-right: 65px' "
Response.Write "        >Specialilty Code: "
Response.Write "      </label> "
Response.Write "      <input "
Response.Write "        type='text' "
Response.Write "        id='speciality-code' "
Response.Write "        name='speciality-code' "
Response.Write "      /><br /><br /> "
Response.Write "    </section> "
Response.Write "    <table cellpadding='2' border='1' width='100%' cellspacing='0'> "
Response.Write "      <thead> "
Response.Write "        <tr> "
Response.Write "          <td colspan='4'> "
Response.Write "            <strong style='padding: 0px 2px'>Diagnosis</strong "
Response.Write "            ><em style='font-size: 12px' "
Response.Write "              >(to be filled by health care providers who have provided out and "
Response.Write "              in-patient services)</em "
Response.Write "            > "
Response.Write "          </td> "
Response.Write "        </tr> "
Response.Write "      </thead> "
Response.Write "      <thead> "
Response.Write "        <tr> "
Response.Write "          <th></th> "
Response.Write "          <th>Description</th> "
Response.Write "          <th>ICD-10</th> "
Response.Write "          <th>G-DRG</th> "
Response.Write "        </tr> "
Response.Write "      </thead> "
Response.Write "      <tbody> "
'response.write "        <tr> "
'response.write "          <td>1</td> "
'response.write "          <td>MALARIA</td> "
'response.write "          <td>B54</td> "
'response.write "          <td>OPDC06C</td> "
'response.write "        </tr> "
'response.write "        <tr> "
'response.write "          <td>2</td> "
'response.write "          <td>ANAEMIA (UNSPECIFIED)</td> "
'response.write "          <td>D64.9</td> "
'response.write "          <td>OPDC06C</td> "
'response.write "        </tr> "
displayDiseases visitationID
Response.Write "      </tbody> "
Response.Write "    </table> "
Response.Write "  </div> "
End Sub
 
Sub CSS_Styles()
  Response.Write "<style>"
  Response.Write "     * {"
  Response.Write "   box-sizing: border-box;"
  Response.Write " }"
  
  Response.Write " .container { "
  Response.Write "   width: 950px;"
  Response.Write "   margin: auto;"
  Response.Write "   font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande',"
  Response.Write "     'Lucida Sans', Arial, sans-serif;"
  Response.Write " }"
  
  Response.Write " .heading {"
  Response.Write "   text-align: center;"
  Response.Write "   margin-bottom: 0px;"
  Response.Write "   font-size: 1.4rem;"
  Response.Write " }"
 
  Response.Write " @media (max-width: 360px) {"
  Response.Write "   .heading {"
  Response.Write "     margin-bottom: 0px;"
  Response.Write "     font-size: 0.625rem;"
  Response.Write "   }"
  Response.Write " }"
  
  Response.Write " #form-info {"
  Response.Write "   display: flex;"
  Response.Write "   justify-content: space-between;"
  Response.Write "   align-items: center;"
  Response.Write "   margin-bottom: 0.625rem;"
  Response.Write " }"

  Response.Write " #claim-form {"
  Response.Write "   margin-bottom: -15px;"
  Response.Write " }"

  Response.Write " #claims-info input {"
  Response.Write "   width: 80px;"
  Response.Write " }"
 
  Response.Write " .client-information {"
  Response.Write "   display: flex;"
  Response.Write "   align-items: center;"
  Response.Write "   justify-content: space-between;"
  Response.Write " }"
  
  Response.Write " .surname-others-dob-hosp input {"
  Response.Write "   width: 120px;"
  Response.Write " }"

  Response.Write " .age-gender-member-claims {"
  Response.Write "   display: flex;"
  Response.Write "   flex-direction: column;"
  Response.Write "   text-align: right;"
  Response.Write " }"
 
  Response.Write " #services-provided {"
  Response.Write "   display: flex;"
  Response.Write "   margin-bottom: 10px;"
  Response.Write "   align-items: space-between;"
  Response.Write "   margin-bottom: 30px;"
  Response.Write "   gap: 20px;"
  Response.Write " }"
 
  Response.Write " #service-type {"
  Response.Write "   border: 1px solid black;"
  Response.Write "   width: 60%;"
  Response.Write "   margin-right: auto;"
  Response.Write "   padding: 5px 5px;"
  Response.Write "   text-align: left;"
  Response.Write " }"
 
  Response.Write " #provision-date {"
  Response.Write "   border: 1px solid black;"
  Response.Write "   width: 35%;"
  Response.Write "   padding: 15px 15px;"
  Response.Write " }"
 
  Response.Write " #attendance-type {"
  Response.Write "   border: 1px solid black;"
  Response.Write "   padding: 5px 5px;"
  Response.Write "   margin-bottom: 30px;"
  Response.Write "   text-align: left;"
  Response.Write " }"
    
  Response.Write " #physician-details {"
  Response.Write "   margin-top: 10px;"
  Response.Write "   border: 1px solid black;"
  Response.Write "   width: 100%;"
  Response.Write "   padding-top: 15px;"
  Response.Write "   padding-left: 15px;"
  Response.Write "   margin-bottom: 30px;"
  Response.Write "   text-align: left;"
  Response.Write " }"
  
  Response.Write " #physician-details input {"
  Response.Write "   border: none;"
  Response.Write " }"
 
  Response.Write "</style>"
End Sub
Sub displayDiseases(visitationID)
    Dim sql, rst, cnt, diseaseID, DiseaseCategoryID, DiseaseTypeID
    Set rst = CreateObject("ADODB.RecordSet")
    cnt = 0
    
    ' Assuming 'conn' is defined elsewhere in your code
    sql = "SELECT diseaseID, DiseaseCategoryID, DiseaseTypeID FROM Diagnosis WHERE VisitationID = '" & visitationID & "'"

    With rst
        .Open sql, conn, 3, 4
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                cnt = cnt + 1
                diseaseID = .fields("DiseaseID").Value
                DiseaseCategoryID = .fields("DiseaseCategoryID").Value
                DiseaseTypeID = .fields("DiseaseTypeID").Value

                Response.Write "<tr>"
                   Response.Write "<td align='center'>" & cnt & "</td>"
                   Response.Write "<td align='left'>" & GetComboName("Disease", diseaseID) & "</td>"
                   Response.Write "<td align='center'>" & DiseaseCategoryID & "</td>"
                   Response.Write "<td align='center'>" & DiseaseTypeID & "</td>"
                Response.Write "</tr>"

                .MoveNext
            Loop
        Else
            Response.Write "No records found"
        End If
        .Close
    End With
End Sub

Sub physicianDetails()
    
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>

