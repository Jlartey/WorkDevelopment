'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim mth, periodStart, periodEnd, rptType, agegroup, arrAG, arrGender, arrTot(20, 20), pos
'addJS
dt = Trim(Request.QueryString("PrintFilter0"))
arPeriod = getDatePeriodFromDelim(dt)
periodStart = arPeriod(0)
periodEnd = arPeriod(1)
rptName = Trim(Request.QueryString("PrintLayoutName"))
InitVariables
DisplayHeader
ProcessCode
addJS

Sub InitVariables()
  mth = ""
  rptType = Trim(Request.QueryString("ReportType"))
  agegroup = "< 28 days||0||0.077**1- 11 mths||0.078||0.99"
  agegroup = agegroup & "**1-4 yrs||1||5**5- 9 yrs||5||10**10- 14 yrs||10||15**15- 17 yrs||15||18"
  agegroup = agegroup & "**18- 19 yrs||18||20**20- 34 yrs||20||35**35- 49 yrs||35||50**50- 59 yrs||50||60"
  agegroup = agegroup & "**60- 69 yrs||60||70**70+ yrs||70||200"
  arrAG = Split(agegroup, "**")
  arrGender = Split("GEN01**GEN02", "**")
End Sub


Sub GetGroupHeading(hdr, lstMorbidity)
  Dim ul, minAge, maxAge, tot, whcls
  ul = UBound(arrAG)
  repo = Request.QueryString("ReportType")
  response.write "<style>table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 90vw; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{background-color: #3C8F6D;color:#fff;font-weight:700;} table#myTable thead .h_title{ background-color: #3C8F6D;color:#fff; } table#myTable thead .h_names{ font-size: 14px;} table#myTable small.blk{display:block;margin:0;padding:0;}</style>"
  response.write "<table id='myTable'> <thead><tr class='h_res'><th colspan='26'>Results are Ready...</th></tr><tr class='h_title'><th colspan='26'>" & repo & " Diseases</th></tr>"
  response.write "<tr class='h_names'><th rowspan='2'>DISEASES(NEW CASES ONLY)</th>"
  For Each GenderID In arrGender
    response.write "<th colspan='12'>" & GetComboName("Gender", GenderID) & "</th>"
  Next
  response.write "<th rowspan='2'>TOTAL</th></tr>"
  response.flush

  response.write "<tr>"
  For Each GenderID In arrGender
    For Each lstAG In arrAG
      arAge = Split(lstAG, "||")
      response.write "<th>" & arAge(0) & "</th>"
    Next
  Next
  response.write "</tr></thead>"

  response.write "<tbody><tr>"
  arrMorbidity2 = Split(lstMorbidity, "**")
  For Each morbidity In arrMorbidity2
    arrMorbidity = Split(morbidity, "||")
    tot = 0
    whcls = ""
    If UBound(arrMorbidity) >= 1 Then
      If Len(Trim(arrMorbidity(1))) > 3 Then
        response.write "<td>" & arrMorbidity(0) & "</td>"
        whcls = arrMorbidity(1)
        response.flush

        For Each GenderID In arrGender
          For i = 0 To ul
            arAge = Split(arrAG(i), "||")
            minAge = arAge(1)
            maxAge = arAge(2)
            v = GetDiagnosisCount(GenderID, minAge, maxAge, periodStart, periodEnd, whcls)
            tot = tot + v
            response.write "<td>" & v & "</td>"
            response.flush
          Next
        Next
        response.write "<td>" & tot & "</td></tr>"
      End If
    End If
  Next
  response.write "</tbody></table>"
End Sub

Sub ProcessCode()
  response.write "<tr>"
  response.write "<td align=""center"" width=""100%"">"
  lstMorbidity = ""
  Select Case UCase(rptType)
  Case UCase("Immunizable")
    description = "Communicable Immunizable"
    ' SetPageMessages description
    lstMorbidity = lstMorbidity & "**AFP (Polio)||And (DiseaseName Like '%FETOPROTEIN%' Or DiseaseName like '%POLIO%') "
    lstMorbidity = lstMorbidity & "**Meningitis ||And (DiseaseName Like '%Meningitis%') "
    lstMorbidity = lstMorbidity & "**Tetanus/Neo-Natal Tetanus ||And (DiseaseName Like '%Tetanus%' Or DiseaseID like 'A33%') "
    lstMorbidity = lstMorbidity & "**Pertussis (Whooping Cough) ||And (DiseaseName Like '%Pertussis%' Or DiseaseName like '%Whooping%Cough%') "
    lstMorbidity = lstMorbidity & "**Diphteria ||And (DiseaseName Like '%Diphteria%' Or DiseaseName like '%Diphtheria%') "
    lstMorbidity = lstMorbidity & "**Measles ||And (DiseaseName Like '%Measles%') "
    lstMorbidity = lstMorbidity & "**Yellow Fever (YF) ||And (DiseaseName Like '%Yellow%Fever%') "
    lstMorbidity = lstMorbidity & "**Tuberculosis (TB) ||And (DiseaseName Like '%Tuberculosis%') "
    lstMorbidity = GetReportCategories(description)
    ' response.write lstMorbidity
    GetGroupHeading description, lstMorbidity
  Case UCase("NonImmunizable")
    description = "Communicable non-immunizable"
    lstMorbidity = ""
    '' Malaria specifics will be put here later
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    lstMorbidity = lstMorbidity & "**<b class='red'>Uncomplicated Malaria Suspected </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Uncomplicated Malaria Suspected Tested </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Uncomplicated Malaria Tested Positive </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Uncomplicated Malaria not Tested but Treated as Malaria </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Uncomplicated Malaria Cases Tested Negative but Treated as Malaria</b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Uncomplicated Malaria in Pregnanacy Suspected </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Uncomplicated Malaria in Pregnanacy Suspected Tested </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Uncomplicated Malaria in Pregnanacy Tested Positive </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Uncomplicated Malaria in Pregnanacy not Tested but Treated as Malaria </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Uncomplicated Malaria in Pregnanacy Tested Negative but Treated as Malaria </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Severe Malaria (Lab-Confirmed) </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Severe Malaria (Non-Lab-Confirmed) </b> ||And 11=22 "

    lstMorbidity = lstMorbidity & "**Typhoid Fever ||And (DiseaseName Like '%Typhoid%') "
    lstMorbidity = lstMorbidity & "**Suspected Cholera ||And DiseaseName Like '%Cholera%' And DiseaseName like '%Diarh%ea%' "
    lstMorbidity = lstMorbidity & "** <b class='red'>Diarrhoea Diseases</b> ||And (DiseaseName Like '%Diarrh%ea%' Or DiseaseName like '%Diarh%ea%') And Not DiseaseName Like '%Cholera%' "
    lstMorbidity = lstMorbidity & "**Schistosomiasis (Bilhazia) ||And (DiseaseName Like '%Schistosomiasis%' Or DiseaseName like '%Bilhazia%') "
    lstMorbidity = lstMorbidity & "**Suspected Guinea Worm ||And (DiseaseName Like '%Guinea%Worm%') "
    lstMorbidity = lstMorbidity & "**Onchocerciasis ||And (DiseaseName Like '%Onchocerciasis%' Or DiseaseName like '%River%Blind%') "
    lstMorbidity = lstMorbidity & "**Leprosy ||And (DiseaseName Like '%Leprosy%') "
    lstMorbidity = lstMorbidity & "**<b class='red'>HIV/AIDS Related Conditions</b> ||And (DiseaseName Like '%HIV%' Or DiseaseName like '%AIDS%') "
    lstMorbidity = lstMorbidity & "**Mumps ||And (DiseaseName Like '%Mumps%') "
    lstMorbidity = lstMorbidity & "**Intestinal Worms ||And (DiseaseName Like '%helminthiasis%' Or DiseaseName like '%hookworm%' Or DiseaseName like '%tapeworm%') "
    lstMorbidity = lstMorbidity & "**Chicken Pox ||And (DiseaseName Like '%Chicken%Pox%' Or DiseaseName like '%Varicella%') "
    lstMorbidity = lstMorbidity & "**Upper Respiratory Tract Infection ||And (DiseaseName Like '%Upper%Respiratory%Tract%Infection%') "
    lstMorbidity = lstMorbidity & "**Pneumonia ||And (DiseaseName Like '%Pneumonia%') "
    lstMorbidity = lstMorbidity & "**Septiceamia/Sepsis ||And (DiseaseName Like '%sepsis%' Or DiseaseName like '%Septic%emia%') "
    lstMorbidity = GetReportCategories(description)
    GetGroupHeading description, lstMorbidity
  Case UCase("NonCommunicable")
    description = "Non-Communicable Diseases"
    lstMorbidity = ""
    lstMorbidity = lstMorbidity & "**Malnutrition ||And (DiseaseName Like '%Malnutrition%' Or DiseaseName like '%Kwashiorkor%' Or DiseaseName like '%Marasm%') "
    lstMorbidity = lstMorbidity & "**Obesity ||And (DiseaseName Like '%Obesity%') "
    lstMorbidity = lstMorbidity & "**Anaemia ||And (DiseaseName Like '%An%emia%') AND Not (DiseaseName like '%pregnancy%') "
    lstMorbidity = lstMorbidity & "**Other Nutritional Diseases ||And (DiseaseName Like '%Nutritional%Diseases%') "
    lstMorbidity = lstMorbidity & "**Hypertension ||And (DiseaseName Like '%Hypertension%' Or DiseaseName Like '%HPT%') "
    lstMorbidity = lstMorbidity & "**Cardiac Diseases ||And (DiseaseName Like '%Cardiac%') "
    lstMorbidity = lstMorbidity & "**Stroke ||And (DiseaseName Like '%Stroke%' Or DiseaseName like '%CVA%' Or (DiseaseName like '%cerebral&infarction%' and Not DiseaseName like '%without%')) "
    lstMorbidity = lstMorbidity & "**Diabetes ||And (DiseaseName Like '%Diabetes%') "
    lstMorbidity = lstMorbidity & "**Rheumatism /Other Joint Pains /Arthritis ||And (DiseaseName Like '%Rheumati%') "
    lstMorbidity = lstMorbidity & "**Sickle Cell Diseases ||And (DiseaseName Like '%Sickle%cell%') "
    lstMorbidity = lstMorbidity & "**Asthma ||And (DiseaseName Like '%Asthma%') "
    lstMorbidity = lstMorbidity & "**<b class='red'>COPD</b> ||And 1=2 "
    lstMorbidity = lstMorbidity & "**Breast Cancer ||And (DiseaseName Like '%Breast%Cancer%') "
    lstMorbidity = lstMorbidity & "**Cervical Cancer ||And (DiseaseName Like '%Cervi%Cancer%' Or DiseaseName like '%Cancer%Cervi%') "
    lstMorbidity = lstMorbidity & "**Lymphoma ||And (DiseaseName Like '%Lymphoma%') "
    lstMorbidity = lstMorbidity & "**<b class='red'>Prostate </b> ||And ((DiseaseName Like '%Prostate%' And DiseaseName like '%Cancer%') Or (DiseaseName like '%mali%neoplasm%')) "
    lstMorbidity = lstMorbidity & "**Liver Disease ||And (DiseaseName Like '%Liver%Disease%') "
    lstMorbidity = lstMorbidity & "**<b class='red'>All Other Cancers </b> ||And 1=2 "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    lstMorbidity = GetReportCategories(description)
    GetGroupHeading description, lstMorbidity
  Case UCase("Mental")
    description = "Mental Health Conditions"
    lstMorbidity = ""
    lstMorbidity = lstMorbidity & "**<b class='red'>Schizophrenia </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Acute Psychotic Disorder </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Mono Symptoms Delusion </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Depression </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Substance Abuse </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Epilepsy </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Autism </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Mental Retardation </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Attention Deficit </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Hyperactivity Disorder </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Conversion Disorder </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Post Traumatic Stress Syndrome </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Generalized Anxiety </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Other Anxiety Disorders </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Neurosis </b> ||And 11=22 "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    lstMorbidity = GetReportCategories(description)
    GetGroupHeading description, lstMorbidity
  Case UCase("Specialized")
    description = "Specialized Conditions"
    lstMorbidity = ""
    lstMorbidity = lstMorbidity & "**<b class='red'>Acute Eye Infection </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class=''>Cataract </b> ||And (DiseaseName Like '%Cataract%') "
    lstMorbidity = lstMorbidity & "**<b class=''>Trachoma </b> ||And (DiseaseName Like '%Trachoma%') "
    lstMorbidity = lstMorbidity & "**<b class='red'>Otitis Media </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Other Acute Ear Infection </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class=''>Dental Carriers </b> ||And (DiseaseName Like '%Dental%Carriers%') "
    lstMorbidity = lstMorbidity & "**<b class='red'>Dental Swellings </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Traumatic Conditions (Oral and Maxillofacial Region) </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Periodental Diseases </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class=''>Cerebral Palsy </b> ||And (DiseaseName Like '%Cerebral%Palsy%') "
    lstMorbidity = lstMorbidity & "**<b class='red'>Liver Diseases </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class=''>Urinary Tract Infection </b> ||And (DiseaseName Like '%Urinary%Tract%Infection%') "
    lstMorbidity = lstMorbidity & "**<b class='red'>Skin Diseases </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class=''>Ulcer <em><sub>(ALL)</sub></em> </b> ||And (DiseaseName Like '%Ulcer%' And Not (DiseaseName Like '%genita%' Or DiseaseName Like '%vagina%' Or DiseaseName Like '%penis%')) "
    ' lstMorbidity = lstMorbidity & "**<b class=''>Ulcer <em><sub>(ALL)</sub></em> </b> ||And (DiseaseName Like '%Ulcer%' And Not (DiseaseName Like '%genita%' Or DiseaseName Like '%vagina%' Or DiseaseName Like '%penis%')) "
    lstMorbidity = lstMorbidity & "**<b class='red'>Kidney Related Diseases </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'>Other Oral Conditions </b> ||And 11=22 "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    lstMorbidity = GetReportCategories(description)
    GetGroupHeading description, lstMorbidity
  Case UCase("Obstetrics")
    description = "Obs & Gynaecological Conditions"
    lstMorbidity = ""
    ' lstMorbidity = lstMorbidity & "**<b class='red'>Gynaecological Conditions </b> ||And 11=22 "
    ' lstMorbidity = lstMorbidity & "**<b class='red'>Pregnanacy Related Complications </b> ||And 11=22 "
    ' lstMorbidity = lstMorbidity & "**<b class='red'>Anaemia in Pregnanacy </b> ||And (DiseaseName Like '%An%emia%' And DiseaseName Like '%Pregnanacy%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    lstMorbidity = GetReportCategories(description)
    GetGroupHeading description, lstMorbidity
  Case UCase("Reproductive")
    description = "Reproductive Tract Diseases"
    lstMorbidity = ""
    lstMorbidity = lstMorbidity & "**<b class=''>Gonorrhoea </b> ||And (DiseaseName Like '%Gonorrh%a%') "
    ' lstMorbidity = lstMorbidity & "**<b class='red'>Genital Ulcer </b> ||And (DiseaseName Like '%Genital%') "
    lstMorbidity = lstMorbidity & "**<b class='red'> </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'> </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'> </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'> </b> ||And 11=22 "
    lstMorbidity = lstMorbidity & "**<b class='red'> </b> ||And 11=22 "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    lstMorbidity = GetReportCategories(description)
    GetGroupHeading description, lstMorbidity
  Case UCase("Injury")
    description = "Injuries and Others"
    lstMorbidity = ""
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%' Or DiseaseName like '%%') "
    ' lstMorbidity = lstMorbidity & "** ||And (DiseaseName Like '%%' Or DiseaseName like '%%') "
    lstMorbidity = GetReportCategories(description)
    GetGroupHeading description, lstMorbidity
  Case UCase("All")
    description = "All"
    lstMorbidity = ""
    lstMorbidity = GetReportCategories(description)
    GetGroupHeading description, lstMorbidity
  Case UCase("Review")
    description = "Review"
    lstMorbidity = GetReportCategories(description)
    ' response.write lstMorbidity
    GetGroupHeading description, lstMorbidity
  Case UCase("AttRef")
  ' case else
  End Select
    response.write "</td>"
    response.write "</tr>"
End Sub

Sub DisplayHeader()
  Dim str
  str = ""
  str = str & "<div class='top-cont'><select id='inpReportType' name='inpReportType'>"
  str = str & "<option disabled selected hidden value='-'>Select Disease Category</option>"
  lstRpt = ""
  lstRpt = lstRpt & "**All||All"
  lstRpt = lstRpt & "**Immunizable||Communicable Immunizable"
  lstRpt = lstRpt & "**NonImmunizable||Communicable Non-Immunizable"
  lstRpt = lstRpt & "**NonCommunicable||Non-Communicable Diseases"
  lstRpt = lstRpt & "**Mental||Mental Health Conditions"
  lstRpt = lstRpt & "**Specialized||Specialized Conditions"
  lstRpt = lstRpt & "**Obstetrics||Obstetrics and Gynaecological Conditions"
  lstRpt = lstRpt & "**Reproductive||Reproductive Tract Diseases"
  lstRpt = lstRpt & "**Injury||Injuries and Others"
  lstRpt = lstRpt & "**AttRef||Re-Attendance and Referrals"
  ' now for each of the above ** loop them inside select option
  arReportType = Split(lstRpt, "**")
  For Each rpt In arReportType
    arRpt = Split(rpt, "||")
    If UBound(arRpt) >= 1 Then
      str = str & "<option value='" & arRpt(0) & "' "
      If UCase(arRpt(0)) = UCase(rptType) Then
        str = str & " selected='selected' "
      End If
      str = str & ">" & arRpt(1) & "</option>"
    End If
  Next
  str = str & "</select> &emsp; "
  str = str & "<input type='button' id='cmdReport' onclick='cmdReportOnClick()' value='Process Report'>"
  response.write str
  response.flush
End Sub

Function GetDiagnosisCount(GenderID, minAgeLmt, maxAgeLmt, startDate, endDate, whcls)
  Dim rst, sql, ot
  ot = 0
  Set rst = server.CreateObject("ADODB.RecordSet")
  sql = "select count(distinct dg.PatientID) as 'Cnt' "
  sql = sql & " from Diagnosis dg, Visitation v Where v.VisitationID=dg.VisitationID "
  sql = sql & " and v.visitdate>='" & startDate & "' and v.visitdate<= '" & endDate & "' "
  sql = sql & " and v.PatientAge >= " & minAgeLmt & " and v.PatientAge < " & maxAgeLmt & " "
  sql = sql & " and v.GenderID='" & GenderID & "' "
  sql = sql & " and (v.MedicalServiceID='M001' Or v.MedicalServiceID='M002') "
  sql = sql & " and v.BranchID = '" & brnch & "' "
  sql = sql & whcls

  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        ot = .fields("Cnt")
        .MoveNext
      Loop
    End If
  End With
  GetDiagnosisCount = ot
End Function

Sub addJS()
  Dim js
  js = "<script type=""text/javascript"">" & vbNewLine
  js = js & "function cmdReportOnClick() { " & vbNewLine
  js = js & " var url, ele, dtStart, dtEnd;" & vbNewLine
  js = js & " ele = document.getElementById('cmdReport'); if (ele) { ele.innerText = ""Processing Report...""; } " & vbNewLine
  js = js & " ele = document.getElementById('icoReport'); if (ele) { ele.innerHTML = '<div class=""spinner-border"" role=""status""><span class=""visually-hidden"">Loading...</span></div>'; } " & vbNewLine
  js = js & " dtStart = '" & periodStart & "';" & vbNewLine
  js = js & " dtEnd = '" & periodEnd & "';" & vbNewLine
  js = js & " url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=dhimsOPDMobidityReport&PositionForTableName=WorkingDay&WorkingDayID=DAY20180222';"
  js = js & " url = url + '&PrintFilter0=' + dtStart + '||' + dtEnd;" & vbNewLine
  js = js & " url = url + '&ReportType=' + GetEleVal('inpReportType');" & vbNewLine
  js = js & "/* alert(url); */" & vbNewLine
  js = js & " window.location.href=processurl(url); " & vbNewLine
  js = js & "} " & vbNewLine
  js = js & "</script>  " & vbNewLine
  response.write js
End Sub

Function GetReportCategories(description)
  Dim sql, rst, ot
  Set rst = CreateObject("ADODB.Recordset")
  sql = "select testvar1bid, testvar1bname from testvar1b "
  sql = sql & "where testvar1aid = '" & rptName & "' "
  If UCase(Trim(description)) <> UCase("ALL") Then
    sql = sql & " and description = '" & description & "'; "
  End If
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        cat_id = Trim(.fields("testvar1bid"))
        tbname = Trim(.fields("testvar1bname"))
        cat_diseases = GetCategoriesDisease(cat_id)
        If Trim(cat_diseases) = "" Then
          new_temp = "1=0"
        Else
          new_temp = "dg.DiseaseID IN ( " & cat_diseases & " ) "
        End If
        ot = ot & "**" & tbname & "||And " & new_temp
        .MoveNext
      Loop
      .Close
    End If
  End With
  Set rst = Nothing
  GetReportCategories = ot
End Function

Function GetCategoriesDisease(cat_id)
  Dim sql, rst, ot, ut
  ut = ""
  sql = "select performvar15name from performvar15 where keyprefix = '" & cat_id & "'"
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        disease_id = Trim(.fields("performvar15name"))
        ut = ut & "'" & disease_id & "', "
        .MoveNext
      Loop
      .Close
      ut = Left(ut, (Len(ut) - Len(", ")))
    End If
  End With
  Set rst = Nothing
  GetCategoriesDisease = ut
End Function

Function getDatePeriodFromDelim(strDelimPeriod)
    Dim arPeriod, periodStart, periodEnd
    Dim arOut(1)
    arPeriod = Split(strDelimPeriod, "||")

    If UBound(arPeriod) >= 0 Then
        periodStart = arPeriod(0)
    End If

    If UBound(arPeriod) >= 1 Then
        periodEnd = arPeriod(1)
    End If

    periodStart = makeDatePeriod(Trim(periodStart), periodEnd, "0:00:00")
    periodEnd = makeDatePeriod(Trim(periodEnd), periodStart, "23:59:59")

    arOut(0) = periodStart
    arOut(1) = periodEnd

    getDatePeriodFromDelim = arOut
End Function

Function makeDatePeriod(strDateStart, defaultDate, strTime)
    If IsDate(strDateStart) Then
            makeDatePeriod = FormatDate(strDateStart) & " " & Trim(strTime)
    Else
        If IsDate(defaultDate) Then
                makeDatePeriod = FormatDate(defaultDate) & " " & Trim(strTime)
        Else
                makeDatePeriod = FormatDate(Now()) & " " & Trim(strTime)
        End If
    End If
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
