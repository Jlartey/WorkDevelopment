'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

response.write "<style> table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 80vw;; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{ background-color: #3C8F6D; color:#fff } table#myTable thead .h_title{ background-color: blanchedalmond; } table#myTable thead .h_names{ font-size: 14px;} table#myTable tbody td{text-align:center;} table#myTable .last{background-color: blanchedalmond;color:#000;font-weight:700;text-align:center;}  </style>"

dt = Request.QueryString("PrintFilter0")
arDt = getDatePeriodFromDelim(dt)
periodStart = arDt(0)
periodEnd = arDt(1)

startDate = FormatWorkingDay(periodStart)
endDate = FormatWorkingDay(periodEnd)

Dim agegroup, gender_male, gender_female, male_admissions, female_admissions, female_deaths, male_deaths
Dim admission_total, death_total, overall_total, female_agegroup_admission_total, female_agegroup_death_total
Dim male_agegroup_admission_total, male_agegroup_death_total
Dim total_row_ele, total_admission_total, total_death_total, grand_grand_total
total_row_ele = ""
gender_female = "gen02"
gender_male = "gen01"

workingMonth = Request.QueryString("PrintFilter1")

hdr = " IPD Statement <br> "
hdr = hdr & " Period: " & GetComboName("WorkingDay", startDate) & " To " & GetComboName("WorkingDay", endDate)

response.write "<table id='myTable'> "
response.write "<thead><tr class='h_res'><th colspan='8'>" & hdr & "</th></tr></thead>"
response.write "    <tr> "
response.write "        <th rowspan=""2"">Cohort</th> "
response.write "        <th colspan=""2"">Admission</th> "
response.write "        <th rowspan=""2"">Admission Total</th> "
response.write "        <th colspan=""2"">Death</th> "
response.write "        <th rowspan=""2"">Death Total</th> "
response.write "        <th rowspan=""2"">Grand Total</th> "
response.write "    </tr>"
response.write "    <tr> "
response.write "        <td>Male</td> "
response.write "        <td>Female</td> "
response.write "        <td>Male</td> "
response.write "        <td>Female</td> "
response.write "    </tr> "

agegroup = "0-28 days||0||0.077**29 days - less than a year||0.077||0.99"
agegroup = agegroup & "**1 - 4 Years||1||5**5 - 9 Years||5||10**10 - 14 Years||10||15**15 - 17 Years||15||18"
agegroup = agegroup & "**18 - 19 Years||18||20**20 - 34 Years||20||35**35 - 49 Years||35||50**50 - 59 Years||50||60"
agegroup = agegroup & "**60 - 69 Years||60||70**70 Yrs & Above||70||200"

arrAG = Split(agegroup, "**")

female_agegroup_admission_total = 0
male_agegroup_admission_total = 0

female_agegroup_death_total = 0
male_agegroup_death_total = 0

total_admission_total = 0
total_death_total = 0

For Each element In arrAG

  female_admissions = 0
  male_admissions = 0

  female_deaths = 0
  male_deaths = 0

  admission_total = 0
  death_total = 0

  Temp = Split(element, "||")
  age_group = Temp(0)
  lower_age_limit = Temp(1)
  upper_age_limit = Temp(2)


  additional_filter = ""
  additional_filter = additional_filter & "v.patientage >= '" & lower_age_limit & "' "
  additional_filter = additional_filter & "and v.patientage < '" & upper_age_limit & "' "

  female_admissions = VisitationWithAdmission(gender_female, additional_filter)
  male_admissions = VisitationWithAdmission(gender_male, additional_filter)

  female_agegroup_admission_total = female_agegroup_admission_total + female_admissions
  male_agegroup_admission_total = male_agegroup_admission_total + male_admissions

  female_deaths = DeathCount(gender_female, additional_filter)
  male_deaths = DeathCount(gender_male, additional_filter)

  female_agegroup_death_total = female_agegroup_death_total + female_deaths
  male_agegroup_death_total = male_agegroup_death_total + male_deaths

  admission_total = female_admissions + male_admissions
  death_total = female_deaths + male_deaths

  total_admission_total = total_admission_total + admission_total
  total_death_total = total_death_total + death_total

  overall_total = admission_total + death_total

  htmlResponse = ""
  htmlResponse = htmlResponse & "<tr>"
  htmlResponse = htmlResponse & "<td>" & age_group & "</td>"
  htmlResponse = htmlResponse & "<td>" & male_admissions & "</td>"
  htmlResponse = htmlResponse & "<td>" & female_admissions & "</td>"
  htmlResponse = htmlResponse & "<td>" & admission_total & "</td>"
  htmlResponse = htmlResponse & "<td>" & male_deaths & "</td>"
  htmlResponse = htmlResponse & "<td>" & female_deaths & "</td>"
  htmlResponse = htmlResponse & "<td>" & death_total & "</td>"
  htmlResponse = htmlResponse & "<td>" & overall_total & "</td>"
  htmlResponse = htmlResponse & "</tr>"

  response.write htmlResponse
Next

grand_grand_total = total_admission_total + total_death_total

total_row_ele = total_row_ele & "<tr class='last'>"
total_row_ele = total_row_ele & "<td><b>Grand Total</b></td>"
total_row_ele = total_row_ele & "<td><b>" & male_agegroup_admission_total & "</b></td>"
total_row_ele = total_row_ele & "<td><b>" & female_agegroup_admission_total & "</b></td>"
total_row_ele = total_row_ele & "<td><b>" & total_admission_total & "</b></td>"
total_row_ele = total_row_ele & "<td><b>" & male_agegroup_death_total & "</b></td>"
total_row_ele = total_row_ele & "<td><b>" & female_agegroup_death_total & "</b></td>"
total_row_ele = total_row_ele & "<td><b>" & total_death_total & "</b></td>"
total_row_ele = total_row_ele & "<td><b>" & grand_grand_total & "</b></td>"
total_row_ele = total_row_ele & "</tr>"

response.write total_row_ele

response.write "</table> "

Function MainAdmission(visitationID)
  Dim admissionIDs, sql, rst, main_admission
  admissionIDs = ""
  sql = "SELECT TOP 1 AdmissionID FROM Admission WHERE visitationid = '" & visitationID & "' "
  sql = sql & "order by AdmissionDate asc"
  admissionIDs = ""
  Set rst = CreateObject("ADODB.Recordset")
  With rst
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
          main_admission = .fields("AdmissionID")
          .MoveNext
          Loop
      End If
      .Close
  End With
  Set rst = Nothing
  MainAdmission = main_admission
End Function

Function VisitationWithAdmission(genid, extra_filter)
Dim originalAdmission, sql, rst, admissionCount, mainAdmissions, visitationID, pDays
' Dim answers(2)
mainAdmissions = ""
admissionCount = 0
pDays = 0

sql = ""
sql = sql & "SELECT distinct v.VisitationID FROM Visitation AS v INNER JOIN Admission AS a ON v.VisitationID = a.VisitationID "
' sql = sql & "SELECT distinct a.VisitationID FROM Admission AS a INNER JOIN Visitation AS v ON a.VisitationID = v.VisitationID "
' sql = sql & "WHERE a.WorkingMonthID = v.WorkingMonthID AND a.WorkingMonthID = '" & workingMonth & "'"
' sql = sql & "WHERE a.WorkingMonthID = v.WorkingMonthID AND a.WorkingDayID between '" & startDate & "' And '" & endDate & "' "
sql = sql & "WHERE v.WorkingDayID between '" & startDate & "' And '" & endDate & "' "
' response.write sql
sql = sql & "and v.genderid = '" & genid & "' "
sql = sql & "and " & extra_filter

Set rst = CreateObject("ADODB.Recordset")
With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            visitationID = Trim(.fields("VisitationID"))
            main_ad = MainAdmission(visitationID)
            mainAdmissions = mainAdmissions & "'" & main_ad & "',"
            admissionCount = AllAdmissionByWard(mainAdmissions)
        .MoveNext
        Loop
    End If
    .Close
End With
Set rst = Nothing
VisitationWithAdmission = admissionCount
End Function

Function AllAdmissionByWard(mainAdmissionsIDs)
  Dim sql, rst, totalCount, main_IDs
  totalCount = 0
  main_IDs = TruncateLastComma(mainAdmissionsIDs)
  sql = "select count(*) as totalCount from Admission where AdmissionID in (" & main_IDs & ")"
  Set rst = CreateObject("ADODB.Recordset")
  With rst
      .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
          totalCount = CInt(.fields("totalCount"))
      End If
      .Close
  End With
  Set rst = Nothing

  AllAdmissionByWard = totalCount
End Function

Function TruncateLastComma(s)
    Dim string_without_trailing_comma
    string_without_trailing_comma = s
    If Right(Trim(s), 2) = "'," Then
        string_without_trailing_comma = Left(Trim(s), Len(s) - 1)
    End If
    TruncateLastComma = string_without_trailing_comma
End Function

Function DeathCount(genid, extra_filter)
    Dim sql, rst, totalDeath
    totalDeath = 0
    Set rst = CreateObject("ADODB.Recordset")
    sql = ""
    sql = sql & "SELECT count(*) as totalDeath FROM Visitation AS v INNER JOIN Admission AS a ON v.VisitationID = a.VisitationID "
    ' sql = sql & "WHERE a.WorkingMonthID = v.WorkingMonthID AND a.WorkingMonthID = '" & workingMonth & "'"
    sql = sql & "WHERE a.WorkingMonthID = v.WorkingMonthID AND a.WorkingDayID between '" & startDate & "' And '" & endDate & "' "
    sql = sql & "and a.GenderID = '" & genid & "' "
    sql = sql & "and " & extra_filter & " "
    sql = sql & "and v.medicaloutcomeid in ('m002', 'm002b')"
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            totalDeath = CInt(.fields("totalDeath"))
        End If
        .Close
    End With
    DeathCount = totalDeath
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
