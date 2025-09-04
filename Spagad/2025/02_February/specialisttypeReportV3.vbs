'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim rst, sql, cnt, gen

mth = Trim(request.querystring("BillMonthID"))
bgcat = Trim(request.querystring("BillGroupCatID"))

Styles
Header

Sub Header()
    Dim billMonths, billGroupCats

    sql = "SELECT BillMonthID, BillMonthName FROM BillMonth ORDER BY BillMonthID DESC"
    
    Set rst1 = CreateObject("ADODB.Recordset")
    rst1.open sql, conn, 3, 4

    billMonths = ""

    With rst1
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                optionHTML = "<option value='" & .fields("BillMonthID") & "'>" & .fields("BillMonthName") & "</option>"
                billMonths = billMonths & optionHTML
                .MoveNext
            Loop
        End If
    End With

    rst1.Close
    Set rst1 = Nothing

    sql = "SELECT BillGroupCatID, BillGroupCatName FROM BillGroupCat"
    
    Set rst2 = CreateObject("ADODB.Recordset")
    rst2.open sql, conn, 3, 4

    billGroupCats = ""

    With rst2
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                optionHTML = "<option value='" & .fields("BillGroupCatID") & "'>" & .fields("BillGroupCatName") & "</option>"
                billGroupCats = billGroupCats & optionHTML
                .MoveNext
            Loop
        End If
    End With

    rst2.Close
    Set rst2 = Nothing

response.write "<div class='filters'>"
    response.write "        <label for='billMonth' class='font-style'>Select BillMonth:</label><br>"
    response.write "        <select id='billMonth' name='billMonth' class='myselect'>"
    response.write billMonths
    response.write "        </select> <br><br>"

    response.write "        <label for='billGroupCat' class='font-style'>Select BillGroup:</label><br>"
    response.write "        <select id='billGroupCat' name='billGroupCat' class='myselect'>"
    response.write billGroupCats
    response.write "        </select>"
    response.write "        <button type='button' onclick='updateUrl()'>Show Data</button> <br />"
response.write "</div>"

response.write "<script>"
    response.write "    function updateUrl() {"
    
    response.write "        const billMonth = Array.from(document.getElementById('billMonth').selectedOptions).map(option => option.value).join(',');"
    response.write "        const billGroupCat = Array.from(document.getElementById('billGroupCat').selectedOptions).map(option => option.value).join(',');"
    response.write "        const baseUrl = 'http://192.168.2.14/hms/wpgPrtPrintLayoutAll.asp';"
    response.write "        const params = new URLSearchParams({"
    response.write "            PrintLayoutName: 'specialisttypeReport',"
    response.write "            PositionForTableName: 'WorkingDay',"
    response.write "            WorkingDayID: '',"
    response.write "            BillMonthID: billMonth,"
    response.write "            BillGroupCatID: billGroupCat"
    response.write "        });"
    response.write "        const newUrl = baseUrl + '?' + params.toString();"
    response.write "        window.location.href = newUrl;"
    response.write "        console.log(newUrl);"
    response.write "    }"
    response.write "</script>"

End Sub

Sub Styles()
    response.write "<style>"
        response.write "    .filters {"
        response.write "        padding: 10px;"
        response.write "        border: 1px solid #ccc;"
        response.write "        background-color: #f9f9f9;"
        response.write "        border-radius: 8px;"
        response.write "        width: 300px;"
        response.write "        margin-bottom: 15px;"
        response.write "        margin-top: 30px; "
        response.write "    }"
        
        response.write "h1 {"
        response.write "    font-size: 25px;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "    margin: 20px 0;"
        response.write "    text-transform: uppercase;"
        response.write "}"
        
        response.write "    .font-style {"
        response.write "        font-family: Arial, sans-serif;"
        response.write "        font-size: 14px;"
        response.write "        font-weight: bold;"
        response.write "        color: #333;"
        response.write "    }"
        
        response.write "    .myselect {"
        response.write "        width: 100%;"
        response.write "        padding: 8px;"
        response.write "        border-radius: 4px;"
        response.write "        border: 1px solid #aaa;"
        response.write "        font-size: 14px;"
        response.write "        background-color: #fff;"
        response.write "        cursor: pointer;"
        response.write "    }"
        
        response.write "    button {"
        response.write "        margin-top: 10px;"
        response.write "        padding: 8px 15px;"
        response.write "        border: none;"
        response.write "        background-color: #007BFF;"
        response.write "        color: white;"
        response.write "        font-size: 14px;"
        response.write "        border-radius: 4px;"
        response.write "        cursor: pointer;"
        response.write "        transition: background 0.3s ease-in-out;"
        response.write "    }"
        
        response.write "    button:hover {"
        response.write "        background-color: #0056b3;"
        response.write "    }"

        response.write "</style>"

End Sub


Set rst = CreateObject("ADODB.Recordset")
Set rst0 = CreateObject("ADODB.Recordset")
response.write "<style> table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 80vw;; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center;font-size:16px; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{ background-color: #FC046A; color:#fff } table#myTable thead .h_title{ background-color: blanchedalmond; } table#myTable thead .h_names{ font-size: 14px;} table#myTable tbody td{text-align:center;} table#myTable .last{background-color: #3C8F6D;color:#fff;font-weight:700;text-align:center;}  </style>"

sql = "Select v.VisitDate, v.PatientID, cast (v.PatientAge as int) as age, v.VisitCost as vCost, v.specialistTypeid, v.specialistid, v.sponsorid,v.InsuranceNo,v.BenefitTypeID From visitation v Where v.billMonthID IN ('" & mth & "') and v.BillGroupCatID='" & bgcat & "' order by v.VisitDate"
'sql = "Select v.VisitDate, v.PatientID, cast (v.PatientAge as int) as age, v.VisitCost, v.specialistTypeid, v.specialistid, v.InsuranceSchemeID, v.sponsorid From visitation v Where v.billMonthID IN ('" & mth & "') and v.specialistTypeID='" & spid & "' order by v.VisitDate"
cnt = 0
Dim tot, gTot
With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    response.write "<table id='myTable'> <thead><tr class='h_res'><th colspan='15'>Found " & rst.RecordCount & " Results...</th></tr> "
    response.write "<tr class='h_title'><th colspan='15'>Generated <u>" & GetComboName("SpecialistType", spid) & "</u> Report on <u>" & GetComboName("WorkingMonth", mth) & "</u></th></tr>"
    response.write "<tr class='h_names'><th>No.</th><th>Date</th><th>Patient Name</th><th>Age</th><th>Sponsor</th><th>Insurance No.</th><th>Relation</th><th>Visit Cost</th><th>Specialist Type</th><th>Specialists</th></tr></thead><tbody>"
  .MoveFirst
  Do While Not .EOF
  cnt = cnt + 1
  tot = .fields("vCost")
  gTot = gTot + tot
  response.write "<tr><td>" & cnt & "</td><td>" & FormatDate(.fields("VisitDate")) & "</td><td style='text-align:left;'>" & GetComboName("Patient", .fields("PatientID")) & "</td><td>" & .fields("age") & "</td><td>" & GetComboName("Sponsor", .fields("SponsorID")) & "</td><td>" & .fields("InsuranceNo") & "</td><td>" & GetComboName("BenefitType", .fields("BenefitTypeID")) & "</td><td>" & (FormatNumber(CStr(tot), 2, , , -1)) & "</td><td>" & GetComboName("SpecialistType", .fields("specialistTypeid")) & "</td><td>" & GetComboName("Specialist", .fields("specialistid")) & "</td></tr>"
  .MoveNext
  Loop
  response.write "<tr><td colspan='7'><b>GRAND TOTAL</b></td><td><b>" & (FormatNumber(CStr(gTot), 2, , , -1)) & "</b></td><td></td><td></td></tr>"
  Else
    response.write "<h1>No records found</h1>"
  End If
  rst.Close
  Set rst = Nothing
End With
response.write "</tbody></table><br><br>"

sql = "SELECT Specialistid AS specialistname, SUM(visitCost) AS amt, COUNT(visitationid) AS cnt FROM Visitation WHERE BillGroupCatID='" & bgcat & "' AND billmonthid='" & mth & "' GROUP BY specialistid ORDER BY SUM(visitcost) desc"
Dim mtot, ctot
With rst0
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    response.write "<table id='myTable'><thead><tr class='h_title'><td colspan='15'><b>SPECIALIST SUMMARY</b></td></tr> <tr class='h_names'><th>SPECIALIST NAME</th><th>COUNT</th><th>AMOUNT</th></tr></thead><tbody>"
  .MoveFirst
  Do While Not .EOF
    mtot = .fields("amt")
    ctot = ctot + mtot
    response.write "<tr><td style='text-align:left;'>" & GetComboName("specialist", .fields("specialistname")) & "</td> <td>" & .fields("cnt") & "</td> <td>" & (FormatNumber(CStr(mtot), 2, , , -1)) & "</td></tr>"
  .MoveNext
  Loop
  response.write "<tr><td colspan='2'><b>TOTAL </b></td><td><b>" & (FormatNumber(CStr(ctot), 2, , , -1)) & "</b></td></tr>"
  End If
  response.write "</tbody></table>"
rst0.Close
Set rst0 = Nothing
End With
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
