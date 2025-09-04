'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim rst, sql, presc, arPeriod, periodStart, periodEnd, dat, grandTot, mth1, mth2, mth, yrs, dispType
response.write Glob_GetBootstrap5()
 dispType = ""
mth = Trim(Request.queryString("printfilter0"))
yrs = Trim(Request.queryString("printfilter1"))
'arPeriod = getDatePeriodFromDelim(Trim(Request.queryString("PrintFilter2")))
 dat = Trim(Request.queryString("PrintFilter2"))
 dispType = Trim(Request.queryString(("dispType")))
    Set rst = CreateObject("ADODB.RecordSet")
    
InitPageScript

'ar = Split(dat, "||")
'
'   If UBound(ar) = 1 Then
'    periodStart = ar(0)
'    periodEnd = ar(1)
'    Else
'   End If
   
   datsplit = Split(dat, "||")

If UBound(datsplit) >= 0 Then
   mthstart = datsplit(0)
   mthend = datsplit(1)
End If

mth1 = FormatWorkingMonth(mthstart)
mth2 = FormatWorkingMonth(mthend)
    
'If (Len(arPeriod) > 0) Then
'periodStart = arPeriod(0)
'periodEnd = arPeriod(1)
'End If
response.write "<table>"
response.write "<tr >"
response.write "<td>"
response.write "       <label for=""statusSelect"" class=""form-label"">Select Status</label>"
response.write "       <select class=""form-select"" name=""stat"" id=""stat"" onchange=""TypeOnchange()"">"
' items option
If dispType = "items" Or dispType = "" Then
    response.write "           <option value=""items"" selected>Report By Items</option>"
Else
    response.write "           <option value=""items"">Report By Items</option>"
End If
' sponsor option
'If dispType = "sponsor" Then
'    response.write "           <option value=""sponsor"" selected>Report By Sponsor</option>"
'Else
'    response.write "           <option value=""sponsor"">Report By Sponsor</option>"
'End If
'' sponsor type option
'If dispType = "sponsorTyp" Then
'    response.write "           <option value=""sponsorTyp"" selected>Report By Sponsor Type</option>"
'Else
'    response.write "           <option value=""sponsorTyp"">Report By Sponsor Type</option>"
'End If
' department option
'If dispType = "department" Then
'    response.write "           <option value=""department"" selected>Report By Department</option>"
''Else
''    response.write "           <option value=""department"">Report By Department</option>"
''End If


response.write "       </select>"
response.write "</td>"

response.write "   </tr>"
response.write "</table>"
processcode



Sub processcode()
    Dim rst, sql

    Set rst = CreateObject("ADODB.Recordset")

'    sql = ""
'sql = sql & " select tc.TreatmentID,sum(qty)as fQty ,sum(FinalAmt)as fAmt from TreatCharges tc"
'sql = sql & "  join treatment t On tc.TreatmentID  = t.TreatmentID"
'sql = sql & "  join visitation v On tc.visitationid = v.visitationid"
'sql = sql & " where tc.treatGroupid = 'T007' AND tc.billgroupid = 'B15' And tc.TreatTypeID NOT IN ('T004','T005','T010') And tc.sponsorid <> 'AFC'"
'If (Len(yrs) > 0) Then
'sql = sql & " AND v.billYearID = '" & yrs & "'"
'End If
'If (Len(mth) > 0) Then
'sql = sql & " AND v.billMonthID = '" & mth & "'"
'End If
'If (Len(dat) > 0) Then
'sql = sql & " AND v.billMonthID between '" & mth1 & "' AND '" & mth2 & "'"
'End If
'sql = sql & " Group By tc.TreatmentID"

sql = ""
sql = sql & " WITH services AS ("
sql = sql & "     SELECT "
sql = sql & "         tb.treatmentid, "
sql = sql & "         SUM(CASE WHEN tb.BillGroupID = 'B15' THEN 1 ELSE 0 END) AS TotalQty, "
sql = sql & "         AVG(tb.UnitCost) AS UnitCost, "
sql = sql & "         SUM(CASE WHEN tb.BillGroupID = 'B15' THEN tb.FinalAmt ELSE 0 END) AS TotalFinalAmt "
sql = sql & "     FROM "
sql = sql & "         (SELECT treatmentid, BillGroupID, UnitCost, FinalAmt "
sql = sql & "          FROM treatcharges tc "
sql = sql & "          WHERE tc.treatGroupid = 'T007' AND tc.billgroupid = 'B15' AND VisitationID IN "
sql = sql & "                (SELECT DISTINCT VisitationID "
sql = sql & "                 FROM Visitation "
sql = sql & "                 WHERE 1 = 1 "

If (Len(yrs) > 0) Then
    sql = sql & "                 AND BillYearID = '" & yrs & "' "
End If
If (Len(mth) > 0) Then
    sql = sql & "                 AND BillMonthID = '" & mth & "' "
End If
If (Len(dat) > 0) Then
    sql = sql & "                 AND BillMonthID BETWEEN '" & mth1 & "' AND '" & mth2 & "' "
End If

sql = sql & "                )"
sql = sql & "         ) AS tb "
sql = sql & "     GROUP BY tb.treatmentid "
sql = sql & " ) "
sql = sql & " "
sql = sql & " SELECT "
sql = sql & "     s.treatmentid, "
sql = sql & "     COALESCE(s.TotalQty, 0) AS NetQty, "
sql = sql & "     COALESCE(s.UnitCost, 0) AS UnitCost, "
sql = sql & "     COALESCE(s.TotalFinalAmt, 0) AS NetFinalAmt "
sql = sql & " FROM "
sql = sql & "     services s "
sql = sql & " WHERE "
sql = sql & "     COALESCE(s.TotalQty, 0) != 0;"


If (Len(yrs) > 0) Then
dtl = UCase(GetComboName("WorkingYear", yrs))
End If
If (Len(mth) > 0) Then
dtl = UCase(GetComboName("WorkingMonth", mth))
End If
If (Len(dat) > 0) Then
dtl = "From " & GetComboName("workingmonth", mth1) & " TO " & GetComboName("workingmonth", mth2) & ""
End If

    response.write "<table class='table table-bordered'>"
    response.write "<thead>"
    response.write "<tr>"
    response.write "<th Colspan= '999'><h5>PHARMACY REVENUE AND EXPENSE " & dtl & "</h5></th>"
    response.write "</tr>"
    response.write "<tr>"
    response.write "<th>#</th>"
    response.write "<th>ID</th>"
    response.write "<th>NAME</th>"
    response.write "<th>COUNT</th>"
    response.write "<th>AMOUNT</th>"
    response.write "</tr>"
    response.write "</thead>"
 
    response.flush
    response.write "<tbody>"

    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If (rst.RecordCount > 0) Then
           response.write "<tr>"
     response.write "<td Colspan= '999' style='color:red'> PROCEDURE AND DELIVERY</td>"
    response.write "</tr>"
            rst.MoveFirst
            Do While Not rst.EOF
                cnt = cnt + 1
                trtID = rst.fields("treatmentid")
                trtName = GetComboName("treatment", rst.fields("treatmentid"))
                qty = rst.fields("NetQty")
                amt = rst.fields("NetFinalAmt")

                response.write "<tr>"
                response.write "<td>" & cnt & "</td>"
                response.write "<td>" & trtID & "</td>"
                response.write "<td>" & trtName & "</td>"
                response.write "<td>" & qty & "</td>"
                response.write "<td>" & FormatNumber(amt, 2) & "</td>"
                response.write "</tr>"
                response.flush
                 totAmt = totAmt + amt
                 
                rst.MoveNext
            Loop
            grandTot = grandTot + totAmt
            response.write "<tr>"
            response.write "<td colspan = '4'><b>TOTALS</b></td>"
            response.write "<td><b>" & FormatNumber(totAmt, 2) & "</b></td>"
            
            response.write "</tr>"
        End If
    End With
    consumables
    response.flush
    processsub
    response.flush
    ward
    response.flush

     
         response.write "<tr class='bg-light'>"
            response.write "<td colspan = '4'><b>GRAND TOTAL </b></td>"
            response.write "<td><b>" & FormatNumber(grandTot, 2) & "</b></td>"
            
            response.write "</tr>"
    response.write "</tbody>"
    response.write "</table>"
End Sub

 response.write "</tbody>"
    response.write "</table>"


Sub processsub()
    Dim rst, sql

    Set rst = CreateObject("ADODB.Recordset")

'    sql = ""
'sql = sql & " select tc.TreatmentID,sum(qty)as fQty ,sum(FinalAmt)as fAmt from TreatCharges tc"
'sql = sql & "  join treatment t On tc.TreatmentID  = t.TreatmentID"
'sql = sql & "  join visitation v On tc.visitationid = v.visitationid"
'sql = sql & " where tc.treatGroupid = 'T001' AND tc.treattypeid = 'T001' AND tc.billgroupid = 'B15' And tc.TreatTypeID NOT IN ('T004','T005','T010') AND tc.sponsorid <> 'AFC'"


sql = ""
sql = sql & " WITH services AS ("
sql = sql & "     SELECT "
sql = sql & "         tb.treatmentid, "
sql = sql & "         SUM(CASE WHEN tb.BillGroupID = 'B15' THEN 1 ELSE 0 END) AS TotalQty, "
sql = sql & "         AVG(tb.UnitCost) AS UnitCost, "
sql = sql & "         SUM(CASE WHEN tb.BillGroupID = 'B15' THEN tb.FinalAmt ELSE 0 END) AS TotalFinalAmt "
sql = sql & "     FROM "
sql = sql & "         (SELECT treatmentid, BillGroupID, UnitCost, FinalAmt "
sql = sql & "          FROM treatcharges tc "
sql = sql & "          WHERE tc.treatGroupid NOT IN ('T007') "
sql = sql & "          AND tc.treattypeid NOT IN ('T008') "
sql = sql & "          AND tc.billgroupid = 'B15' "
sql = sql & "          AND VisitationID IN "
sql = sql & "                (SELECT DISTINCT VisitationID "
sql = sql & "                 FROM Visitation "
sql = sql & "                 WHERE 1 = 1 "

If (Len(yrs) > 0) Then
    sql = sql & "                 AND BillYearID = '" & yrs & "' "
End If

If (Len(mth) > 0) Then
    sql = sql & "                 AND BillMonthID = '" & mth & "' "
End If

If (Len(dat) > 0) Then
    sql = sql & "                 AND BillMonthID BETWEEN '" & mth1 & "' AND '" & mth2 & "' "
End If

sql = sql & "                )"
sql = sql & "         ) AS tb "
sql = sql & "     GROUP BY tb.treatmentid "
sql = sql & " ) "
sql = sql & " "
sql = sql & " SELECT "
sql = sql & "     s.treatmentid, "
sql = sql & "     COALESCE(s.TotalQty, 0) AS NetQty, "
sql = sql & "     COALESCE(s.UnitCost, 0) AS UnitCost, "
sql = sql & "     COALESCE(s.TotalFinalAmt, 0) AS NetFinalAmt "
sql = sql & " FROM "
sql = sql & "     services s "
sql = sql & " WHERE "
sql = sql & "     COALESCE(s.TotalQty, 0) != 0;"





'
'If (Len(yrs) > 0) Then
'sql = sql & " AND v.billYearID = '" & yrs & "'"
'End If
'If (Len(mth) > 0) Then
'sql = sql & " AND v.billMonthID = '" & mth & "'"
'End If
'If (Len(dat) > 0) Then
'sql = sql & " AND v.billMonthID between '" & mth1 & "' AND '" & mth2 & "'"
'End If
'sql = sql & " Group By tc.TreatmentID"


If (Len(yrs) > 0) Then
dtl = UCase(GetComboName("WorkingYear", yrs))
End If
If (Len(mth) > 0) Then
dtl = UCase(GetComboName("WorkingMonth", mth))
End If
If (Len(dat) > 0) Then
dtl = "From " & mth1 & " AND " & mth1 & ""
End If


    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If (rst.RecordCount > 0) Then
            response.write "<tr>"
     response.write "<td Colspan= '999' style='color:red'> Pharmacy AND OTHER SERVICES</td>"
    response.write "</tr>"
            rst.MoveFirst
            Do While Not rst.EOF
                cnt = cnt + 1
                trtID = rst.fields("treatmentid")
                trtName = GetComboName("treatment", rst.fields("treatmentid"))
                qty = rst.fields("NetQty")
                amt = rst.fields("NetFinalAmt")

                response.write "<tr>"
                response.write "<td>" & cnt & "</td>"
                response.write "<td>" & trtID & "</td>"
                response.write "<td>" & trtName & "</td>"
                response.write "<td>" & qty & "</td>"
                response.write "<td>" & FormatNumber(amt, 2) & "</td>"
                response.write "</tr>"
                response.flush
                 totAmt = totAmt + amt
                  
                rst.MoveNext
            Loop
            grandTot = grandTot + totAmt
            response.write "<tr>"
            response.write "<td colspan = '4'><b>TOTALS</b></td>"
            response.write "<td><b>" & FormatNumber(totAmt, 2) & "</b></td>"
            
            response.write "</tr>"
        End If
    End With

   
End Sub


Sub consumables()
    Dim rst, sql

    Set rst = CreateObject("ADODB.Recordset")


 If (dispType = "items" Or Len(dispType) < 1) Then
    typ = "drugid"
    qry = ""
 ElseIf (dispType = "sponsor") Then
    typ = "sponsorid"
    qry = "visitationid,"
  ElseIf (dispType = "department") Then
    typ = "jobscheduleid"
    qry = "visitationid,"
    
ElseIf (dispType = "sponsorTyp") Then
    typ = "Insurancetypeid"
    qry = "visitationid,"
    Else
    typ = "drugid"
    qry = ""
 End If

sql = ""
sql = sql & " WITH Sales AS ("
sql = sql & "     SELECT "
sql = sql & "         tb." & typ & ", "
If (typ <> "drugid") Then
    sql = sql & "         tb.drugid, "
End If
If (dispType = "sponsor" Or dispType = "department" Or dispType = "sponsorTyp") Then
sql = sql & "         COUNT(DISTINCT CASE WHEN tb.BillGroupID = 'B15' THEN tb.visitationid ELSE Null END) AS TotalQty, "
Else
sql = sql & "         SUM(CASE WHEN tb.BillGroupID = 'B15' THEN tb.Qty ELSE 0 END) AS TotalQty, "
End If
sql = sql & "         AVG(tb.UnitCost) AS UnitCost, "
sql = sql & "         SUM(CASE WHEN tb.BillGroupID = 'B15' THEN tb.FinalAmt ELSE 0 END) AS TotalFinalAmt "
sql = sql & "     FROM "
sql = sql & "         (SELECT " & qry & " drugid, BillGroupID, UnitCost,Qty, FinalAmt,sponsorid,jobscheduleid,insurancetypeid "
sql = sql & "          FROM DrugSaleItems "
sql = sql & "          WHERE VisitationID IN "
sql = sql & "                (SELECT DISTINCT VisitationID "
sql = sql & "                 FROM Visitation "
sql = sql & "                 WHERE 1 = 1 "

If (Len(yrs) > 0) Then
    sql = sql & "                 AND BillYearID = '" & yrs & "' "
End If

If (Len(mth) > 0) Then
    sql = sql & "                 AND BillMonthID = '" & mth & "' "
End If

If (Len(dat) > 0) Then
    sql = sql & "                 AND BillMonthID BETWEEN '" & mth1 & "' AND '" & mth2 & "' "
End If

sql = sql & "                ) "
sql = sql & "         UNION ALL "
sql = sql & "         SELECT " & qry & "drugid, BillGroupID, UnitCost,Qty, DispenseAmt2 AS FinalAmt,sponsorid,jobscheduleid,insurancetypeid "
sql = sql & "         FROM DrugSaleItems2 "
sql = sql & "         WHERE VisitationID IN "
sql = sql & "                (SELECT DISTINCT VisitationID "
sql = sql & "                 FROM Visitation "
sql = sql & "                 WHERE 1 = 1 "

If (Len(yrs) > 0) Then
    sql = sql & "                 AND BillYearID = '" & yrs & "' "
End If

If (Len(mth) > 0) Then
    sql = sql & "                 AND BillMonthID = '" & mth & "' "
End If

If (Len(dat) > 0) Then
    sql = sql & "                 AND BillMonthID BETWEEN '" & mth1 & "' AND '" & mth2 & "' "
End If

sql = sql & "                ) "
sql = sql & "         ) AS tb "
sql = sql & "     GROUP BY tb." & typ & " "
If (typ <> "drugid") Then
    sql = sql & ", tb.drugid"
End If
sql = sql & " ), "
sql = sql & " Returns AS ("
sql = sql & "     SELECT "
sql = sql & "         tb." & typ & ", "
If (typ <> "drugid") Then
    sql = sql & "         tb.drugid, "
End If
sql = sql & "         SUM(CASE WHEN tb.BillGroupID = 'B15' THEN tb.Qty ELSE 0 END) AS ReturnQty, "
sql = sql & "         AVG(tb.UnitCost) AS UnitCost, "
sql = sql & "         SUM(CASE WHEN tb.BillGroupID = 'B15' THEN tb.FinalAmt ELSE 0 END) AS ReturnFinalAmt "
sql = sql & "     FROM "
sql = sql & "         (SELECT " & qry & "drugid, BillGroupID,UnitCost,ReturnQty as Qty, FinalAmt,sponsorid,jobscheduleid,insurancetypeid "
sql = sql & "          FROM DrugReturnItems "
sql = sql & "          WHERE VisitationID IN "
sql = sql & "                (SELECT DISTINCT VisitationID "
sql = sql & "                 FROM Visitation "
sql = sql & "                 WHERE 1 = 1 "

If (Len(yrs) > 0) Then
    sql = sql & "                 AND BillYearID = '" & yrs & "' "
End If

If (Len(mth) > 0) Then
    sql = sql & "                 AND BillMonthID = '" & mth & "' "
End If

If (Len(dat) > 0) Then
    sql = sql & "                 AND BillMonthID BETWEEN '" & mth1 & "' AND '" & mth2 & "' "
End If

sql = sql & "                ) "
sql = sql & "         UNION ALL "
sql = sql & "         SELECT " & qry & "drugid, BillGroupID, UnitCost,ReturnQty as Qty, MainItemValue1 AS FinalAmt,sponsorid,jobscheduleid,insurancetypeid "
sql = sql & "         FROM DrugReturnItems2 "
sql = sql & "         WHERE VisitationID IN "
sql = sql & "                (SELECT DISTINCT VisitationID "
sql = sql & "                 FROM Visitation "
sql = sql & "                 WHERE 1 = 1 "

If (Len(yrs) > 0) Then
    sql = sql & "                 AND BillYearID = '" & yrs & "' "
End If

If (Len(mth) > 0) Then
    sql = sql & "                 AND BillMonthID = '" & mth & "' "
End If

If (Len(dat) > 0) Then
    sql = sql & "                 AND BillMonthID BETWEEN '" & mth1 & "' AND '" & mth2 & "' "
End If

sql = sql & "                ) "
sql = sql & "         ) AS tb "
sql = sql & "     GROUP BY tb." & typ & " "
If (typ <> "drugid") Then
    sql = sql & " , tb.drugid "
End If
sql = sql & " ) "

sql = sql & " SELECT "
sql = sql & "     s." & typ & ","
If (dispType = "sponsor" Or dispType = "department" Or dispType = "sponsorTyp") Then
sql = sql & "    COALESCE(SUM(s.TotalQty), 0) AS NetQty, "
Else
sql = sql & "      COALESCE(SUM(s.TotalQty), 0) - COALESCE(SUM(r.ReturnQty), 0) AS NetQty, "
End If
sql = sql & "     COALESCE(AVG(s.UnitCost), AVG(r.UnitCost)) AS UnitCost,  "
sql = sql & "     COALESCE(SUM(s.TotalFinalAmt), 0) - COALESCE(SUM(r.ReturnFinalAmt), 0) AS NetFinalAmt"
sql = sql & " FROM "
sql = sql & "     Sales s "
sql = sql & " LEFT JOIN "
sql = sql & "     Returns r ON s." & typ & " = r." & typ & ""
sql = sql & "  AND s.drugid = r.drugid "
sql = sql & "  GROUP BY s." & typ & ""

sql = sql & "  HAVING COALESCE(SUM(s.TotalQty), 0) - COALESCE(SUM(r.ReturnQty), 0) >= 1;"








If (Len(yrs) > 0) Then
dtl = UCase(GetComboName("WorkingYear", yrs))
End If
If (Len(mth) > 0) Then
dtl = UCase(GetComboName("WorkingMonth", mth))
End If
If (Len(dat) > 0) Then
dtl = "From " & mth1 & " TO " & mth1 & ""
End If

'    response.write "<table class='table table-bordered'>"
'    response.write "<thead>"
'    response.write "<tr>"
'    response.write "<th Colspan= '999'><h5>Pharmacy CONSUMABLE REVENUE  " & dtl & "</h5></th>"
'    response.write "</tr>"
'    response.write "<tr>"
'    response.write "<th>#</th>"
'    response.write "<th>ITEM ID</th>"
'    response.write "<th>ITEM NAME</th>"
'    response.write "<th>COUNT</th>"
'    response.write "<th>AMOUNT</th>"
'    response.write "</tr>"
'    response.write "</thead>"
'    response.write "<tbody>"

    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If (rst.RecordCount > 0) Then
            response.write "<tr>"
         response.write "<td Colspan= '999' style='color:red'> MEDICAL ITEMS / CONSUMABLES</td>"
        response.write "</tr>"
            rst.MoveFirst
            Do While Not rst.EOF
                cnt = cnt + 1
                'trtID = rst.fields("drugid")
                trtID = rst.fields("" & typ & "")
                'trtName = GetComboName("drug", rst.fields("drugid"))
                If (dispType = "items" Or dispType = "") Then
                trtName = GetComboName("drug", rst.fields("drugid"))
                ElseIf (dispType = "sponsor") Then
                trtName = GetComboName("sponsor", rst.fields("sponsorid"))
                ElseIf (dispType = "department") Then
                 trtName = GetComboName("Jobschedule", rst.fields("jobscheduleid"))
                 
                 ElseIf (dispType = "sponsorTyp") Then
                 trtName = GetComboName("InsuranceType", rst.fields("InsuranceTypeid"))
                End If

                qty = rst.fields("NetQty")
                amt = rst.fields("NetFinalAmt")

                response.write "<tr>"
                response.write "<td>" & cnt & "</td>"
                response.write "<td>" & trtID & "</td>"
                response.write "<td>" & trtName & "</td>"
                response.write "<td>" & qty & "</td>"
                response.write "<td>" & FormatNumber(amt, 2) & "</td>"
                response.write "</tr>"
                response.flush
                 totAmt = totAmt + amt
                
                rst.MoveNext
            Loop
              grandTot = grandTot + totAmt
            response.write "<tr>"
            response.write "<td colspan = '4'><b>TOTALS</b></td>"
            response.write "<td><b>" & FormatNumber(totAmt, 2) & "</b></td>"
            
            response.write "</tr>"
        End If
    End With

'    response.write "</tbody>"
'    response.write "</table>"
End Sub


Sub ward()
    Dim rst, sql

    Set rst = CreateObject("ADODB.Recordset")

'    sql = ""
'sql = sql & " select tc.TreatmentID,sum(qty)as fQty ,sum(FinalAmt)as fAmt from TreatCharges tc"
'sql = sql & " Inner join treatment t On tc.TreatmentID  = t.TreatmentID"
'sql = sql & " where tc.treatTypeid = 'T008'"
'If (Len(yrs) > 0) Then
'sql = sql & " AND tc.WorkingYearID = '" & yrs & "'"
'End If
'If (Len(mth) > 0) Then
'sql = sql & " AND tc.WorkingMonthID = '" & mth & "'"
'End If
'If (Len(dt) > 0) Then
'sql = sql & " AND tc.consultreviewdate between '" & periodStart & "' AND '" & periodEnd & "'"
'End If
'sql = sql & " Group By tc.TreatmentID"


sql = ""
sql = sql & " WITH services AS ("
sql = sql & "     SELECT "
sql = sql & "         tb.treatmentid, "
sql = sql & "         SUM(CASE WHEN tb.BillGroupID = 'B15' THEN 1 ELSE 0 END) AS TotalQty, "
sql = sql & "         AVG(tb.UnitCost) AS UnitCost, "
sql = sql & "         SUM(CASE WHEN tb.BillGroupID = 'B15' THEN tb.FinalAmt ELSE 0 END) AS TotalFinalAmt "
sql = sql & "     FROM "
sql = sql & "         (SELECT treatmentid, BillGroupID, UnitCost, FinalAmt "
sql = sql & "          FROM treatcharges tc "
sql = sql & "          WHERE tc.treatGroupid NOT IN ('T007') AND tc.treatTypeid = 'T008' "
sql = sql & "          AND VisitationID IN "
sql = sql & "                (SELECT DISTINCT VisitationID "
sql = sql & "                 FROM Visitation "
sql = sql & "                 WHERE 1 = 1 "

If (Len(yrs) > 0) Then
    sql = sql & "                 AND BillYearID = '" & yrs & "' "
End If

If (Len(mth) > 0) Then
    sql = sql & "                 AND BillMonthID = '" & mth & "' "
End If

If (Len(dat) > 0) Then
    sql = sql & "                 AND BillMonthID BETWEEN '" & mth1 & "' AND '" & mth2 & "' "
End If

sql = sql & "                ) "
sql = sql & "         ) AS tb "
sql = sql & "     GROUP BY tb.treatmentid "
sql = sql & " ) "
sql = sql & " "
sql = sql & " SELECT "
sql = sql & "     s.treatmentid, "
sql = sql & "     COALESCE(s.TotalQty, 0) AS NetQty, "
sql = sql & "     COALESCE(s.UnitCost, 0) AS UnitCost, "
sql = sql & "     COALESCE(s.TotalFinalAmt, 0) AS NetFinalAmt "
sql = sql & " FROM "
sql = sql & "     services s "
sql = sql & " WHERE "
sql = sql & "     COALESCE(s.TotalQty, 0) != 0;"






If (Len(yrs) > 0) Then
dtl = UCase(GetComboName("WorkingYear", yrs))
End If
If (Len(mth) > 0) Then
dtl = UCase(GetComboName("WorkingMonth", mth))
End If
If (Len(dt) > 0) Then
dtl = "From " & periodStart & " AND " & periodEnd & ""
End If

'    response.write "<table class='table table-bordered'>"
'    response.write "<thead>"
'    response.write "<tr>"
'    response.write "<th Colspan= '999'><h5>Pharmacy WARD REVENUE  " & dtl & "</h5></th>"
'    response.write "</tr>"
'    response.write "<tr>"
'    response.write "<th>#</th>"
'    response.write "<th>TREATMENT ID</th>"
'    response.write "<th>TREATMENT NAME</th>"
'    response.write "<th>COUNT</th>"
'    response.write "<th>AMOUNT</th>"
'    response.write "</tr>"
'    response.write "</thead>"
'    response.write "<tbody>"

    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If (rst.RecordCount > 0) Then
     response.write "<tr>"
     response.write "<td Colspan= '999' style='color:red'> WARD BILLING</td>"
    response.write "</tr>"
            rst.MoveFirst
            Do While Not rst.EOF
                cnt = cnt + 1
                trtID = rst.fields("treatmentid")
                trtName = GetComboName("treatment", rst.fields("treatmentid"))
                qty = rst.fields("NetQty")
                amt = rst.fields("NetFinalAmt")

                response.write "<tr>"
                response.write "<td>" & cnt & "</td>"
                response.write "<td>" & trtID & "</td>"
                response.write "<td>" & trtName & "</td>"
                response.write "<td>" & qty & "</td>"
                response.write "<td>" & FormatNumber(amt, 2) & "</td>"
                response.write "</tr>"
                response.flush
                 totAmt = totAmt + amt
                 
                rst.MoveNext
            Loop
             grandTot = grandTot + totAmt
            response.write "<tr>"
            response.write "<td colspan = '4'><b>TOTALS</b></td>"
            response.write "<td><b>" & FormatNumber(totAmt, 2) & "</b></td>"
            
            response.write "</tr>"
        End If
    End With

'    response.write "</tbody>"
'    response.write "</table>"
End Sub





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


Sub InitPageScript()


  Dim htStr
  'Client Script
  htStr = ""
  htStr = htStr & "<script id=""scptPrintLayoutExtraScript"" LANGUAGE=""javascript"">" & vbCrLf
  htStr = htStr & vbCrLf
    htStr = htStr & "function TypeOnchange(){" & vbCrLf
    htStr = htStr & "    var ur, typ;" & vbCrLf
    htStr = htStr & "    typ = GetEleVal('stat');" & vbCrLf
    'htStr = htStr & "    ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=SponsorBillV3B&PositionForTableName=Sponsor&PrintFilter=" & mth & "&PrintFilter0=" & mth & "&PrintFilter1=" & spn & "&PrintFilter2=" & brn & "&PrintFilter3=" & qtr & "&SponsorID=&dispType=' + typ;" & vbCrLf
    'htStr = htStr & "    ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=PharmacyRevenueReportByDrugs&PositionForTableName=Sponsor&PrintFilter=" & arPeriod & "&PrintFilter0=" & arPeriod & "&PrintFilter1=" & mth & "&PrintFilter2=" & spn & "&PrintFilter3=" & brn & "&PrintFilter4=" & qtr & "&SponsorID=&dispType=' + typ;" & vbCrLf
    htStr = htStr & "    ur = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=PharmacyRevenueReportByDrugs&PositionForTableName=WorkingDay&PrintFilter=" & mth & "&PrintFilter0=" & mth & "&PrintFilter1=" & yrs & "&PrintFilter2=" & dat & "&WorkingDayID=&dispType=' + typ;" & vbCrLf
    htStr = htStr & "    window.location.href = processurl(ur);" & vbCrLf
    htStr = htStr & "}" & vbCrLf


  htStr = htStr & "</script>"
  response.write htStr
  js = js & "<script>" & vbCrLf
  js = js & "  " & vbCrLf
  js = js & "  " & vbCrLf
  js = js & "</script>"
  response.write js
End Sub


'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
