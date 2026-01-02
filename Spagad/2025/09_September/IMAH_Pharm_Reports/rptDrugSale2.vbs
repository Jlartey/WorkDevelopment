'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

server.scripttimeout = 60 * 60 * 60


Dim sql, sql2, rst, pst
Dim retAmt, TotQty, pUnit, dpt, totAmt, dt, SysUsr, periodStart, periodEnd, ar, cnt, DisQty, RetQty, soldQty, soldAmt
Dim cntOut

retAmt = 0
TotQty = 0
totAmt = 0

dt = Trim(Request.QueryString("printfilter0"))
SysUsr = Trim(Request.QueryString("printfilter1"))
SysUsr = ""

periodStart = CStr(Now)
periodEnd = CStr(Now)

ar = Split(dt, "||")
If UBound(ar) = 1 Then
    periodStart = ar(0)
    periodEnd = ar(1)
End If

response.write _
    "<style>" & _
        "table.tb {border-collapse:collapse; font-size:11px; width:100%}" & _
        "table.tb td, table.tb th { padding:8px 8px;}" & _
        "table.tb tr td {padding:2 10 2 10; }" & _
        "table.tb tr th {padding:2 10 2 10; text-align: left;color: #000;}" & _
        " .heading{text-align:left;margin:20px 0 10px 0;color:#00e;}" & _
        " a{margin: 0 5px;color:#00e;font-size:11px;cursor:pointer;padding:0 2px;}" & _
        " .a:hover {}" & _
        "h4 {margin:0px;}" & _
    "</style>"

response.write "<table>"
    response.write "<tr>"
        response.write "<td style='text-align:center'>"
                AddReportHeader
        response.write "</td>"
    response.write "</tr>"
response.write "</table>"

response.write "<hr/>"
response.write "<div style = 'text-align:center'><h4> DRUG SALES REPORT</h4></div>"
response.write "<div style = 'text-align:center;font-style:italic;font-size:11px;'> Store: " & GetComboName("DrugStore", GetDrugStore()) & "</div>"
response.write "<div style = 'text-align:center;font-style:italic;font-size:11px;'> From " & FormatDate(periodStart) & " to " & FormatDate(periodEnd) & "</div>"
If Len(Trim(SysUsr)) > 0 Then
    response.write "<div style = 'text-align:center;font-size:11px;'> For: " & GetComboName("Staff", GetComboNameFld("SystemUser", SysUsr, "StaffID")) & "</div>"
End If
response.write "<hr/>"

Set rst = server.CreateObject("ADODB.RecordSet")
Set pst = server.CreateObject("ADODB.RecordSet")

cnt = 1

pUnit = GetComboNameFld("JobSchedule", jSchd, "UnitID")
dpt = UCase(GetComboNameFld("JobSchedule", jSchd, "DepartmentID"))

sql = "SELECT DrugID, DrugName, AVG(UnitCost) AS UnitCost, SUM(Qty) AS DispenseQty, SUM(Amount) AS Amount FROM "
sql = sql & " (SELECT ds.DrugID, d.DrugName, ds.UnitCost, ds.Qty AS Qty, ds.FinalAmt AS Amount FROM DrugSaleItems as ds, Drug AS d "
sql = sql & "  WHERE ds.DrugID = d.DrugID AND JobscheduleId IN ('M0603', 'S22', 'M0602', 'M0601') "
sql = sql & "  AND DispenseDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' AND ds.DrugCategoryID<>'D002' "
If Len(Trim(SysUsr)) > 0 Then
    sql = sql & " AND SystemUserID = '" & SysUsr & "' "
End If
If dpt = "DPT002" Then
    sql = sql & " AND JobScheduleID IN (select jobScheduleID from JobSchedule where UnitID='" & pUnit & "')"
End If
sql = sql & " UNION ALL "
sql = sql & " SELECT ds.DrugID, d.DrugName, ds.UnitCost, ds.DispenseAmt1 AS Qty, ds.DispenseAmt2 AS Amount FROM DrugSaleItems2 as ds, Drug AS d "
sql = sql & " WHERE ds.DrugID = d.DrugID AND JobscheduleId IN ('M0603', 'S22', 'M0602', 'M0601') "
sql = sql & " AND DispenseDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' AND ds.DrugCategoryID<>'D002' "
If Len(Trim(SysUsr)) > 0 Then
    sql = sql & " AND SystemUserID = '" & SysUsr & "' "
End If
If dpt = "DPT002" Then
    sql = sql & " AND JobScheduleID IN (select jobScheduleID from JobSchedule where UnitID='" & pUnit & "')"
End If
sql = sql & " ) AS dsp GROUP BY DrugID, DrugName"

With rst
        .open sql, conn, 3, 4

        If .recordCount > 0 Then
                response.write _
                "<table  border='1' class ='tb'>" & _
                    "<tr style=""background-image:url('images/SysTheme/N010/bgMnuListHdrColorA.bmp')"">" & _
                        "<th>No.</th>" & _
                        "<th>DrugID</th>" & _
                        "<th>DrugName</th>" & _
                        "<th>Avg. UnitCost</th>" & _
                        "<th>Qty Issued</th>" & _
                        "<th>Qty Returned</th>" & _
                        "<th>Qty Sold</th>" & _
                        "<th>Amount</th>" & _
                        "<th>Starting Stock</th>" & _
                        "<th>Ending Stock</th>" & _
                    "</tr>"

                .MoveFirst

                Do While Not .EOF
                DisQty = FormatNumber(CDbl(.fields("DispenseQty")))
                RetQty = FormatNumber(CDbl(ReturnDrug2(.fields("DrugID"))))
                soldQty = DisQty - RetQty

                soldAmt = FormatNumber(CDbl(.fields("Amount")) - CDbl(retAmt))
                        response.write _
                                "<tr>" & _
                                "<td>" & cnt & "</td>" & _
                                "<td>" & UCase(.fields("DrugID")) & "</td>" & _
                                "<td>" & UCase(.fields("DrugName")) & "</td>" & _
                                "<td>" & FormatNumber(.fields("UnitCost")) & "</td>" & _
                                "<td>" & DisQty & "</td>" & _
                                "<td>" & RetQty & "</td>" & _
                                "<td>" & soldQty & "</td>" & _
                                "<td>" & soldAmt & "</td>" & _
                                "<td>" & GetDrugLevelByDate("M06", .fields("DrugID"), periodStart) & "</td>" & _
                                "<td>" & GetDrugLevelByDate("M06", .fields("DrugID"), periodEnd) & "</td>" & _
                            "</tr>"

                            TotQty = TotQty + soldQty
                            totAmt = totAmt + soldAmt
                            cnt = cnt + 1

                            .MoveNext
                            response.flush
                Loop
                response.write _
                                "<tr>" & _
                                "<td colspan='6' style = 'text-align:right;'><b>TOTAL</b></td>" & _
                                "<td><b>" & FormatNumber(CDbl(TotQty)) & "</b></td>" & _
                                "<td><b>" & FormatNumber(CDbl(totAmt)) & "</b></td>" & _
                            "</tr>"

                response.write "</table>"
        Else
                response.write "No Record Matching Selection Criteria found"
        End If
        .Close
End With


Function ReturnDrug2(DrgID)
        retAmt = 0

    sql2 = "SELECT SUM(ReturnQty) AS ReturnQty, SUM(Amount) AS ReturnAmt FROM "
    sql2 = sql2 & " ( SELECT ReturnQty, FinalAmt As Amount FROM DrugReturnItems WHERE DrugID ='" & DrgID & "' "
    sql2 = sql2 & " AND ReturnDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'  AND JobscheduleId IN ( 'M0603', 'S22', 'M0602', 'M0601') AND DrugCategoryID<>'D002' "
    If Len(Trim(SysUsr)) > 0 Then
        sql2 = sql2 & " AND SystemUserID = '" & SysUsr & "' "
    End If
    sql2 = sql2 & " UNION ALL "
    'sql2 = sql2 & "SELECT DispenseAmt1 As ReturnQty, DispenseAmt2 AS Amount FROM DrugReturnItems2 WHERE DrugID ='" & DrgID & "' "
    sql2 = sql2 & "SELECT ReturnQty As ReturnQty, MainItemValue1 AS Amount FROM DrugReturnItems2 WHERE DrugID ='" & DrgID & "' "
    sql2 = sql2 & " AND ReturnDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'  AND JobscheduleId IN ( 'M0603', 'S22', 'M0602', 'M0601') AND DrugCategoryID<>'D002' "
    If Len(Trim(SysUsr)) > 0 Then
        sql2 = sql2 & " AND SystemUserID = '" & SysUsr & "' "
    End If
    sql2 = sql2 & ") AS l"

    With pst
            .open sql2, conn, 3, 4

            If .recordCount > 0 Then

                .MoveFirst

                If Not IsNull(.fields("ReturnQty")) Then
                    ReturnDrug2 = .fields("ReturnQty")
                    retAmt = .fields("ReturnAmt")
                Else
                    ReturnDrug2 = 0
                    retAmt = 0
                End If
            End If

            .Close
    End With
End Function

Function GetDrugLevelByDate(sto, drugId, dat1)

    Dim dat2, ot
    dat2 = Now()

    Dim rst, sql, wkd, sp, rst1, tot, sm, cnt, smtot, gTot1, gTot2, gTot3, cnt2, sm2, ag, ag2, avaQty, posIn, posOut

    Dim drg, dt, dtIn, dtOut, md, qty, unt, ex, jb, br, totIn, cntIn, cntOut, totOut, currIn, currOut, disIn, disOut

    Dim supIn, supOut, reqIn1, reqOut1, reqIn2, reqOut2, disIn1, disOut1, disIn2, disOut2, adjIn, adjOut, PosCur

    gTot1 = 0
    gTot2 = 0
    gTot3 = 0
    cntIn = 0
    cntOut = 0
    dtIn = ""
    dtOut = ""
    totIn = 0
    totOut = 0
    currIn = 0
    currOut = 0

    supIn = 0
    supOut = 0
    reqIn1 = 0
    reqOut1 = 0
    reqIn2 = 0
    reqOut2 = 0
    disIn = 0
    disOut = 0
    adjIn = 0
    adjOut = 0

    jb = GetComboNameFld("DrugStore", sto, "JobScheduleID")
    br = GetComboNameFld("DrugStore", sto, "BranchID")
    avaQty = GetCurrDrugStockLev(sto, drugId)
    md = ""
    posIn = 0
    posOut = 0
    PosCur = 0
    Set rst = CreateObject("ADODB.Recordset")
    Set rst1 = CreateObject("ADODB.Recordset")

    'POS 4 INCOMING FROM SUPPLIER
    sql = "select sum(totalcost) as tot,sum(qty) as qt from incomingdrugitems "
    sql = sql & " where branchid='" & br & "' and jobscheduleid='" & jb & "' AND DrugCategoryID<>'D002' "
    sql = sql & " and entrydate between '" & dat1 & "' and '" & dat2 & "' and drugid='" & drugId & "'"

    With rst1
        .open sql, conn, 3, 4

        If .recordCount > 0 Then
            .MoveFirst

            If Not IsNull(.fields("qt")) Then
                If IsNumeric(.fields("qt")) Then
                    qty = .fields("qt")
                    totIn = totIn + qty
                    supIn = qty
                End If
            End If
        End If

        .Close
    End With

    'POS 5 RETURN TO SUPPLIER
    sql = "select sum(totalcost) as tot,sum(returnqty) as qt from drugtosupplieritem "
    sql = sql & " where jobscheduleid='" & jb & "'" ' and  branchid='" & br & "'"
    sql = sql & " and returndate between '" & dat1 & "' and '" & dat2 & "' and drugid='" & drugId & "' AND DrugCategoryID<>'D002'"

    With rst1
        .open sql, conn, 3, 4

        If .recordCount > 0 Then
            .MoveFirst

            If Not IsNull(.fields("qt")) Then
                If IsNumeric(.fields("qt")) Then
                    qty = .fields("qt")
                    totOut = totOut + qty
                    supOut = qty
                End If
            End If
        End If

        .Close
    End With

    'POS 6 SINGLE INCOMING REQUISITIONS
    sql = "select sum(issuedqty) as qt,sum(issuetotalcost) as tot from drugrequest "
    sql = sql & " where issueddate between '" & dat1 & "' and '" & dat2 & "' and drugid='" & drugId & "'"
    sql = sql & " and issuedqty>0 and itemrequeststageid='3' and drugStoreID='" & sto & "'"

    With rst1
        .open sql, conn, 3, 4

        If .recordCount > 0 Then
            .MoveFirst

            If Not IsNull(.fields("qt")) Then
                If IsNumeric(.fields("qt")) Then
                    qty = .fields("qt")
                    totIn = totIn + qty
                    reqIn1 = qty
                End If
            End If
        End If

        .Close
    End With

    'POS 7 SINGLE OUTGOING REQUISITION
    sql = "select sum(issuedqty) as qt,sum(issuetotalcost) as tot from drugrequest "
    sql = sql & " where issueddate between '" & dat1 & "' and '" & dat2 & "' and drugid='" & drugId & "'"
    sql = sql & " and issuedqty>0 and itemrequeststageid='3' and drugRequestStoreID='" & sto & "'"

    With rst1
        .open sql, conn, 3, 4

        If .recordCount > 0 Then
            .MoveFirst

            If Not IsNull(.fields("qt")) Then
                If IsNumeric(.fields("qt")) Then
                    qty = .fields("qt")
                    totOut = totOut + qty
                    reqOut1 = qty
                End If
            End If
        End If

        .Close
    End With

    'POS 8 MULTIPLE INCOMING REQUISITIONS
    sql = "select sum(issuedqty) as qt,sum(issuetotalcost) as tot from drugacceptitems "
    sql = sql & " where acceptdate1 between '" & dat1 & "' and '" & dat2 & "' and drugid='" & drugId & "'"
    sql = sql & " and issuedqty>0  and drugStoreID='" & sto & "' AND DrugCategoryID<>'D002'"

    With rst1
        .open sql, conn, 3, 4

        If .recordCount > 0 Then
            .MoveFirst

            If Not IsNull(.fields("qt")) Then
                If IsNumeric(.fields("qt")) Then
                    qty = .fields("qt")
                    totIn = totIn + qty
                    reqIn2 = qty
                End If
            End If
        End If

        .Close
    End With

    'POS 9 MULTIPLE OUTGOING REQUISITION
    sql = "select sum(issuedqty) as qt,sum(issuetotalcost) as tot from drugacceptitems "
    sql = sql & " where acceptdate1 between '" & dat1 & "' and '" & dat2 & "' and drugid='" & drugId & "'"
    sql = sql & " and issuedqty>0  and drugRequestStoreID='" & sto & "' AND DrugCategoryID<>'D002'"

    With rst1
        .open sql, conn, 3, 4

        If .recordCount > 0 Then
            .MoveFirst

            If Not IsNull(.fields("qt")) Then
                If IsNumeric(.fields("qt")) Then
                    qty = .fields("qt")
                    totOut = totOut + qty
                    reqOut2 = qty
                End If
            End If
        End If

        .Close
    End With

    'POS 10A OUTGOING SALE 1
    sql = "select sum(finalAmt) as tot,sum(qty) as qt from drugsaleitems "
    sql = sql & " where jobscheduleid='" & jb & "'" ' and  branchid='" & br & "'"
    sql = sql & " and dispensedate between '" & dat1 & "' and '" & dat2 & "' and drugid='" & drugId & "' AND DrugCategoryID<>'D002'"

    With rst1
        .open sql, conn, 3, 4

        If .recordCount > 0 Then
            .MoveFirst

            If Not IsNull(.fields("qt")) Then
                If IsNumeric(.fields("qt")) Then
                    qty = .fields("qt")
                    totOut = totOut + qty
                    disOut1 = qty
                End If
            End If
        End If

        .Close
    End With

    'POS 11A RETURN FROM SALES 1
    sql = "select sum(finalamt) as tot,sum(returnqty) as qt from drugreturnitems "
    sql = sql & " where jobscheduleid='" & jb & "'" ' and  branchid='" & br & "'"
    sql = sql & " and returndate between '" & dat1 & "' and '" & dat2 & "' and drugid='" & drugId & "' AND DrugCategoryID<>'D002'"

    With rst1
        .open sql, conn, 3, 4

        If .recordCount > 0 Then
            .MoveFirst

            If Not IsNull(.fields("qt")) Then
                If IsNumeric(.fields("qt")) Then
                    qty = .fields("qt")
                    totIn = totIn + qty
                    disIn1 = qty
                End If
            End If
        End If

        .Close
    End With

    'POS 10B OUTGOING SALES BY PRESCRIPTION
    sql = "select sum(finalAmt) as tot,sum(qty) as qt from drugsaleitems2 "
    sql = sql & " where jobscheduleid='" & jb & "'" ' and  branchid='" & br & "'"
    sql = sql & " and dispensedate between '" & dat1 & "' and '" & dat2 & "' and drugid='" & drugId & "' AND DrugCategoryID<>'D002'"

    With rst1
        .open sql, conn, 3, 4

        If .recordCount > 0 Then
            .MoveFirst

            If Not IsNull(.fields("qt")) Then
                If IsNumeric(.fields("qt")) Then
                    qty = .fields("qt")
                    totOut = totOut + qty
                    disOut2 = qty
                End If
            End If
        End If

        .Close
    End With


    ot = avaQty - totIn + totOut
    GetDrugLevelByDate = CStr(ot)
    Set rst = Nothing
    Set rst1 = Nothing
End Function

Function GetCurrDrugStockLev(sto, drg)

    Dim sql, rst, ot, dat2

    ot = 0
    Set rst = CreateObject("ADODB.Recordset")
    dat2 = CStr(Now())
    sql = "select availableqty from drugstocklevel "
    sql = sql & " where drugid='" & drg & "' and drugstoreid='" & sto & "' AND DrugCategoryID<>'D002'"

    With rst
        .open sql, conn, 3, 4

        If .recordCount > 0 Then
            .MoveFirst
            ot = .fields("availableqty")
            dat2 = CStr(Now())
        End If

        .Close
    End With

    GetCurrDrugStockLev = ot
    Set rst = Nothing
End Function
Function GetDrugStore()
    Dim ot, sql, rst

    sql = " select DrugStoreID from DrugStore where JobScheduleID='" & jSchd & "' "
    Set rst = CreateObject("ADODB.RecordSet")

    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        ot = rst.fields("DrugStoreID")
    Else
        rst.Close
        sql = "select DrugStoreID from DrugStore2 where JobScheduleID='" & jSchd & "' "
        rst.open sql, conn, 3, 4
        If rst.recordCount > 0 Then
            ot = rst.fields("DrugStoreID")
        End If
    End If
    Set rst = Nothing
    GetDrugStore = ot
End Function


'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
