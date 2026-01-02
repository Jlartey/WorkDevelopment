'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim dateRange, MedicalServiceID, linkType, drugId, SponsorID

dateRange = Request.QueryString("PrintFilter0")

If dateRange <> "" Then
    dateRange = Split(dateRange, "||")
End If

response.write Styles
'linkType = request.querystring("lnkType")
MedicalServiceID = Request.QueryString("MedicalServiceID")
SponsorID = Request.QueryString("SponsorID")

If UCase(uname) = "KENNETH" And False Then

Else


    If MedicalServiceID = "" Then
        'AddReportHeader
        PrintHeader dateRange, SponsorID
        PrintSponsorReport dateRange, SponsorID
    Else

        drugId = Request.QueryString("DrugID")
        PrintHeader2 dateRange, SponsorID
        PrintSponsorReportForDrug MedicalServiceID, dateRange, drugId, SponsorID
    End If

'Else

End If

Sub PrintSponsorReport2(dateRange, SponsorID)
    Dim sql, rst, str, lastMedicalServiceID, tot

    sql = "SELECT SponsorName"
    sql = sql & "  , DrugID, DrugName"
    sql = sql & "  , DIssuedQty "
    sql = sql & "  , DReturnQty"
    sql = sql & "  , PIssuedQty"
    sql = sql & "  , PReturnQty"
    sql = sql & "  , SUM(DIssuedQty) + SUM (PIssuedQty) AS IssuedQty "
    sql = sql & "  , SUM(DReturnQty) + SUM (PReturnQty) AS ReturnQty "
    sql = sql & "  , DrugID, DrugName"
    sql = sql & "  , DrugID, DrugName"
    sql = sql & "  , DrugID, DrugName"
    sql = sql & "  , DrugID, DrugName"
    sql = sql & "   FROM"
    sql = sql & "       ("
    sql = sql & "           SELECT MedicalServiceID, DrugID, DrugName "
    sql = sql & "               FROM DrugSaleItems "
    sql = sql & "               WHERE 1=1"
    sql = sql & "                   AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' AND DrugCategoryID<>'D002' "
    sql = sql & "       )"
    sql = sql & "        AS Dispenses"
    sql = sql & " "

    str = GetReportHeader

    rst = CreateObject("ADODB.RecordSet")

    response.write str
End Sub


Function GetReportHeader(dateRange, title)
    Dim str
    str = str & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
    str = str & "<tbody><tr><td align=""Center"" valign=""top""><img border=""0"" src=""images/logo.jpg"""
    str = str & "align=""Center""> </td><td align=""center"" height=""20"" bgcolor=""white"" "
    str = str & "style=""font-size: 14pt; color: #CCCCFF;  font-family:Arial"">"
    str = str & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" "
    str = str & "style=""font-family: Arial; font-size: 10pt; color: #000000"">"
    str = str & "<tbody><tr><td valign=""top""></td><td valign=""top""><b>IMaH&nbsp;&nbsp;HOSPITAL</b>"
    str = str & "</td><td valign=""top""></td><td valign=""top""><b>/</b></td></tr><tr>"
    str = str & "<td valign=""top""></td><td valign=""top""><b>Hospital</b></td>"
    str = str & "<td valign=""top""></td><td valign=""top""><b>/</b></td></tr><tr>"
    str = str & "<td valign=""top""></td><td valign=""top""><b>/</b></td>"
    str = str & "<td valign=""top""></td><td valign=""top""><b>/</b></td>"
    str = str & "</tr><tr><td valign=""top""></td><td valign=""top""><b>/</b></td>"
    str = str & "<td valign=""top""></td></tr><tr></tr></tbody></table></td></tr>"
    str = str & "<tr><td colspan=""2""><hr/></td><tr/>"
    str = str & "<tr><td colspan=""2"" style=""text-align: center;""><h2>" & title & "</h2></td><tr/>"
    str = str & "<tr><td colspan=""2"" style=""text-align: center;"">From <span>" & dateRange(0) & "</span> To <span>" & dateRange(1) & "</span> </td><tr/>"
    str = str & "<tr><td colspan=""2""><hr/></td><tr/>"
    str = str & "</tbody></table>"

    GetReportHeader = str

End Function

Sub PrintSponsorReport(dateRange, SponsorID)
    Dim sql, rst, lastMedicalServiceID, str, tot
    Dim cAmt, cQty, cRetQty, cIQty, cpdRetQty, cpdIQty, cddRetQty, cddIQty
    Dim globalCount, whcls, dpt, pUnit

    cAmt = 0
    cQty = 0
    cRetQty = 0
    cIQty = 0
    cpdRetQty = 0
    cpdIQty = 0
    cddRetQty = 0
    cddIQty = 0
    tot = 0

    pUnit = GetComboNameFld("JobSchedule", jSchd, "UnitID")
    dpt = UCase(GetComboNameFld("JobSchedule", jSchd, "DepartmentID"))

    'DispenseDate
    whcls = whcls & "  BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "'"
    If dpt = "DPT002" Then
        whcls = whcls & " AND JobScheduleID IN (select jobScheduleID from JobSchedule where UnitID='" & pUnit & "')"
    End If
    whcls = whcls & ""
    If SponsorID <> "" Then
        whcls = whcls & " AND SponsorID='" & SponsorID & "'"
    End If
    'WhCls = WhCls & " AND DrugID='" & drugID & "'  "
    'WhCls = WhCls & " AND MedicalServiceID='" & MedicalServiceID & "'"


    sql = "SELECT "
    sql = sql & "   Mydrug.MedicalServiceID, MyDrug.DrugID, SUM(DDispQty) AS [Direct Dispenses],                "
    sql = sql & "   SUM([Ret. Dir. Disp]) AS [Ret. Dir. Disp],  SUM(PDispQty) AS [Prescribed Dispenses], "
    sql = sql & "   SUM([Ret. Presc. Disp]) AS [Ret. Presc. Disp],  SUM(Qty) AS [Total Issued],          "
    sql = sql & "   SUM([Ret. Dir. Disp]) + SUM([Ret. Presc. Disp]) AS [Total Returned],                 "
    sql = sql & "   (SUM(Qty) - (  SUM([Ret. Dir. Disp]) + SUM([Ret. Presc. Disp]))) AS [Qty Sold],      "
    sql = sql & " (  ( SUM(ISNULL(FinalSaleAmt,0)) + SUM(ISNULL(FinalSaleAmt2,0))  )                     "
    sql = sql & "         - ( SUM(ISNULL(FinalRetAmt1,0)) + SUM(ISNULL(FinalRetAmt2, 0)) )               "
    sql = sql & "   ) AS [FinalAmt]                                                                      "
    sql = sql & "                                                                                        "
    sql = sql & " FROM                                                                                   "
    sql = sql & "   (                                                                                    "
    sql = sql & "   SELECT DrugID,MedicalServiceID, Qty, JobScheduleID, DispenseDate,                           "
    sql = sql & "          Qty AS [DDispQty], 0 AS [PDispQty], 0 AS [Ret. Presc. Disp],                  "
    sql = sql & "          0 AS [Ret. Dir. Disp], finalAmt AS [FinalSaleAmt], 0 AS [FinalSaleAmt2],      "
    sql = sql & "          0 AS [FinalRetAmt1], 0 AS [FinalRetAmt2]                                      "
    sql = sql & "   FROM DrugSaleItems     WHERE 1=1 AND DrugCategoryID<>'D002' AND DispenseDate " & whcls & "                      "
    sql = sql & "                                                                                        "
    sql = sql & "   UNION ALL                                                                            "
    sql = sql & "   SELECT DrugID"
    sql = sql & "          , (SELECT MedicalServiceID FROM DrugSale WHERE m.DrugSaleID=DrugSale.DrugSaleID) AS MedicalServiceID "
    sql = sql & "          , DispenseAmt1 AS [Qty], JobScheduleID,                       "
    sql = sql & "          DispenseDate, 0 AS [DDispQty], DispenseAmt1 AS [PDispQty],                    "
    sql = sql & "          0 AS [Ret. Presc. Disp], 0 AS [Ret. Dir. Disp],                               "
    sql = sql & "          0 AS [FinalSaleAmt], DispenseAmt2 AS [FinalSaleAmt2],                         "
    sql = sql & "          0 AS [FinalRetAmt1], 0 AS [FinalRetAmt2]                                      "
    sql = sql & "                                                                                        "
    sql = sql & "   FROM DrugSaleItems2 AS m   WHERE 1=1 AND DrugCategoryID<>'D002' AND DispenseDate " & whcls & "                     "
    sql = sql & "                                                                                        "
    sql = sql & "   UNION ALL                                                                            "
    sql = sql & "   SELECT DrugID"
    sql = sql & "          , (SELECT MedicalServiceID FROM DrugSale WHERE m.DrugSaleID=DrugSale.DrugSaleID) AS MedicalServiceID "
    sql = sql & "          , 0 AS [Qty], JobScheduleID, ReturnDate AS DispenseDate,      "
    sql = sql & "          0 AS [DDispQty], 0 AS [PDispQty], 0 AS [Ret. Presc. Disp],                    "
    sql = sql & "          ReturnQty AS [Ret. Dir. Disp],                                                "
    sql = sql & "          0 AS [FinalSaleAmt], 0 AS [FinalSaleAmt2],                                    "
    'sql = sql & "          ReturnQty AS [FinalRetAmt1], 0 AS [FinalRetAmt2]                              "
    sql = sql & "          FinalAmt AS [FinalRetAmt1], 0 AS [FinalRetAmt2]                               "
    sql = sql & "   FROM DrugReturnItems AS m   WHERE 1=1 AND DrugCategoryID<>'D002' AND ReturnDate " & whcls & "                        "
    sql = sql & "                                                                                        "
    sql = sql & "   UNION ALL                                                                            "
    sql = sql & "   SELECT DrugID"
    sql = sql & "          , (SELECT MedicalServiceID FROM DrugSale WHERE m.DrugSaleID=DrugSale.DrugSaleID) AS MedicalServiceID "
    sql = sql & "          , 0 AS [Qty], JobScheduleID, ReturnDate AS DispenseDate,      "
    sql = sql & "          0 AS [DDispQty], 0 AS [PDispQty], ReturnQty AS [Ret. Presc. Disp],            "
    sql = sql & "          0 AS [Ret. Dir. Disp],                                                        "
    sql = sql & "          0 AS [FinalSaleAmt], 0 AS [FinalSaleAmt2],                                    "
    sql = sql & "          0 AS [FinalRetAmt1], MainItemValue1 AS [FinalRetAmt2]                         "
    sql = sql & "   FROM DrugReturnItems2 AS m  WHERE 1=1 AND DrugCategoryID<>'D002' AND ReturnDate " & whcls & "                       "
    sql = sql & "                                                                                        "
    sql = sql & "   )       AS MyDrug                                                                    "
    sql = sql & "                                                                                        "
    sql = sql & "   GROUP BY MyDrug.DrugID, MyDrug.MedicalServiceID                                             "
    sql = sql & "                                                                                        "
    sql = sql & "   ORDER BY MedicalServiceID, DrugID                                                           "

    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4

    lastMedicalServiceID = ""

    If rst.recordCount > 0 Then

        rst.movefirst

        Do While Not rst.EOF
            globalCount = globalCount + 1
            If lastMedicalServiceID <> rst.fields("MedicalServiceID") Then
                If lastMedicalServiceID <> "" Then
                    'end previous table
                    str = str & "</tbody>"

                    str = str & "<tfoot>"
                        str = str & "<tr>"
                            str = str & "<td colspan='2' class='move-right'> Sub Total</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cddIQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cddRetQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cpdIQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cpdRetQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cIQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cRetQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cAmt) & "</td>"
                        str = str & "</tr>"
                    str = str & "</tfoot>"
                    str = str & "</table><br/>"
                End If
                tot = cAmt + tot
                cAmt = 0: cQty = 0: cRetQty = 0: cIQty = 0: cpdRetQty = 0: cpdIQty = 0: cddRetQty = 0: cddIQty = 0
                'start a new table
                str = str & "<table border='1' class='table-style' style='width:100%'>"
                str = str & " <colgroup>"
                    str = str & " <col class=''>"
                    str = str & " <col class=''>"
                    str = str & " <col class='' style='background-color: #c6e2fd'>"
                    str = str & " <col class='' >"
                    str = str & " <col class='' style='background-color: #c6e2fd'>"
                    str = str & " <col class='' >"
                    str = str & " <col class='' style='background-color: #c6e2fd'>"
                    str = str & " <col class=''>"
                    str = str & " <col class=''>"
                    str = str & " <col class='' style='background-color: #c6e2fd'>"
                str = str & " </colgroup>"
                str = str & "<thead>"
                    str = str & "<tr class='MedicalService'>"
                        str = str & "<th colspan='10'>" & GetComboName("MedicalService", rst.fields("MedicalServiceID")) & "</th>"
                    str = str & "</tr>"
                    str = str & "<tr class='heading'>"
                        str = str & "<th rowspan='2'> Drug ID </th>"
                        str = str & "<th rowspan='2'> Drug Name </th>"
                        str = str & "<th colspan='2'> Direct Dispenses </th>"
                        str = str & "<th colspan='2'> Prescribed Dispense </th>"
                        str = str & "<th rowspan='2'> Total Issued </th>"
                        str = str & "<th rowspan='2'> Total Returned </th>"
                        str = str & "<th rowspan='2'> Quantity Sold </th>"
                        str = str & "<th rowspan='2'> Amount </th>"
                    str = str & "</tr>"
                    str = str & "<tr class='heading'>"
                        str = str & "<th> Issued Qty. </th>"
                        str = str & "<th> Returned Qty. </th>"
                        str = str & "<th> Issued Qty. </th>"
                        str = str & "<th> Returned Qty. </th>"

                    str = str & "</tr>"
                str = str & "</thead>"
                str = str & "<tbody>"
            End If


            str = str & "<tr>"
                str = str & "<td class='td-col drug' onclick=""run_cmd('" & rst.fields("DrugID") & "', '" & rst.fields("MedicalServiceID") & "')"">" & rst.fields("DrugID") & "</td>"
                str = str & "<td class='td-col drug' onclick=""run_cmd('" & rst.fields("DrugID") & "', '" & rst.fields("MedicalServiceID") & "')"">" & GetComboName("Drug", (rst.fields("DrugID"))) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Direct Dispenses")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Ret. Dir. Disp")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Prescribed Dispenses")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Ret. Presc. Disp")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Total Issued")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Total Returned")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Qty Sold")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("FinalAmt")) & "</td>"
            str = str & "</tr>"

            cAmt = cAmt + FormatNumber(rst.fields("FinalAmt"))
            cQty = cQty + (rst.fields("Qty Sold"))
            cRetQty = cRetQty + (rst.fields("Total Returned"))
            cIQty = cIQty + (rst.fields("Total Issued"))
            cpdRetQty = cpdRetQty + (rst.fields("Ret. Presc. Disp"))
            cpdIQty = cpdIQty + (rst.fields("Prescribed Dispenses"))
            cddRetQty = cddRetQty + (rst.fields("Ret. Dir. Disp"))
            cddIQty = cddIQty + (rst.fields("Direct Dispenses"))

            lastMedicalServiceID = rst.fields("MedicalServiceID")
            rst.MoveNext
            response.write str
            response.Flush
            str = ""

        Loop
         'end last table
'        str = str & "</tbody>"
'        str = str & "</table><br/>"

        'end last table
        str = str & "</tbody>"

        str = str & "<tfoot>"
            str = str & "<tr>"
                str = str & "<td colspan='2' class='move-right'> Sub Total</td>"
                str = str & "<td class='move-right'>" & cddIQty & "</td>"
                str = str & "<td class='move-right'>" & cddRetQty & "</td>"
                str = str & "<td class='move-right'>" & cpdIQty & "</td>"
                str = str & "<td class='move-right'>" & cpdRetQty & "</td>"
                str = str & "<td class='move-right'>" & cIQty & "</td>"
                str = str & "<td class='move-right'>" & cRetQty & "</td>"
                str = str & "<td class='move-right'>" & cQty & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(cAmt) & "</td>"
            str = str & "</tr>"
        str = str & "</tfoot>"

        str = str & "</table><br/>"
        tot = cAmt + tot
        rst.Close
        Set rst = Nothing
    End If

    str = str & "<div style='text-align:center;'><hr color='black'/>"
    str = str & "<b>Total Amt. </b><span style='font-size: 15pt; font-weight: bold;'>" & FormatNumber(tot) & "</span>"
    str = str & "<hr color='black'/></div>"

    response.write str
End Sub

Sub PrintHeader(dateRange, SponsorID)
    response.write "<hr/>"
    response.write "<h2 style='font-size:12px;text-align:center;'>Drug Sale Report (Service Type) </h2>"
    response.write "<hr/>"
    response.write "<h2 style='font-size:12px;text-align:center;'>Store: " & GetComboName("DrugStore", GetDrugStore()) & "</h2>"
    response.write "<div style='text-align: left;font-size:12px;text-align:center;' ><b>Start Date/Time: </b>" & dateRange(0) & "</div>"
    response.write "<div style='text-align: left;font-size:12px;text-align:center;' ><b>End Date/Time: </b>" & dateRange(1) & "</div>"
    response.write "<hr/>"
    response.write "<div style='text-align: left;font-size:12px;' >" & AddSponsorLink(dateRange, SponsorID) & "</div>"
    response.write "<hr/>"

End Sub
Sub PrintHeader2(dateRange, SponsorID)
    response.write "<hr/>"
    response.write "<h2 style='font-size:12px;text-align:center;'>Drug Sale Report (Service Type) </h2>"
    response.write "<hr/>"
    response.write "<h2 style='font-size:12px;text-align:center;'>Store: " & GetComboName("DrugStore", GetDrugStore()) & "</h2>"
    response.write "<div style='text-align: left;font-size:12px;text-align:center;' ><b>Start Date/Time: </b>" & dateRange(0) & "</div>"
    response.write "<div style='text-align: left;font-size:12px;text-align:center;' ><b>End Date/Time: </b>" & dateRange(1) & "</div>"
    response.write "<hr/>"
'    response.write "<div style='text-align: left' >" & AddSponsorLink(dateRange, sponsorID) & "</div>"
    response.write "<hr/>"

End Sub
Function AddSponsorLink(dateRange, SponsorID)
    Dim str, rst, sql, lnk, blLnk

    lnk = ""
    lnk = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=rptDrugSaleByMedicalService&PositionForTableName=WorkingDay&PrintFilter=" & dateRange(0) & "||" & dateRange(1) & "&PrintFilter0=" & dateRange(0) & "||" & dateRange(1) & "&WorkingDayID="


    sql = "SELECT SponsorName, SponsorID FROM Sponsor where SponsorStatusID='S001' ORDER BY SponsorName "

    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4

    blLnk = lnk & "&SponsorID="
    str = "<div><ol><li><a href='" & blLnk & "'> All Sponsors</a></li>"
    If rst.recordCount > 0 Then

        Do While Not rst.EOF
            blLnk = lnk & "&SponsorID=" & rst.fields("sponsorID")
            str = str & "<li><a href='" & blLnk & "' > " & rst.fields("SponsorName") & "</a></li>"
            rst.MoveNext
        Loop

        rst.Close
        Set rst = Nothing
        str = str & "</ol>"
    End If

    str = str & "<hr/>"
    If SponsorID <> "" Then
        str = str & "<br/><h3>" & GetComboName("Sponsor", SponsorID) & "</h3><br/>"
    Else
        str = str & "<br/>For All Sponsors <br/>"
    End If

    str = str & "</div>"
    AddSponsorLink = str

End Function

Function MakeMedicalServiceLink(drugId, spnID, dateRange)
    Dim drugName

    drugName = GetComboName("Drug", drugId)
    lnk = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=rptDrugSale&PositionForTableName=WorkingDay&PrintFilter=" & dateRange(0) & "||" & dateRange(1) & "&PrintFilter0=" & dateRange(0) & "||" & dateRange(1) & "&WorkingDayID="
    lnk = lnk & "&DrugID=" & drugId & "&MedicalServiceID=" & spnID
    MakeMedicalServiceLink = "<a href='" & lnk & "' > " & drugName & "</a>"
End Function


Sub PrintSponsorReportForDrug(MedicalServiceID, dateRange, drugId, SponsorID)
    Dim sql, rst, lastMedicalServiceID, str, tot
    Dim cAmt, cQty, cRetQty, cIQty, cpdRetQty, cpdIQty, cddRetQty, cddIQty
    Dim globalCount, whcls, dpt, pUnit

    'Response.Write "<br/>MedicalService ID " & MedicalServiceID & " DrugID " & drugID

    pUnit = GetComboNameFld("JobSchedule", jSchd, "UnitID")
    dpt = UCase(GetComboNameFld("JobSchedule", jSchd, "DepartmentID"))

    'DispenseDate
    whcls = whcls & "  BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "'"
    If dpt = "DPT002" Then
        sql = sql & " AND JobScheduleID IN (select jobScheduleID from JobSchedule where UnitID='" & pUnit & "')"
    End If
    whcls = whcls & ""
    whcls = whcls & " AND DrugID='" & drugId & "'  "
    whcls = whcls & " AND (SELECT MedicalServiceID FROM DrugSale WHERE m.DrugSaleID=DrugSale.DrugSaleID)='" & MedicalServiceID & "'"
    If SponsorID <> "" Then
        whcls = whcls & " AND SponsorID='" & SponsorID & "'"
    End If

    sql = "SELECT "
    sql = sql & "   DrugSaleID, PatientID,Mydrug.MedicalServiceID, MyDrug.DrugID,                               "
    sql = sql & "   SUM(DDispQty) AS [Direct Dispenses],                                                 "
    sql = sql & "   SUM([Ret. Dir. Disp]) AS [Ret. Dir. Disp],  SUM(PDispQty) AS [Prescribed Dispenses], "
    sql = sql & "   SUM([Ret. Presc. Disp]) AS [Ret. Presc. Disp],  SUM(Qty) AS [Total Issued],          "
    sql = sql & "   SUM([Ret. Dir. Disp]) + SUM([Ret. Presc. Disp]) AS [Total Returned],                 "
    sql = sql & "   (SUM(Qty) - (  SUM([Ret. Dir. Disp]) + SUM([Ret. Presc. Disp]))) AS [Qty Sold],      "
    sql = sql & " (  ( SUM(ISNULL(FinalSaleAmt,0)) + SUM(ISNULL(FinalSaleAmt2,0))  )                     "
    sql = sql & "         - ( SUM(ISNULL(FinalRetAmt1,0)) + SUM(ISNULL(FinalRetAmt2, 0)) )               "
    sql = sql & "   ) AS [FinalAmt]                                                                      "
    sql = sql & "                                                                                        "
    sql = sql & " FROM                                                                                   "
    sql = sql & "   (                                                                                    "
    sql = sql & "   SELECT DrugSaleID, PatientID,DrugID,MedicalServiceID, Qty, JobScheduleID, DispenseDate,     "
    sql = sql & "          Qty AS [DDispQty], 0 AS [PDispQty], 0 AS [Ret. Presc. Disp],                  "
    sql = sql & "          0 AS [Ret. Dir. Disp], finalAmt AS [FinalSaleAmt], 0 AS [FinalSaleAmt2],      "
    sql = sql & "          0 AS [FinalRetAmt1], 0 AS [FinalRetAmt2]                                      "
    sql = sql & "   FROM DrugSaleItems AS m  WHERE 1=1 AND DrugCategoryID<>'D002' AND DispenseDate " & whcls & "                        "
    sql = sql & "                                                                                        "
    sql = sql & "   UNION                                                                                "
    sql = sql & "   SELECT DrugSaleID, PatientID,DrugID "
    sql = sql & "          , (SELECT MedicalServiceID FROM DrugSale WHERE m.DrugSaleID=DrugSale.DrugSaleID) AS MedicalServiceID "
    sql = sql & "          , DispenseAmt1 AS [Qty], JobScheduleID, "
    sql = sql & "          DispenseDate, 0 AS [DDispQty], DispenseAmt1 AS [PDispQty],                    "
    sql = sql & "          0 AS [Ret. Presc. Disp], 0 AS [Ret. Dir. Disp],                               "
    sql = sql & "          0 AS [FinalSaleAmt], DispenseAmt2 AS [FinalSaleAmt2],                         "
    sql = sql & "          0 AS [FinalRetAmt1], 0 AS [FinalRetAmt2]                                      "
    sql = sql & "                                                                                        "
    sql = sql & "   FROM DrugSaleItems2 AS m   WHERE 1=1 AND DrugCategoryID<>'D002' AND DispenseDate " & whcls & "                       "
    sql = sql & "                                                                                        "
    sql = sql & "   UNION                                                                                "
    sql = sql & "   SELECT DrugSaleID, PatientID,DrugID"
    sql = sql & "          , (SELECT MedicalServiceID FROM DrugSale WHERE m.DrugSaleID=DrugSale.DrugSaleID) AS MedicalServiceID "
    sql = sql & "           , 0 AS [Qty], JobScheduleID,            "
    sql = sql & "          ReturnDate AS DispenseDate,                                                   "
    sql = sql & "          0 AS [DDispQty], 0 AS [PDispQty], 0 AS [Ret. Presc. Disp],                    "
    sql = sql & "          ReturnQty AS [Ret. Dir. Disp],                                                "
    sql = sql & "          0 AS [FinalSaleAmt], 0 AS [FinalSaleAmt2],                                    "
    'sql = sql & "          ReturnQty AS [FinalRetAmt1], 0 AS [FinalRetAmt2]                              "
    sql = sql & "          FinalAmt AS [FinalRetAmt1], 0 AS [FinalRetAmt2]                               "
    sql = sql & "   FROM DrugReturnItems  AS m  WHERE 1=1 AND DrugCategoryID<>'D002' AND ReturnDate " & whcls & "                       "
    sql = sql & "                                                                                        "
    sql = sql & "   UNION                                                                                "
    sql = sql & "   SELECT DrugSaleID, PatientID,DrugID"
    sql = sql & "          , (SELECT MedicalServiceID FROM DrugSale WHERE m.DrugSaleID=DrugSale.DrugSaleID) AS MedicalServiceID "
    sql = sql & "          , 0 AS [Qty], JobScheduleID,            "
    sql = sql & "          ReturnDate AS DispenseDate,                                                   "
    sql = sql & "          0 AS [DDispQty], 0 AS [PDispQty], ReturnQty AS [Ret. Presc. Disp],            "
    sql = sql & "          0 AS [Ret. Dir. Disp],                                                        "
    sql = sql & "          0 AS [FinalSaleAmt], 0 AS [FinalSaleAmt2],                                    "
    sql = sql & "          0 AS [FinalRetAmt1], MainItemValue1 AS [FinalRetAmt2]                         "
    sql = sql & "   FROM DrugReturnItems2 AS m  WHERE 1=1 AND DrugCategoryID<>'D002' AND ReturnDate " & whcls & "                        "
    sql = sql & "                                                                                        "
    sql = sql & "   )       AS MyDrug                                                                    "
    sql = sql & "                                                                                        "
    sql = sql & "   GROUP BY MyDrug.DrugID, MyDrug.MedicalServiceID,DrugSaleID, PatientID                       "
    sql = sql & "                                                                                        "
    sql = sql & "   ORDER BY MedicalServiceID, DrugID                                                           "

    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4

    lastMedicalServiceID = ""

    If rst.recordCount > 0 Then

        rst.movefirst

        Do While Not rst.EOF
            globalCount = globalCount + 1
            If lastMedicalServiceID <> rst.fields("MedicalServiceID") Then
                If lastMedicalServiceID <> "" Then
                    'end previous table
                    str = str & "</tbody>"

                    str = str & "<tfoot>"
                        str = str & "<tr>"
                            str = str & "<td colspan='2' class='move-right'> Sub Total</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cddIQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cddRetQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cpdIQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cpdRetQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cIQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cRetQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cQty) & "</td>"
                            str = str & "<td class='move-right'>" & FormatNumber(cAmt) & "</td>"
                        str = str & "</tr>"
                    str = str & "</tfoot>"
                    str = str & "</table><br/>"
                End If

                cAmt = 0: cQty = 0: cRetQty = 0: cIQty = 0: cpdRetQty = 0: cpdIQty = 0: cddRetQty = 0: cddIQty = 0
                'start a new table
                str = str & "<table border='1' class='table-style' style='width:100%;'>"
                str = str & "<thead >"
                    str = str & "<tr class='MedicalService'>"
                        str = str & "<th colspan='10'>" & GetComboName("MedicalService", rst.fields("MedicalServiceID")) & "</th>"
                    str = str & "</tr>"
                    str = str & "<tr class='heading'>"
                        str = str & "<th rowspan='2'> Dispense ID </th>"
                        str = str & "<th rowspan='2'> Patient Name </th>"
                        str = str & "<th colspan='2'> Direct Dispenses </th>"
                        str = str & "<th colspan='2'> Prescribed Dispense </th>"
                        str = str & "<th rowspan='2'> Total Issued </th>"
                        str = str & "<th rowspan='2'> Total Returned </th>"
                        str = str & "<th rowspan='2'> Quantity Sold </th>"
                        str = str & "<th rowspan='2'> Amount </th>"
                    str = str & "</tr>"
                    str = str & "<tr class='heading'>"
                        str = str & "<th> Issued Qty. </th>"
                        str = str & "<th> Returned Qty. </th>"
                        str = str & "<th> Issued Qty. </th>"
                        str = str & "<th> Returned Qty. </th>"

                    str = str & "</tr>"
                str = str & "</thead>"
                str = str & "<tbody>"
            End If


            str = str & "<tr>"
                str = str & "<td class='td-col'>" & MakeDispenseLink(rst.fields("DrugSaleID")) & "</td>"
                str = str & "<td class='td-col'>" & GetComboName("Patient", (rst.fields("PatientID"))) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Direct Dispenses")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Ret. Dir. Disp")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Prescribed Dispenses")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Ret. Presc. Disp")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Total Issued")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Total Returned")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("Qty Sold")) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(rst.fields("FinalAmt")) & "</td>"
            str = str & "</tr>"

            cAmt = cAmt + FormatNumber(rst.fields("FinalAmt"))
            cQty = cQty + (rst.fields("Qty Sold"))
            cRetQty = cRetQty + (rst.fields("Total Returned"))
            cIQty = cIQty + (rst.fields("Total Issued"))
            cpdRetQty = cpdRetQty + (rst.fields("Ret. Presc. Disp"))
            cpdIQty = cpdIQty + (rst.fields("Prescribed Dispenses"))
            cddRetQty = cddRetQty + (rst.fields("Ret. Dir. Disp"))
            cddIQty = cddIQty + (rst.fields("Direct Dispenses"))
            tot = cAmt + tot

            lastMedicalServiceID = rst.fields("MedicalServiceID")
            rst.MoveNext


        Loop
         'end last table
'        str = str & "</tbody>"
'        str = str & "</table><br/>"

        'end previous table
        str = str & "</tbody>"

        str = str & "<tfoot>"
            str = str & "<tr>"
                str = str & "<td colspan='2' class='move-right'> Sub Total</td>"
                str = str & "<td class='move-right'>" & FormatNumber(cddIQty) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(cddRetQty) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(cpdIQty) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(cpdRetQty) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(cIQty) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(cRetQty) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(cQty) & "</td>"
                str = str & "<td class='move-right'>" & FormatNumber(cAmt) & "</td>"
            str = str & "</tr>"
        str = str & "</tfoot>"
        str = str & "</table><br/>"

        rst.Close
        Set rst = Nothing
    End If

'    Response.Write "<div><b>Total Amt. <b/>" & tot & "</div>"
    response.write "<div><b>Drug Name: <b/>" & GetComboName("Drug", drugId) & "</div><br/>"
    response.write str
End Sub

Function MakeDispenseLink(drugsaleid)
    Dim lnk
    lnk = "wpgDrugSale.asp?PageMode=ProcessSelect&DrugSaleID=" & drugsaleid
    MakeDispenseLink = "<a href='" & lnk & "'>" & drugsaleid & "</a>"
End Function
Function Styles()
    Dim str
     str = str & "<style>"
        str = str & ".td-col{ white-space: nowrap;}"
        str = str & ".drug{ cursor: pointer; }"
        str = str & ".move-right{ white-space: nowrap; text-align: right; }"
        str = str & "table.table-style{ border-spacing: unset; border: 1px solid silver;}"
        str = str & "table.table-style td, table.table-style th{ font-size:11px;}"
        str = str & ".MedicalService{ background-color: #c6e2fd; font-size: larger;}"
        str = str & ".heading{ background-color: #c6e2fd;}"
        str = str & " tfoot tr td { font-weight: bold; background-color: unset; font-size: larger;}"
        str = str & ".td-col{ white-space: nowrap; padding: 3 10 3 10;}"
        str = str & ".move-right{ white-space: nowrap; text-align: right; }"
    str = str & "</style>"

    str = str & "<script>"
        str = str & " function run_cmd (drugid, spnID) {"
        str = str & "   console.log(drugid); " & vbCrLf
        str = str & "   window.open (window.location.href + '&MedicalServiceID=' + spnID + '&DrugID=' + drugid); " & vbCrLf
        str = str & " }"
    str = str & "</script>"

    Styles = str
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
