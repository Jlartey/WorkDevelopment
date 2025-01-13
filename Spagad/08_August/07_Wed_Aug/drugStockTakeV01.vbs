'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Function GetCurrDrugStockLev(sto, drg)
Dim sql, rst, ot
ot = 0
Set rst = CreateObject("ADODB.Recordset")
sql = "select availableqty from drugstocklevel "
sql = sql & " where drugid='" & drg & "' and drugstoreid='" & sto & "'"
With rst
.open qryPro.FltQry(sql), conn, 3, 4
If .RecordCount > 0 Then
.MoveFirst
ot = .fields("availableqty")
End If
.Close
End With
GetCurrDrugStockLev = ot
Set rst = Nothing
End Function

Sub TrackDrugStock()
Dim sql, rst, ot, cnt, rst0, sto, drg, stoNm, drgNm, expiryDate
Dim recPerPg, abPos, recPos, recCnt, num, pgStr, pgUrl
Dim stkTot, stkVal, stkPr, hrf, avaQty, tot, gTot
ot = ""
Set rst0 = CreateObject("ADODB.Recordset")
Set rst = CreateObject("ADODB.Recordset")
recPerPg = 5000
  sql = "select * from drugstore where drugstoreid='" & stoSlt & "' order by drugstoreid"
  rst0.open qryPro.FltQry(sql), conn, 3, 4
  If rst0.RecordCount > 0 Then
    rst0.MoveFirst
    cnt = 0
    tot = 0
    gTot = 0
    Do While Not rst0.EOF
      sto = rst0.fields("DrugStoreID")
      stoNm = rst0.fields("DrugStoreName")
      response.write "<table border=""1"" cellpadding=""2"" cellspacing=""0"" style=""font-size: 8pt; font-family: Arial; border-collapse:collapse;page-break-after:always"">"
'      response.write "<tr>"
'      response.write "<td colspan=""7"" align=""center""><b>" & UCase(stoNm) & " [" & UCase(sto) & "]</b> : DRUG STOCK VALUE AS AT [" & FormatDateDetail(dat1) & "]</td>"
'      response.write "</tr>"
      'Open Drug List
'      sql = "select drug.drugname,drug.RetailUnitCost,drugstocklevel.drugid from drugstocklevel,drug "
'      sql = sql & " where drugstocklevel.drugid=drug.drugid"
'      sql = sql & " and drug.DrugStatusID <> 'IST002' " '@bless'
'      sql = sql & " and drugstocklevel.drugstockstatusid='D001'"
'      sql = sql & " and drugstocklevel.drugstoreid='" & sto & "'"
'      sql = sql & " order by drug.drugName"

     'My addition
    sql = "SELECT drug.drugname, drug.RetailUnitCost, drugstocklevel.drugid, "
    sql = sql & "CASE WHEN IncomingDrugItems.ExpiryDate IS NULL THEN 'N/A' ELSE CONVERT(VARCHAR(20), IncomingDrugItems.ExpiryDate, 103) END AS ExpiryDate "
    sql = sql & "FROM drugstocklevel "
    sql = sql & "JOIN drug ON drugstocklevel.drugid = drug.drugid "
    sql = sql & "JOIN IncomingDrugItems ON drug.drugid = IncomingDrugItems.DrugID "
    sql = sql & "WHERE drug.DrugStatusID <> 'IST002' "
    sql = sql & "AND drugstocklevel.drugstockstatusid = 'D001' "
    sql = sql & "AND drugstocklevel.drugstoreid = '" & sto & "' "
    sql = sql & "ORDER BY drug.drugName"
    
    'response.write sql
      
      rst.CacheSize = CInt(recPerPg)
      rst.open qryPro.FltQry(sql), conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .PageSize = CInt(recPerPg)
          'Generate Paging Links
          pgUrl = GetPageUrlInfo()
          pgStr = "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size: 8pt; font-weight:bold; border-collapse:collapse""><tr>"
          pgStr = pgStr & "<td style=""font-size: 9pt; font-weight:normal"">FIND&nbsp;:&nbsp;</td>"
          For num = 1 To .PageCount
            .AbsolutePage = num
            pgStr = pgStr & "<td><a href=""" & Replace(pgUrl, "POSITIONFORRECPAGENUMBER", CStr(num)) & """>" & Left(.fields("DrugName"), 3) & "</a></td>"
          Next
          pgStr = pgStr & "</tr></table>"
          
          'Paging
'          response.write "<tr>"
'          response.write "<td colspan=""8"">" & pgStr & "</td>"
'          response.write "</tr>"
          'Header
          response.write "<tr>"
          response.write "<td><b>NO.</b></td>"
          response.write "<td><b>DRUGID</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>DRUG</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>U.PRICE</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>UOM</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>CURR. LEV</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>TOT</b></td>"
          response.write "<td valign=""top"" align=""Center"" width=""140""><b>STOCK TAKE</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>EXPIRY DATE</b></td>"
          
          response.write "</tr>"
          
          abPos = GetPrintPagePos(.PageCount)
          .AbsolutePage = abPos
          recPos = (abPos - 1) * CInt(recPerPg)
          cnt = 0
          recCnt = 0
          Do While Not .EOF And recCnt < CInt(recPerPg)
            cnt = cnt + 1
            recCnt = recCnt + 1
            recPos = recPos + 1
            drg = .fields("drugid")
            drgNm = .fields("drugName")
            expiryDate = .fields("ExpiryDate")
            
            avaQty = GetCurrDrugStockLev(sto, drg)
            stkTot = avaQty
            stkPr = .fields("RetailUnitCost") 'GetComboNameFld("Drug", drg, "RetailUnitCost") 'GetDrugPrice(drg, sto)
            tot = CDbl(stkTot) * CDbl(stkPr)
            gTot = gTot + tot
            'stkVal = CDbl(stkTot) * CDbl(stkPr)
            'totalStkVal = totalStkVal + stkVal
            response.write "<tr>"
            response.write "<td><b>" & CStr(recPos) & "</b></td>"
            response.write "<td>" & UCase(drg) & "</td>"
            response.write "<td>" & UCase(drgNm) & "</td>"
            response.write "<td align=""right"">" & FormatNumber(stkPr, 4, , , -1) & "</td>" 'Price
            response.write "<td></td>"
            response.write "<td align=""right"">" & CStr(stkTot) & "</td>"
            response.write "<td align=""right"">" & FormatNumber(tot, 2, , , -1) & "</td>" 'Price
            response.write "<td align=""right""></td>"
            response.write "<td align=""right""> " & expiryDate & "</td>"
            
            'hrf = "wpgSelectPrintLayout.asp?PositionForTableName=DrugStockLevel&DrugID=" & drg & "&DrugStoreID=" & sto
            'response.write "<td align=""center""><a target=""_Blank"" href=""" & hrf & """>Track</a></td>"
            response.write "</tr>"
            
            response.flush
            .MoveNext
          Loop
          response.write "<tr>"
          response.write "<td>.</td>"
          response.write "<td align=""center"" colspan=""5""><b>TOTALS</b></td>"
          response.write "<td align=""right""><b>" & FormatNumber(gTot, 2, , , -1) & "</b></td>"
          response.write "<td colspan=""2""></td>"
          response.write "</tr>"
          'Paging
'          response.write "<tr>"
'          response.write "<td colspan=""8"">" & pgStr & "</td>"
'          response.write "</tr>"
          response.write "</table>"
        End If
      End With
      'Close Drug List
      rst.Close
      response.write "</table>"
      rst0.MoveNext
    Loop
  End If 'rst0
  rst0.Close
  Set rst = Nothing
  Set rst0 = Nothing
End Sub
Function GetPageUrlInfo()
  Dim arr, ul, num, ky, qKy
  ky = ""
  num = 0
  num = num + 1
  ky = "?PrintPagePos=POSITIONFORRECPAGENUMBER"
  For Each qKy In request.querystring
    If UCase(Trim(qKy)) <> "PRINTPAGEPOS" Then
      num = num + 1
      If num > 1 Then
       ky = ky & "&"
      End If
      ky = ky & qKy & "=" & Trim(request.querystring(qKy))
    End If
  Next
  GetPageUrlInfo = ky
End Function
'GetPrintPagePos
  Function GetPrintPagePos(pgCnt)
    Dim currPos, ot
    ot = 1
    currPos = request("PrintPagePos")
    If IsNumeric(currPos) Then
      If CInt(currPos) <= 0 Then
        ot = 1
      ElseIf CInt(currPos) <= pgCnt Then
        ot = CInt(currPos)
      ElseIf CInt(currPos) > pgCnt Then
        ot = CInt(pgCnt)
      End If
    End If
    GetPrintPagePos = ot
  End Function
Function GetDrugPrice(drg, sto)
    Dim ot, sql, rstDst
    Set rstDst = CreateObject("ADODB.Recordset")
    ot = ""
    sql = "select * from drugpricematrix "
    sql = sql & " where drugstoreid='" & sto & "'"
    sql = sql & " and drugid='" & drg & "'"
    rstDst.open qryPro.FltQry(sql), conn, 3, 4
    If rstDst.RecordCount > 0 Then
      rstDst.MoveFirst
      ot = rstDst.fields("itemunitcost")
    End If
    rstDst.Close
    Set rstDst = Nothing
    GetDrugPrice = ot
  End Function
'////////////////////////////////////START SCRIPT //////////////////////////////////
Dim dat1, dat2, prtSql, arr, num, ul, inpFlt, fnd, stoSlt, totalStkVal
Dim arrIn(1000, 10)
Dim arrOut(1000, 10)
Dim cntIn, cntOut, posIn, posOut, PosCur
Dim totIn, totOut, currIn, currOut, avaQty
totalStkVal = 0
server.scripttimeout = 1800
dat1 = CStr(Now())
dat2 = dat1
stoSlt = jSchd 'Trim(GetRecordField("DrugStoreID"))


response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""0"" cellpadding=""0"">"
response.write "<tr>"
AddReportHeader
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
response.write "</tr>"

'Title
response.write "<tr>"
response.write "<td align=""center"" bgcolor=""#FFFFFF"" >"
response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" bgcolor=""White"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:9pt"">"
response.write "<tr>"
response.write "<td>TITLE : </td>"
response.write "<td>DRUGS STOCK TAKE SHEET</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td>STORE : </td>"
response.write "<td>" & UCase(GetComboName("DrugStore", stoSlt)) & " [" & UCase(stoSlt) & "]</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td>DATE : </td>"
response.write "<td> AS AT " & (FormatDateDetail(CDate(dat1))) & "</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td></tr>"


response.write "<tr>"
response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""left"">"
response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 9pt; font-family: Arial"">"
response.write "<tr>"
response.write "<td align=""center"">"

response.write "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
'response.write "<tr><td valign=""top"" colspan=""3""><hr color=""#999999"" size=""1"">"
'response.write "</td></tr>"
response.write "<tr><td valign=""top"" colspan=""3"">"
TrackDrugStock
response.write "</td>"
response.write "</tr>"
response.write "</table>"

response.write "</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
'None
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>

