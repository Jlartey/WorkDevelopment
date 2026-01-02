Sub TrackDrugStock()
Dim sql, rst, ot, cnt, rst0, sto, drg, stoNm, drgNm
Dim recPerPg, abPos, recPos, recCnt, num, pgStr, pgUrl
Dim stkTot, stkVal, stkPr, hrf, avaQty, tot, gTot, batch, expiryDate

ot = ""
Set rst0 = CreateObject("ADODB.Recordset")
Set rst = CreateObject("ADODB.Recordset")
recPerPg = 5000

' === ADD CSS FOR PRINT STYLING ===
response.write "<style type=""text/css"">"
response.write "@page { size: A4 landscape; margin: 10mm; }"
response.write "@media print {"
response.write "  body { font-family: Arial, sans-serif; font-size: 9pt; }"
response.write "  table { page-break-inside: avoid; }"
response.write "  thead { display: table-header-group; }"
response.write "  tr { page-break-inside: avoid; page-break-after: auto; }"
response.write "  .no-print { display: none; }"
response.write "}"
response.write "table.stocktable {"
response.write "  width: 100%; border-collapse: collapse; font-size: 9pt; font-family: Arial, Helvetica, sans-serif;"
response.write "}"
response.write "table.stocktable th {"
response.write "  border: 1px solid #000; padding: 8px; text-align: center; background-color: #f0f0f0; font-weight: bold; font-size: 9pt;"
response.write "}"
response.write "table.stocktable td {"
response.write "  border: 1px solid #000; padding: 8px; vertical-align: middle; height: 40px;"
response.write "}"
response.write "table.stocktable td.right { text-align: right; }"
response.write "table.stocktable td.center { text-align: center; }"
response.write "table.stocktable tr:nth-child(even) { background-color: #f9f9f9; }"
response.write ".wide-col { width: 180px; } /* For Stock Take and Remarks */"
response.write ".batch-col { width: 120px; }"
response.write ".rack-title { font-size: 16pt; font-weight: bold; text-align: center; padding: 15px; background-color: #e0e0e0; }"
response.write "</style>"

  sql = "select * from drugstore where drugstoreid='" & stoSlt & "' order by drugstoreid"
  rst0.open sql, conn, 3, 4
  If rst0.recordCount > 0 Then
    rst0.movefirst
    cnt = 0
    tot = 0
    gTot = 0
    Do While Not rst0.EOF
      sto = rst0.fields("DrugStoreID")
      stoNm = rst0.fields("DrugStoreName")

      sql = "SELECT DISTINCT "
      sql = sql & "drug.drugname, "
      sql = sql & "drugstocklevel.drugid, "
      sql = sql & "drug.RetailUnitCost, "
      sql = sql & "IncomingDrugItems.expirydate, "
      sql = sql & "drugtype.drugtypename, "
      sql = sql & "COALESCE(performvar22.performvar22name, 'No Rack') AS performvar22name, "
      sql = sql & "unitofmeasure.unitofmeasurename, "
      sql = sql & "IncomingDrugItems.PurchaseOrderInfo1, "
      sql = sql & "performvar22.performvar22id "
      sql = sql & "FROM drugstocklevel "
      sql = sql & "INNER JOIN drug ON drugstocklevel.drugid = drug.drugid "
      sql = sql & "LEFT JOIN performvar24 ON drug.drugid = performvar24.performvar24name "
      sql = sql & "LEFT JOIN performvar22 ON performvar24.description = performvar22.performvar22id "
      sql = sql & "INNER JOIN drugtype ON drug.drugtypeid = drugtype.drugtypeid "
      sql = sql & "INNER JOIN unitofmeasure ON unitofmeasure.unitofmeasureid = drug.unitofmeasureid "
      sql = sql & "LEFT JOIN IncomingDrugItems ON IncomingDrugItems.DrugID = DrugStockLevel.DrugID "
      sql = sql & "WHERE drug.DrugStatusID <> 'IST002' AND drug.drugcategoryid <> 'D002' "
      sql = sql & "AND drugstocklevel.drugstockstatusid = 'D001' "
      sql = sql & "AND drugstocklevel.drugstoreid = '" & sto & "' "
      If Len(rackid) > 0 Then
          sql = sql & "AND performvar22.rack = '" & rackid & "' "
      End If
      sql = sql & "ORDER BY performvar22.performvar22id, drug.drugname, PurchaseOrderInfo1"

      rst.open sql, conn, 3, 4

      If rst.recordCount > 0 Then
        ' === RACK TITLE (if filtered) ===
        If Len(rackid) > 0 Then
          response.write "<div class=""rack-title"">RACK: " & UCase(rackid) & "</div>"
        End If

        ' === MAIN STOCK TAKE TABLE ===
        response.write "<table class=""stocktable"" border=""1"">"
        response.write "<thead>"
        response.write "<tr>"
        response.write "<th style=""width:40px"">NO.</th>"
        response.write "<th style=""width:80px"">DRUG ID</th>"
        response.write "<th style=""width:220px"">DRUG NAME</th>"
        response.write "<th style=""width:80px"">UNIT PRICE</th>"
        response.write "<th style=""width:60px"">UOM</th>"
        response.write "<th style=""width:80px"">TYPE</th>"
        response.write "<th class=""batch-col"">BATCH NO.</th>"
        response.write "<th style=""width:100px"">RACK</th>"
        response.write "<th style=""width:80px"">CURRENT LEVEL</th>"
        response.write "<th style=""width:90px"">TOTAL VALUE</th>"
        response.write "<th style=""width:90px"">BIN CARD</th>"
        response.write "<th style=""width:70px"">ADJ.</th>"
        response.write "<th style=""width:90px"">EXPIRY</th>"
        response.write "<th class=""wide-col"">STOCK TAKE</th>"
        response.write "<th class=""wide-col"">REMARKS</th>"
        response.write "</tr>"
        response.write "</thead>"
        response.write "<tbody>"

        abPos = GetPrintPagePos(rst.PageCount)
        rst.AbsolutePage = abPos
        recPos = (abPos - 1) * recPerPg
        recCnt = 0

        Do While Not rst.EOF And recCnt < recPerPg
          recPos = recPos + 1
          drg = rst.fields("drugid")
          drgNm = rst.fields("drugname")
          typ = rst.fields("drugtypename")
          rck = rst.fields("performvar22name")
          uom = rst.fields("unitofmeasurename")
          batch = Nz(rst.fields("PurchaseOrderInfo1"), "")
          expiryDate = Nz(rst.fields("expirydate"), "")

          avaQty = GetCurrDrugStockLev(sto, drg)
          stkPr = CDbl(Nz(rst.fields("RetailUnitCost"), 0))
          tot = avaQty * stkPr
          gTot = gTot + tot

          response.write "<tr>"
          response.write "<td class=""center""><b>" & recPos & "</b></td>"
          response.write "<td class=""center"">" & UCase(drg) & "</td>"
          response.write "<td>" & UCase(drgNm) & "</td>"
          response.write "<td class=""right"">" & FormatNumber(stkPr, 4) & "</td>"
          response.write "<td class=""center"">" & UCase(uom) & "</td>"
          response.write "<td>" & UCase(typ) & "</td>"
          response.write "<td>" & UCase(batch) & "</td>"
          response.write "<td>" & UCase(rck) & "</td>"
          response.write "<td class=""right"">" & avaQty & "</td>"
          response.write "<td class=""right"">" & FormatNumber(tot, 2) & "</td>"
          response.write "<td></td>"  ' BIN CARD
          response.write "<td></td>"  ' ADJUSTMENT
          response.write "<td class=""center"">" & expiryDate & "</td>"
          response.write "<td></td>"  ' STOCK TAKE - wide empty cell
          response.write "<td></td>"  ' REMARKS - wide empty cell
          response.write "</tr>"

          rst.MoveNext
          recCnt = recCnt + 1
        Loop

        ' === TOTALS ROW ===
        response.write "<tr style=""background-color: #e0e0e0; font-weight: bold;"">"
        response.write "<td colspan=""9"" class=""right"">GRAND TOTAL:</td>"
        response.write "<td class=""right"">" & FormatNumber(gTot, 2) & "</td>"
        response.write "<td colspan=""5""></td>"
        response.write "</tr>"

        response.write "</tbody>"
        response.write "</table>"

        ' Add space between racks/stores
        response.write "<div style=""page-break-after: always;""></div>"
      End If

      rst.Close
      rst0.MoveNext
    Loop
  End If

  rst0.Close
  Set rst = Nothing
  Set rst0 = Nothing
End Sub