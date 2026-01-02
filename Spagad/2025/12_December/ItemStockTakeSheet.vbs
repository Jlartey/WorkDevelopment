'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
rackid = Request.QueryString("RackID")
Function GetCurrItemStockLev(sto, drg)
Dim sql, rst, ot
ot = 0
Set rst = CreateObject("ADODB.Recordset")
sql = "select availableqty from Itemstocklevel "
sql = sql & " where Itemid='" & drg & "' and Itemstoreid='" & sto & "'"
With rst
.open sql, conn, 3, 4
If .recordCount > 0 Then
.movefirst
ot = .fields("availableqty")
End If
.Close
End With
GetCurrItemStockLev = ot
Set rst = Nothing
End Function

Sub TrackItemStock()
Dim sql, rst, ot, cnt, rst0, sto, drg, stoNm, drgNm
Dim recPerPg, abPos, recPos, recCnt, num, pgStr, pgUrl
Dim stkTot, stkVal, stkPr, hrf, avaQty, expiryDate, batch
ot = ""
Set rst0 = CreateObject("ADODB.Recordset")
Set rst = CreateObject("ADODB.Recordset")
recPerPg = 5000
  sql = "select * from Itemstore where Itemstoreid='" & stoSlt & "' order by Itemstoreid"
  rst0.open sql, conn, 3, 4
  If rst0.recordCount > 0 Then
    rst0.movefirst
    cnt = 0
    
    Do While Not rst0.EOF
      sto = rst0.fields("ItemStoreID")
      stoNm = rst0.fields("ItemStoreName")
      response.write "<table border=""1"" cellpadding=""3"" cellspacing=""0"" style=""font-size: 8pt; font-family: Arial; border-collapse:collapse;page-break-after:always"">"
      'Open Item List
'       sql = "select Items.Itemname,items.RetailUnitCost,Itemstocklevel.Itemid,itemcategory.itemcategoryname,COALESCE(performvar22.performvar22name, 'No Rack') AS performvar22name,unitofmeasure.unitofmeasurename "
'       sql = sql & " from Itemstocklevel "
'       sql = sql & " inner join Items on itemstocklevel.itemid=items.itemid "
'       sql = sql & " left join performvar24 on items.itemid=performvar24.performvar24name "
'       sql = sql & " left join performvar22 on performvar24.description=performvar22.performvar22id "
'       sql = sql & " inner join itemcategory on items.itemcategoryid=itemcategory.itemcategoryid "
'       sql = sql & " inner join unitofmeasure on items.unitofmeasureid=unitofmeasure.unitofmeasureid "
'       sql = sql & " where items.ItemStatusID <> 'IST002' and Itemstocklevel.Itemstockstatusid='I001'"
'       sql = sql & " and Itemstocklevel.Itemstoreid='" & sto & "'"
' If Len(rackid) > 0 Then
'       sql = sql & " AND performvar22.rack = '" & rackid & "'"
' End If
'       sql = sql & " order by PerformVar22.PerformVar22ID"
       
sql = "SELECT DISTINCT "
sql = sql & "Items.ItemName, "
sql = sql & "ItemStockLevel.itemid, "
sql = sql & "Items.RetailUnitCost, "
sql = sql & "IncomingStockItems.expirydate, "
sql = sql & "ItemCategory.ItemCategoryName, "
sql = sql & "COALESCE(performvar22.performvar22name, 'No Rack') AS performvar22name, "
sql = sql & "unitofmeasure.unitofmeasurename, "
sql = sql & "IncomingStockItems.PurchaseOrderInfo1, "
sql = sql & "performvar22.performvar22id "
sql = sql & "from ItemStockLevel "
sql = sql & "inner join Items on ItemStockLevel.itemid=items.itemid "
sql = sql & "LEFT join performvar24 on items.itemid=performvar24.performvar24name "
sql = sql & "left join performvar22 on performvar24.description=performvar22.performvar22id "
sql = sql & "inner join itemcategory on items.itemcategoryid=itemcategory.itemcategoryid "
sql = sql & "inner join unitofmeasure on items.unitofmeasureid=unitofmeasure.unitofmeasureid "
sql = sql & "LEFT JOIN IncomingStockItems ON IncomingStockItems.ItemID = ItemStockLevel.ItemID "
sql = sql & "where items.ItemStatusID <> 'IST002' and Itemstocklevel.Itemstockstatusid='I001' "
sql = sql & "and Itemstocklevel.Itemstoreid ='" & sto & "' "
If Len(rackid) > 0 Then
      sql = sql & " AND performvar22.rack = '" & rackid & "'"
End If
sql = sql & "order by PerformVar22.PerformVar22ID, Itemname ASC, PurchaseOrderInfo1"
 
 

dispOption "S20"
      
      rst.CacheSize = CInt(recPerPg)
      rst.open sql, conn, 3, 4
      With rst
        If .recordCount > 0 Then
          .PageSize = CInt(recPerPg)
          'Generate Paging Links
          pgUrl = GetPageUrlInfo()
          pgStr = "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size: 8pt; font-weight:bold; border-collapse:collapse""><tr>"
          pgStr = pgStr & "<td style=""font-size: 9pt; font-weight:normal"">FIND&nbsp;:&nbsp;</td>"
          For num = 1 To .PageCount
            .AbsolutePage = num
            pgStr = pgStr & "<td><a href=""" & Replace(pgUrl, "POSITIONFORRECPAGENUMBER", CStr(num)) & """>" & Left(.fields("ItemName"), 3) & "</a></td>"
          Next
          pgStr = pgStr & "</tr></table>"
          
If Len(rackid) > 0 Then
          response.write "<tr>"
          response.write "<th colspan=""14"" align=""center"">" & UCase(rackid) & "</th>"
          response.write "</tr>"
End If
          response.write "<tr>"
          response.write "<td><b>NO.</b></td>"
          response.write "<td><b>ITEMID</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>ITEM</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>U.PRICE</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>UOM</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>CATEGORY</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>BATCH</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>RACK</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>CURR. LEV</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>BIN CARD NO.</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>ADJUSTMENT</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>EXPIRY DATE</b></td>"
          response.write "<td valign=""top"" align=""Center"" width=""140""><b>STOCK TAKE</b></td>"
          response.write "<td valign=""top"" align=""Center""><b>REMARKS</b></td>"
          '
          
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
            drg = .fields("Itemid")
            drgNm = .fields("ItemName")
            cat = .fields("itemcategoryname")
            rck = .fields("performvar22name")
            uom = .fields("unitofmeasurename")
            expiryDate = .fields("expirydate")
            batch = .fields("PurchaseOrderInfo1")
            
            avaQty = GetCurrItemStockLev(sto, drg)
            stkTot = avaQty
            stkPr = .fields("RetailUnitCost") ' GetComboNameFld("Items", drg, "RetailUnitCost") 'GetItemPrice(drg, sto)
            'stkVal = CDbl(stkTot) * CDbl(stkPr)
            'totalStkVal = totalStkVal + stkVal
            response.write "<tr>"
            response.write "<td><b>" & CStr(recPos) & "</b></td>"
            response.write "<td>" & UCase(drg) & "</td>"
            response.write "<td>" & UCase(drgNm) & "</td>"
            response.write "<td align=""right"">" & FormatNumber(stkPr, 4, , , -1) & "</td>" 'Price
            response.write "<td align=""left"">" & UCase(uom) & "</td>"
            response.write "<td align=""left"">" & UCase(cat) & "</td>"
            response.write "<td align=""left"">" & UCase(batch) & "</td>"
            response.write "<td align=""left"">" & UCase(rck) & "</td>"
            response.write "<td align=""right"">" & CStr(stkTot) & "</td>"
            response.write "<td align=""right""></td>" 'Bin card
            response.write "<td align=""right""></td>" 'Adjustment
            response.write "<td align=""right"">" & UCase(expiryDate) & "</td>"
            response.write "<td align=""right""></td>"
            response.write "<td align=""right""></td>"
            response.write "</tr>"
            
            response.flush
            .MoveNext
          Loop
          
          response.write "</table>"
        End If
      End With
      'Close Item List
      rst.Close
      response.write "</table>"
      rst0.MoveNext
    Loop
  End If 'rst0
  rst0.Close
  Set rst = Nothing
  Set rst0 = Nothing
End Sub
Sub dispOption(strTable)
    Dim sql1, rst1, str
    Set rst1 = CreateObject("ADODB.RecordSet")
    sql1 = "SELECT distinct rack "
    sql1 = sql1 & " FROM PerformVar22 WHERE description='active' AND keyprefix='S20'"
    
'    response.write sql
    
    response.write "<style>"
    response.write ".cta-form{"
    response.write "  margin-bottom: 25px;"
    response.write "  padding: 6px;"
    response.write "  font-size: 15px;"
    response.write "  font-family: inherit;"
    response.write "  color: inherit;"
    response.write "  border: none;"
    response.write "  background-color: blanchedalmond;"
    response.write "  border-radius: 9px;"
    response.write "  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);"
    response.write "}"
    response.write "</style>"
    
    With rst1
        .open qryPro.FltQry(sql1), conn, 3, 4
        If .recordCount > 0 Then
            response.write " <select name='RackID' id='RackID' class='cta-form'> "
            response.write "   <option disabled selected hidden> Select Rack </option> "
            .movefirst
                Do While Not .EOF
                    rckID = .fields("rack")
                    'rckName = UCase(.fields("PerformVar22Name"))
                    response.write "   <option value='" & rckID & "'> " & rckID & " </option> "
                    .MoveNext
                Loop
            response.write " </select> "
        End If
        .Close
    End With
    
    response.write vbCrLf & "<script>"
    response.write vbCrLf & "   const InpOption = document.getElementById('RackID')"
    response.write vbCrLf & "   InpOption.addEventListener('change', updateURL);"
    response.write vbCrLf & "   function updateURL(){"
    response.write vbCrLf & "       let currURL = new URL(window.location.href);"
    response.write vbCrLf & "       let params = new URLSearchParams(currURL.search);"
    response.write vbCrLf & "       params.set('RackID', InpOption.value);"
    response.write vbCrLf & "       currURL.search = params.toString();"
    response.write vbCrLf & "       window.location.href = currURL.toString();"
'    response.write vbCrLf & "       console.log(currURL.toString());"
    response.write vbCrLf & "   };"
    response.write vbCrLf & "</script>"
    
End Sub
Function GetPageUrlInfo()
  Dim arr, ul, num, ky, qKy
  ky = ""
  num = 0
  num = num + 1
  ky = "?PrintPagePos=POSITIONFORRECPAGENUMBER"
  For Each qKy In Request.QueryString
    If UCase(Trim(qKy)) <> "PRINTPAGEPOS" Then
      num = num + 1
      If num > 1 Then
       ky = ky & "&"
      End If
      ky = ky & qKy & "=" & Trim(Request.QueryString(qKy))
    End If
  Next
  GetPageUrlInfo = ky
End Function
'GetPrintPagePos
  Function GetPrintPagePos(pgCnt)
    Dim currPos, ot
    ot = 1
    currPos = Request("PrintPagePos")
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
Function GetItemPrice(drg, sto)
    Dim ot, sql, rstDst
    Set rstDst = CreateObject("ADODB.Recordset")
    ot = ""
    sql = "select * from Itempricematrix "
    sql = sql & " where Itemstoreid='" & sto & "'"
    sql = sql & " and Itemid='" & drg & "'"
    rstDst.open sql, conn, 3, 4
    If rstDst.recordCount > 0 Then
      rstDst.movefirst
      ot = rstDst.fields("itemunitcost")
    End If
    rstDst.Close
    Set rstDst = Nothing
    GetItemPrice = ot
  End Function
'////////////////////////////////////START SCRIPT //////////////////////////////////
Dim dat1, dat2, prtSql, arr, num, ul, inpFlt, fnd, stoSlt, totalStkVal
Dim arrIn(1000, 10)
Dim arrOut(1000, 10)
Dim cntIn, cntOut, posIn, posOut, PosCur
Dim totIn, totOut, currIn, currOut, avaQty
totalStkVal = 0

dat1 = CStr(Now())
dat2 = dat1
stoSlt = jSchd ' Trim(GetRecordField("ItemStoreID"))



response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table class='mytable' border=""0""  width=""95vw"" cellspacing=""0"" cellpadding=""0"">"
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
response.write "<td>ITEMS STOCK TAKE SHEET</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td>STORE : </td>"
response.write "<td>" & UCase(GetComboName("ItemStore", stoSlt)) & " [" & UCase(stoSlt) & "]</td>"
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
TrackItemStock
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

