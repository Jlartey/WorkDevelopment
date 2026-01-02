Dim VisitID, SponsorID, chkedID, splitKeyID, table, chkQty
Dim drgs, itms, labs, trts, cslt, beds, diss, performArr1, performArr2
VisitID = Trim(Request.QueryString("VisitID"))
vst = Trim(Request.QueryString("vst"))
chkedID = Trim(Request.QueryString("chkedID"))
chkQty = Trim(Request.QueryString("chkQty"))
proformaStat = Trim(Request.QueryString("proformaStat"))
' chkQty = CInt("2")
drgs = Trim(Request.QueryString("drugs"))
itms = Trim(Request.QueryString("items"))
labs = Trim(Request.QueryString("labs"))
trts = Trim(Request.QueryString("trts"))
srhDrug = Trim(Request.QueryString("srhDrg"))
srhItem = Trim(Request.QueryString("srhItm"))
srhLab = Trim(Request.QueryString("srhLab"))
srhTreat = Trim(Request.QueryString("srhTreat"))
splitKeyID = Split(chkedID, "-")
performArr1 = ""
chked = ""
table = ""
insType = GetComboNameFld("Visitation", vst, "InsuranceTypeID")
If UBound(splitKeyID) > 0 Then
    performArr1 = splitKeyID(1)
    performArr2 = Split(performArr1, "||")
    chked = performArr2(0)
    table = splitKeyID(0)
    cst = CDbl(performArr2(1))
    cQty = CDbl(chkQty)
    eql = (cst * cQty)
    KeyPrefix = "dhd"
    desc = chked & "||" & eql
End If

response.ContentType = "text/html"

If VisitID <> "" And chked <> "" And table = "drg" Then
  ProcessDrug VisitID, eql, desc
ElseIf VisitID <> "" And chked <> "" And table = "itm" Then
  ProcessItems VisitID, eql
ElseIf VisitID <> "" And chked <> "" And table = "lab" Then
  ProcessLabs VisitID, eql
ElseIf VisitID <> "" And chked <> "" And table = "trt" Then
  ProcessTreatments VisitID, eql
End If

If Len(drgs) > 1 Then
  ListDrugs vst, desc
End If
If Len(itms) > 1 Then
  ListItems vst
End If
If Len(labs) > 1 Then
  ListLabs vst
End If
If Len(trts) > 1 Then
  ListTreats vst
End If

If Len(srhDrug) > 0 Then
  ListSrhDrugs vst, srhDrug, chkQty
ElseIf Len(srhItem) > 0 Then
  ListSrhItems vst, srhItem
ElseIf Len(srhLab) > 0 Then
  ListSrhLabs vst, srhLab
ElseIf Len(srhTreat) > 0 Then
  ListSrhTreats vst, srhTreat
End If

If Trim(proformaStat) <> "" Then
  UpdateVisitationFld proformaStat
End If

Sub ProcessDrug(VisitID, eql, desc)
  Dim rst, rst2, rst3, sql, sql2, sql3, isNew, exempt, tmp, tmp2, uniqueItemID
  isNew = False
  sql = "Select * from PerformVAr59 where "
  sql = sql & " PerformVar59Name='" & VisitID & "' AND Description = '" & performArr1 & "'"
  sql3 = "SELECT * FROM drug WHERE drugid='" & chked & "' AND DrugStatusid='IST001'"

  Set rst = CreateObject("ADODB.RecordSet")
  Set rst3 = CreateObject("ADODB.RecordSet")
  Set resp = CreateObject("Scripting.Dictionary")
    rst3.open qryPro.FltQry(sql3), conn, 3, 4
    If rst3.RecordCount > 0 Then
      rst3.MoveFirst
      ' check Exemption
      rst.open qryPro.FltQry(sql), conn, 3, 4
      If rst.RecordCount = 0 Then
      uniqueItemID = GenerateUniqueID
      rst.AddNew
      rst.fields("PerformVar59ID") = uniqueItemID
      rst.fields("PerformVar59Name") = VisitID
      rst.fields("Description") = performArr1
        rst.fields("KeyPrefix") = "DRUG" & "||" & chkQty
      rst.updatebatch
      resp("status1") = "SAVED"
      Else
      conn.execute qryPro.FltQry("DELETE FROM PerformVAr59 WHERE PerformVar59Name='" & VisitID & "' AND Description = '" & performArr1 & "'")
      resp("status1") = "DELETED"
      End If
    End If
    rst.Close
    rst3.Close
    
  If IsObject(response) Then
    response.Clear
    response.write gUtils.JSONStringify(resp)
    response.Flush
  End If
End Sub

Sub ProcessItems(VisitID, eql)
  Dim rst, rst2, rst3, sql, sql2, sql3, isNew, exempt, tmp, tmp2, uniqueItemID
  isNew = False
  sql = "Select * from PerformVAr59 where "
  sql = sql & " PerformVar59Name='" & VisitID & "' AND Description = '" & performArr1 & "'"
  sql3 = "SELECT * FROM Items WHERE itemid='" & chked & "' AND itemStatusid='IST001'"
  response.write sql3

  Set rst = CreateObject("ADODB.RecordSet")
  Set rst3 = CreateObject("ADODB.RecordSet")
  Set resp = CreateObject("Scripting.Dictionary")
  
  rst3.open qryPro.FltQry(sql3), conn, 3, 4
  If rst3.RecordCount > 0 Then
    rst3.MoveFirst
    ' check Exemption
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount = 0 Then
      uniqueItemID = GenerateUniqueID
      rst.AddNew
      rst.fields("PerformVar59ID") = uniqueItemID
      rst.fields("PerformVar59Name") = VisitID
      rst.fields("Description") = performArr1
      rst.fields("KeyPrefix") = "ITEMS" & "||" & chkQty
      rst.updatebatch
      resp("status1") = "SAVED"
    Else
      conn.execute qryPro.FltQry("DELETE FROM PerformVAr59 WHERE PerformVar59Name='" & VisitID & "' AND Description = '" & performArr1 & "'")
      resp("status1") = "DELETED"
    End If
  End If
  rst.Close
  rst3.Close
  
  If IsObject(response) Then
    response.Clear
    response.write gUtils.JSONStringify(resp)
    response.Flush
  End If
End Sub

Sub ProcessLabs(VisitID, eql)
  Dim rst, rst2, rst3, sql, sql2, sql3, isNew, exempt, tmp, tmp2, uniqueItemID
  isNew = False
  sql = "Select * from PerformVAr59 where "
  sql = sql & " PerformVar59Name='" & VisitID & "' AND Description = '" & performArr1 & "'"
  sql3 = "SELECT * FROM LabTest WHERE LabTestid='" & chked & "' AND TestStatusid='TST001'"

  Set rst = CreateObject("ADODB.RecordSet")
  Set rst3 = CreateObject("ADODB.RecordSet")
  Set resp = CreateObject("Scripting.Dictionary")
    ' check labs
    rst3.open qryPro.FltQry(sql3), conn, 3, 4
    If rst3.RecordCount > 0 Then
      rst3.MoveFirst
      ' check Exemption
      rst.open qryPro.FltQry(sql), conn, 3, 4
      If rst.RecordCount = 0 Then
        uniqueItemID = GenerateUniqueID
        rst.AddNew
        rst.fields("PerformVar59ID") = uniqueItemID
        rst.fields("PerformVar59Name") = VisitID
        rst.fields("Description") = performArr1
        rst.fields("KeyPrefix") = "LABTEST" & "||" & chkQty
        rst.updatebatch
        resp("status1") = "SAVED"
      Else
      conn.execute qryPro.FltQry("DELETE FROM PerformVAr59 WHERE PerformVar59Name='" & VisitID & "' AND Description = '" & performArr1 & "'")
        resp("status1") = "DELETED"
      End If
    End If
    rst.Close
    rst3.Close
  
  If IsObject(response) Then
    response.Clear
    response.write gUtils.JSONStringify(resp)
    response.Flush
  End If
End Sub

Sub ProcessTreatments(VisitID, eql)
  Dim rst, rst2, rst3, sql, sql2, sql3, isNew, exempt, tmp, tmp2, uniqueItemID
  isNew = False
  sql = "Select * from PerformVAr59 where "
  sql = sql & " PerformVar59Name='" & VisitID & "' AND Description = '" & performArr1 & "'"
  sql3 = "SELECT * FROM Treatment WHERE Treatmentid='" & chked & "' AND TreatInfo1='NO'"

  Set rst = CreateObject("ADODB.RecordSet")
  Set rst3 = CreateObject("ADODB.RecordSet")
  Set resp = CreateObject("Scripting.Dictionary")
    rst3.open qryPro.FltQry(sql3), conn, 3, 4
    If rst3.RecordCount > 0 Then
      rst3.MoveFirst
      ' check Exemption
      rst.open qryPro.FltQry(sql), conn, 3, 4
      If rst.RecordCount = 0 Then
        uniqueItemID = GenerateUniqueID
        rst.AddNew
        rst.fields("PerformVar59ID") = uniqueItemID
        rst.fields("PerformVar59Name") = VisitID
        rst.fields("Description") = performArr1
          rst.fields("KeyPrefix") = "TREATMENT" & "||" & chkQty
        rst.updatebatch
        resp("status1") = "SAVED"
      Else
      conn.execute qryPro.FltQry("DELETE FROM PerformVAr59 WHERE PerformVar59Name='" & VisitID & "' AND Description = '" & performArr1 & "'")
        resp("status1") = "DELETED"
      End If
    End If
    rst.Close
    rst3.Close
  
  If IsObject(response) Then
    response.Clear
    response.write gUtils.JSONStringify(resp)
    response.Flush
  End If
End Sub

Sub ListSrhDrugs(VisitID, srhDrug, chkQty)
  Dim sql, sql2, ot, rst, rst2, cnt, cost, drgQtyArr, drgQty, drgQty1
  Set rst = CreateObject("ADODB.Recordset")
  Set rst2 = CreateObject("ADODB.Recordset")
  'srhDisease
  sql = "SELECT TOP 100 drug.drugid, drugpricematrix2.itemunitcost"
  sql = sql & " FROM Drug"
  sql = sql & " join drugpricematrix2 on drugpricematrix2.drugid=drug.drugid"
  sql = sql & " WHERE drug.DrugName LIKE '%" & srhDrug & "%' AND drug.drugstatusid='IST001'"
  sql = sql & " and drugpricematrix2.insurancetypeid='" & insType & "'"
  cnt = 0
  drgQty = 1
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      response.write "<table id='myTable' border='1'>"
      response.write "<tr><th>No.</th><th>ID</th><th>DrugName</th><th>Quantity</th><th>UnitCost</th><th style='width:180px!important;'>Status</th></tr>"
      Do While Not .EOF
        cnt = cnt + 1
        drugId = .fields("drugid")
        cost = .fields("itemunitcost")
        dsc = drugId & "||" & cost
        sql2 = "SELECT * from PerformVar59 WHERE PerformVar59Name = '" & VisitID & "' AND Description = '" & dsc & "'"
        With rst2
            .open qryPro.FltQry(sql2), conn, 3, 4
            If .RecordCount > 0 Then
                .MoveFirst
                drgQtyArr = Split(.fields("KeyPrefix"), "||")
                drgQty1 = drgQtyArr(1)
                response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & drugId & "</td><td>" & GetComboName("Drug", drugId) & "</td><td><input type=""number"" id='drg-" & drugId & "' class=""Quantity"" name=""Quantity"" value=" & drgQty1 & " min=""1""></td><td>" & (FormatNumber(CStr(cost), 2, , , -1)) & "</td><td>" & BuildDrugCheckBtn(drugId, VisitID, cost, desc) & "</td></tr>"
            Else
                response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & drugId & "</td><td>" & GetComboName("Drug", drugId) & "</td><td><input type=""number"" id='drg-" & drugId & "' class=""Quantity"" name=""Quantity"" value=" & drgQty & " min=""1""></td><td>" & (FormatNumber(CStr(cost), 2, , , -1)) & "</td><td>" & BuildDrugCheckBtn(drugId, VisitID, cost, desc) & "</td></tr>"
            End If
            .Close
        End With
      .MoveNext
      Loop
      response.write "</table>"
    Else
      response.write "<tr><td colspan='4'><h3 style='color:red;'>No Drugs of Such Name</h3></td></tr>"
      response.write "</table>"
    End If
  .Close
  End With
  Set rst = Nothing
End Sub

Sub ListSrhItems(VisitID, srhItem)
  Dim sql, sql2, ot, rst, rst2, cnt, cost, itmQtyArr, itmQty, itmQty1
  Set rst = CreateObject("ADODB.Recordset")
  Set rst2 = CreateObject("ADODB.Recordset")
  'srhDisease
  sql = "SELECT TOP 100 items.itemid, itempricematrix2.itemunitcost FROM Items join itempricematrix2 ON itempricematrix2.itemid=items.itemid WHERE items.ItemName LIKE '%" & srhItem & "%' AND items.itemStatusid='IST001' and itempricematrix2.insurancetypeid='" & insType & "'"
  cnt = 0
  itmQty = 1
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      response.write "<table id='myTable' border='1'>"
      response.write "<tr><th>No.</th><th>ID</th><th>Name</th><th>Quantity</th><th>UnitCost</th><th style='width:180px!important;'>Status</th></tr>"
      Do While Not .EOF
        cnt = cnt + 1
        itemID = .fields("itemid")
        cost = .fields("itemunitcost")
        dsc = itemID & "||" & cost
        sql2 = "SELECT * from PerformVar59 WHERE PerformVar59Name = '" & VisitID & "' AND Description = '" & dsc & "'"
        With rst2
            .open qryPro.FltQry(sql2), conn, 3, 4
            If .RecordCount > 0 Then
                .MoveFirst
                itmQtyArr = Split(.fields("KeyPrefix"), "||")
                itmQty1 = itmQtyArr(1)
                response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & itemID & "</td><td>" & GetComboName("items", itemID) & "</td><td><input type=""number"" id='itm-" & itemID & "' class=""Quantity"" name=""Quantity"" value=" & itmQty1 & " min=""1""></td><td>" & (FormatNumber(CStr(cost), 2, , , -1)) & "</td><td>" & BuildItemCheckBtn(itemID, VisitID, cost) & "</td></tr>"
            Else
                response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & itemID & "</td><td>" & GetComboName("items", itemID) & "</td><td><input type=""number"" id='itm-" & itemID & "' class=""Quantity"" name=""Quantity"" value=" & itmQty & " min=""1""></td><td>" & (FormatNumber(CStr(cost), 2, , , -1)) & "</td><td>" & BuildItemCheckBtn(itemID, VisitID, cost) & "</td></tr>"
            End If
            .Close
        End With
      .MoveNext
      Loop
      response.write "</table>"
    Else
      response.write "<tr><td colspan='4'><h3 style='color:red;'>No Items of Such Name</h3></td></tr>"
      response.write "</table>"
    End If
  .Close
  End With
  Set rst = Nothing
End Sub

Sub ListSrhLabs(VisitID, srhLab)
  Dim sql, ot, rst, cnt, cost, labQtyArr, labQty, labQty1, rst2, sql2
  Set rst = CreateObject("ADODB.Recordset")
  Set rst2 = CreateObject("ADODB.Recordset")
  'srhDisease
  sql = "SELECT TOP 100 LabTest.labtestid, LabTestCostMatrix.unitcost FROM LabTest join LabTestCostMatrix on LabTestCostMatrix.labtestid=labtest.labtestid WHERE LabTest.LabTestName LIKE '%" & srhLab & "%' AND LabTest.TestStatusid='TST001' and LabTestCostMatrix.insurancetypeid='" & insType & "' and LabTestCostMatrix.agegroupid='A002'"
  cnt = 0
  labQty = 1
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      response.write "<table id='myTable' border='1'>"
      response.write "<tr><th>No.</th><th>ID</th><th>Name</th><th>Quantity</th><th>UnitCost</th><th style='width:180px!important;'>Status</th></tr>"
      Do While Not .EOF
        cnt = cnt + 1
        labID = .fields("labtestid")
        cost = .fields("unitcost")
        dsc = labID & "||" & cost
        sql2 = "SELECT * from PerformVar59 WHERE PerformVar59Name = '" & VisitID & "' AND Description = '" & dsc & "'"
        With rst2
            .open qryPro.FltQry(sql2), conn, 3, 4
            If .RecordCount > 0 Then
                .MoveFirst
                labQtyArr = Split(.fields("KeyPrefix"), "||")
                labQty1 = labQtyArr(1)
                response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & labID & "</td><td>" & GetComboName("LabTest", labID) & "</td><td><input type=""number"" id='lab-" & labID & "' class=""Quantity"" name=""Quantity"" value=" & labQty1 & " min=""1""></td><td>" & (FormatNumber(CStr(cost), 2, , , -1)) & "</td><td>" & BuildLabCheckBtn(labID, VisitID, cost) & "</td></tr>"
            Else
                response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & labID & "</td><td>" & GetComboName("LabTest", labID) & "</td><td><input type=""number"" id='lab-" & labID & "' class=""Quantity"" name=""Quantity"" value=" & labQty & " min=""1""></td><td>" & (FormatNumber(CStr(cost), 2, , , -1)) & "</td><td>" & BuildLabCheckBtn(labID, VisitID, cost) & "</td></tr>"
            End If
            .Close
        End With
      .MoveNext
      Loop
      response.write "</table>"
    Else
      response.write "<tr><td colspan='4'><h3 style='color:red;'>No Labs of Such Name</h3></td></tr>"
      response.write "</table>"
    End If
  .Close
  End With
  Set rst = Nothing
End Sub

Sub ListSrhTreats(VisitID, srhTreat)
  Dim sql, ot, rst, cnt, cost, trtQtyArr, trtQty, trtQty1, rst2, sql2
  Set rst = CreateObject("ADODB.Recordset")
  Set rst2 = CreateObject("ADODB.Recordset")
  'srhDisease
  sql = "SELECT TOP 100 Treatment.treatmentid, TreatCostMatrix.UnitCost FROM Treatment join TreatCostMatrix on TreatCostMatrix.treatmentid=treatment.treatmentid WHERE Treatment.TreatmentName LIKE '%" & srhTreat & "%' AND Treatment.TreatInfo1='NO' and TreatCostMatrix.insurancetypeid='" & insType & "' and TreatCostMatrix.agegroupid='A002' AND TreatCostMatrix.medicalserviceid='M003'"
  cnt = 0
  trtQty = 1
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      response.write "<table id='myTable' border='1'>"
      response.write "<tr><th>No.</th><th>ID</th><th>Name</th><th>Quantity</th><th>UnitCost</th><th style='width:180px!important;'>Status</th></tr>"
      Do While Not .EOF
        cnt = cnt + 1
        treatID = .fields("treatmentid")
        cost = .fields("unitcost")
        dsc = treatID & "||" & cost
        sql2 = "SELECT * from PerformVar59 WHERE PerformVar59Name = '" & VisitID & "' AND Description = '" & dsc & "'"
        With rst2
            .open qryPro.FltQry(sql2), conn, 3, 4
            If .RecordCount > 0 Then
                .MoveFirst
                trtQtyArr = Split(.fields("KeyPrefix"), "||")
                trtQty1 = trtQtyArr(1)
                response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & treatID & "</td><td>" & GetComboName("Treatment", treatID) & "</td><td><input type=""number"" id='trt-" & treatID & "' class=""Quantity"" name=""Quantity"" value=" & trtQty1 & " min=""1""></td><td>" & (FormatNumber(CStr(cost), 2, , , -1)) & "</td><td>" & BuildTrtCheckBtn(treatID, VisitID, cost) & "</td></tr>"
            Else
                response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & treatID & "</td><td>" & GetComboName("Treatment", treatID) & "</td><td><input type=""number"" id='trt-" & treatID & "' class=""Quantity"" name=""Quantity"" value=" & trtQty & " min=""1""></td><td>" & (FormatNumber(CStr(cost), 2, , , -1)) & "</td><td>" & BuildTrtCheckBtn(treatID, VisitID, cost) & "</td></tr>"
            End If
            .Close
        End With
      .MoveNext
      Loop
      response.write "</table>"
    Else
      response.write "<tr><td colspan='4'><h3 style='color:red;'>No Procedures of Such Name</h3></td></tr>"
      response.write "</table>"
    End If
  .Close
  End With
  Set rst = Nothing
  Set rst2 = Nothing
End Sub

Sub ListDrugs(VisitID, chkQty)
  Dim rst, rst2, ot, sql, drugId, cnt, price, drgQtyArr, drgQty, drgQty1
  Set rst = server.CreateObject("ADODB.Recordset")
  Set rst2 = server.CreateObject("ADODB.Recordset")
  sql = "SELECT TOP 100 drug.drugid, drugpricematrix2.ItemUnitCost FROM drug join drugpricematrix2 on drugpricematrix2.drugid=drug.drugid "
  sql = sql & " WHERE drug.drugStatusid='IST001' AND drugpricematrix2.insurancetypeid='" & insType & "'"
  cnt = 0
  drgQty = 1
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      response.write "<table id='myTable' border='1'>"
      response.write "<tr><th>No.</th><th>ID</th><th>DrugName</th><th>Quantity</th><th>UnitCost</th><th style='width:180px!important;'>Status</th></tr>"
      Do While Not .EOF
        cnt = cnt + 1
        drugId = .fields("drugid")
        price = .fields("ItemUnitCost")
        dsc = drugId & "||" & price
        sql2 = "SELECT * from PerformVar59 WHERE PerformVar59Name = '" & VisitID & "' AND Description = '" & dsc & "'"
        With rst2
            .open qryPro.FltQry(sql2), conn, 3, 4
            If .RecordCount > 0 Then
                .MoveFirst
                drgQtyArr = Split(.fields("KeyPrefix"), "||")
                drgQty1 = drgQtyArr(1)
                ' drgQty1 = chkQty
                response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & drugId & "</td><td>" & GetComboName("Drug", drugId) & "</td><td><input type=""number"" id='drg-" & drugId & "' class=""Quantity"" name=""Quantity"" value=" & drgQty1 & " min=""1""></td><td>" & (FormatNumber(CStr(price), 2, , , -1)) & "</td><td>" & BuildDrugCheckBtn(drugId, VisitID, price, desc) & "</td></tr>"
            Else
                response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & drugId & "</td><td>" & GetComboName("Drug", drugId) & "</td><td><input type=""number"" id='drg-" & drugId & "' class=""Quantity"" name=""Quantity"" value=" & drgQty & " min=""1""></td><td>" & (FormatNumber(CStr(price), 2, , , -1)) & "</td><td>" & BuildDrugCheckBtn(drugId, VisitID, price, desc) & "</td></tr>"
            End If
            .Close
        End With
      .MoveNext
      Loop
      response.write "</table>"
    End If
  .Close
  End With
  Set rst = Nothing
  Set rst2 = Nothing
End Sub

Sub ListItems(VisitID)
  Dim rst, ot, sql, itemID, cnt, price, itmQtyArr, itmQty, itmQty1, rst2, dsc, sql2
  Set rst = server.CreateObject("ADODB.Recordset")
  Set rst2 = server.CreateObject("ADODB.Recordset")
  sql = "SELECT TOP 100 items.itemid, itempricematrix2.itemunitcost FROM items join itempricematrix2 on itempricematrix2.itemid=items.itemid "
  sql = sql & " WHERE items.itemStatusid='IST001' and itempricematrix2.insurancetypeid='" & insType & "'"
  cnt = 0
  itmQty = 1
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      response.write "<table id='myTable' border='1'>"
      response.write "<tr><th>No.</th><th>ID</th><th>Name</th><th>Quantity</th><th>UnitCost</th><th style='width:180px!important;'>Status</th></tr>"
      Do While Not .EOF
        cnt = cnt + 1
        itemID = .fields("itemid")
        price = .fields("itemunitcost")
        dsc = itemID & "||" & price
        sql2 = "SELECT * from PerformVar59 WHERE PerformVar59Name = '" & VisitID & "' AND Description = '" & dsc & "'"
        With rst2
          rst2.open qryPro.FltQry(sql2), conn, 3, 4
          If rst2.RecordCount > 0 Then
            .MoveFirst
            itmQtyArr = Split(.fields("KeyPrefix"), "||")
            itmQty1 = itmQtyArr(1)
            response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & itemID & "</td><td>" & GetComboName("Items", itemID) & "</td><td><input type=""number"" id='itm-" & itemID & "' class=""Quantity"" name=""Quantity"" value=" & itmQty1 & " min=""1""></td><td>" & (FormatNumber(CStr(price), 2, , , -1)) & "</td><td>" & BuildItemCheckBtn(itemID, VisitID, price) & "</td></tr>"
          Else
            response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & itemID & "</td><td>" & GetComboName("Items", itemID) & "</td><td><input type=""number"" id='itm-" & itemID & "' class=""Quantity"" name=""Quantity"" value=" & itmQty & " min=""1""></td><td>" & (FormatNumber(CStr(price), 2, , , -1)) & "</td><td>" & BuildItemCheckBtn(itemID, VisitID, price) & "</td></tr>"
          End If
        rst2.Close
        End With
      .MoveNext
      Loop
      response.write "</table>"
    End If
  .Close
  End With
  Set rst = Nothing
  Set rst2 = Nothing
End Sub

Sub ListLabs(VisitID)
  Dim rst, rst2, ot, sql, testID, cnt, price, labQtyArr, labQty, labQty1
  Set rst = server.CreateObject("ADODB.Recordset")
  Set rst2 = server.CreateObject("ADODB.Recordset")
  sql = "SELECT TOP 100 LabTest.labtestid, labtestcostmatrix.unitcost FROM LabTest join labtestcostmatrix on labtestcostmatrix.labtestid=labtest.labtestid "
  sql = sql & " WHERE LabTest.TestStatusid='TST001' and labtestcostmatrix.insurancetypeid='" & insType & "' and labtestcostmatrix.agegroupid='A002'"
  cnt = 0
  labQty = 1
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      response.write "<table id='myTable' border='1'>"
      response.write "<tr><th>No.</th><th>ID</th><th>Name</th><th>Quantity</th><th>UnitCost</th><th style='width:180px!important;'>Status</th></tr>"
      Do While Not .EOF
        cnt = cnt + 1
        testID = .fields("labtestid")
        price = .fields("unitcost")
        dsc = testID & "||" & price
        sql2 = "SELECT * from PerformVar59 WHERE PerformVar59Name = '" & VisitID & "' AND Description = '" & dsc & "'"
        With rst2
          .open qryPro.FltQry(sql2), conn, 3, 4
          If .RecordCount > 0 Then
            .MoveFirst
            labQtyArr = Split(.fields("KeyPrefix"), "||")
            labQty1 = labQtyArr(1)
            response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & testID & "</td><td>" & GetComboName("LabTest", testID) & "</td><td><input type=""number"" id='lab-" & testID & "' class=""Quantity"" name=""Quantity"" value=" & labQty1 & " min=""1""></td><td>" & (FormatNumber(CStr(price), 2, , , -1)) & "</td><td>" & BuildLabCheckBtn(testID, VisitID, price) & "</td></tr>"
          Else
            response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & testID & "</td><td>" & GetComboName("LabTest", testID) & "</td><td><input type=""number"" id='lab-" & testID & "' class=""Quantity"" name=""Quantity"" value=" & labQty & " min=""1""></td><td>" & (FormatNumber(CStr(price), 2, , , -1)) & "</td><td>" & BuildLabCheckBtn(testID, VisitID, price) & "</td></tr>"
          End If
        .Close
        End With
      .MoveNext
      Loop
      response.write "</table>"
    End If
  .Close
  End With
  Set rst = Nothing
  Set rst2 = Nothing
End Sub

Sub ListTreats(VisitID)
  Dim rst, ot, sql, trtID, cnt, price, trtQtyArr, trtQty, trtQty1, rst2, dsc, sql2
  Set rst = server.CreateObject("ADODB.Recordset")
  Set rst2 = server.CreateObject("ADODB.Recordset")
  insGr = GetComboNameFld("Visitation", VisitID, "InsuranceGroupID")
  sql = "SELECT TOP 100 Treatment.treatmentid, TreatCostMatrix.UnitCost FROM Treatment JOIN TreatCostMatrix on TreatCostMatrix.treatmentid=Treatment.treatmentid "
  sql = sql & " WHERE TreatInfo1='NO' and TreatCostMatrix.insuranceTypeid='" & insType & "' and TreatCostMatrix.agegroupid='A002' AND TreatCostMatrix.medicalserviceid='M003'"
  cnt = 0
  trtQty = 1
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      response.write "<table id='myTable' border='1'>"
      response.write "<tr><th>No.</th><th>ID</th><th>Name</th><th>Quantity</th><th>UnitCost</th><th style='width:180px!important;'>Status</th></tr>"
      Do While Not .EOF
        cnt = cnt + 1
        trtID = .fields("treatmentid")
        price = .fields("UnitCost")
        dsc = trtID & "||" & price
        sql2 = "SELECT * from PerformVar59 WHERE PerformVar59Name = '" & VisitID & "' AND Description = '" & dsc & "'"
        With rst2
          rst2.open qryPro.FltQry(sql2), conn, 3, 4
          If rst2.RecordCount > 0 Then
            .MoveFirst
            trtQtyArr = Split(.fields("KeyPrefix"), "||")
            trtQty1 = trtQtyArr(1)
            response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & trtID & "</td><td>" & GetComboName("Treatment", trtID) & "</td><td><input type=""number"" id='trt-" & trtID & "' class=""Quantity"" name=""Quantity"" value=" & trtQty1 & " min=""1""></td><td>" & (FormatNumber(CStr(price), 2, , , -1)) & "</td><td>" & BuildTrtCheckBtn(trtID, VisitID, price) & "</td></tr>"
          Else
            response.write vbCrLf & "<tr><td>" & cnt & "</td><td>" & trtID & "</td><td>" & GetComboName("Treatment", trtID) & "</td><td><input type=""number"" id='trt-" & trtID & "' class=""Quantity"" name=""Quantity"" value=" & trtQty & " min=""1""></td><td>" & (FormatNumber(CStr(price), 2, , , -1)) & "</td><td>" & BuildTrtCheckBtn(trtID, VisitID, price) & "</td></tr>"
          End If
        rst2.Close
        End With
      .MoveNext
      Loop
      response.write "</table>"
    End If
  .Close
  End With
  Set rst = Nothing
End Sub

Sub UpdateVisitationFld(proformaStat)

  Dim sql, valueArr, valueStatus, visitationID
  valueArr = Split(proformaStat, "||")
  If UBound(valueArr) > 0 Then
    valueStatus = valueArr(0)
    visitationID = valueArr(1)

    If UCase(valueStatus) = "P000" Then
      sql = "UPDATE Visitation SET VisitInfo5 = ''"
      sql = sql & " WHERE VisitationID = '" & visitationID & "'"
      conn.execute qryPro.FltQry(sql)
      response.write "Visitation Table Updated with " & valueStatus
    Else
      sql = "UPDATE Visitation SET VisitInfo5 = '" & valueStatus & "'"
      sql = sql & " WHERE VisitationID = '" & visitationID & "'"
      conn.execute qryPro.FltQry(sql)
      response.write "Visitation Table Updated with " & valueStatus
    End If
  End If

End Sub

Function BuildItemCheckBtn(itemID, VisitID, price)
  Dim ot, sql, rst
  Set rst = server.CreateObject("ADODB.Recordset")
  btn = itemID & "||" & price
  sql = "Select * from PerformVAr59 where "
  sql = sql & " PerformVar59Name='" & VisitID & "' AND Description = '" & btn & "'"
  ot = ""
  With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    ot = "<label><input type='checkbox' name='itm-" & itemID & "' value='itm-" & itemID & "' onclick=""DoAjaxUpdate('itm-" & itemID & "||" & price & "')"" checked> <span>Remove</span></label>"
  Else
    ot = "<label><input type='checkbox' name='itm-" & itemID & "' value='itm-" & itemID & "' onclick=""DoAjaxUpdate('itm-" & itemID & "||" & price & "')""> <span>Add</span></label>"
  End If
  .Close
  End With
  BuildItemCheckBtn = ot
  Set rst = Nothing
End Function

Function BuildLabCheckBtn(testID, VisitID, price)
  Dim ot, sql, rst, btn
  Set rst = server.CreateObject("ADODB.Recordset")
  btn = testID & "||" & price
  sql = "Select * from PerformVAr59 where "
  sql = sql & " PerformVar59Name='" & VisitID & "' AND Description = '" & btn & "'"
  ot = ""
  With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    ot = "<label><input type='checkbox' name='lab-" & testID & "' value='lab-" & testID & "' onclick=""DoAjaxUpdate('lab-" & testID & "||" & price & "')"" checked> <span>Remove</span></label>"
  Else
    ot = "<label><input type='checkbox' name='lab-" & testID & "' value='lab-" & testID & "' onclick=""DoAjaxUpdate('lab-" & testID & "||" & price & "')""> <span>Add</span></label>"
  End If
  .Close
  End With
  BuildLabCheckBtn = ot
  Set rst = Nothing
End Function

Function BuildDrugCheckBtn(drugId, VisitID, price, desc)
  Dim ot, sql, rst, btn, drgPriceTot, tot
  Set rst = server.CreateObject("ADODB.Recordset")
'  If IsNumeric(chkQty) Then
'    drgPriceTot = CInt(chkQty) * CInt(price)
'  End If
  btn = drugId & "||" & price
  sql = "Select * from PerformVAr59 where "
  sql = sql & " PerformVar59Name='" & VisitID & "' AND Description = '" & btn & "'"
  ot = ""
  With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    ot = "<label><input type='checkbox' name='drg-" & drugId & "' value='drg-" & drugId & "' onclick=""DoAjaxUpdate('drg-" & drugId & "||" & price & "')"" checked> <span>Remove</span></label>"
  Else
    ot = "<label><input type='checkbox' name='drg-" & drugId & "' value='drg-" & drugId & "' onclick=""DoAjaxUpdate('drg-" & drugId & "||" & price & "')""> <span>Add</span></label>"
  End If
  .Close
  End With
  BuildDrugCheckBtn = ot
  Set rst = Nothing
End Function

Function BuildTrtCheckBtn(trtID, VisitID, price)
  Dim ot, sql, rst, btn
  Set rst = server.CreateObject("ADODB.Recordset")
  btn = trtID & "||" & price
  sql = "Select * from PerformVAr59 where "
  sql = sql & " PerformVar59Name='" & VisitID & "' AND Description = '" & btn & "'"
  ot = ""
  With rst
  .open qryPro.FltQry(sql), conn, 3, 4
  If .RecordCount > 0 Then
    ot = "<label><input type='checkbox' name='trt-" & trtID & "' value='trt-" & trtID & "' onclick=""DoAjaxUpdate('trt-" & trtID & "||" & price & "')"" checked> <span>Remove</span></label>"
  Else
    ot = "<label><input type='checkbox' name='trt-" & trtID & "' value='trt-" & trtID & "' onclick=""DoAjaxUpdate('trt-" & trtID & "||" & price & "')""> <span>Add</span></label>"
  End If
  .Close
  End With
  BuildTrtCheckBtn = ot
  Set rst = Nothing
End Function

Function GenerateUniqueID()
  Dim randomNumber
  Randomize ' Initialize random number generator
  randomNumber = Int((999999 - 100000 + 1) * Rnd + 100000)
  Dim uniqueID
  uniqueID = chked & CStr(randomNumber)
  GenerateUniqueID = uniqueID
End Function

