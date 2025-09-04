Dim operation, secret_key
secret_key = "sk_live_89dddf9020b72b2fcb9424c3cbc7f1da2830d99e"
operation = Trim(Request("operation"))

response.Clear
response.contentType = "application/json"
If UCase(operation) = UCase("initialize-payment") Then
    response.write JSONStringify(InitializePayment())
ElseIf UCase(operation) = UCase("complete-transaction") Then
    response.write JSONStringify(CompleteTransaction())
ElseIf UCase(operation) = UCase("GetItems") Then
    response.write JSONStringify(getItems())
ElseIf UCase(operation) = UCase("GetReceipts") Then
    response.write JSONStringify(getReceipts())
ElseIf UCase(operation) = UCase("GetMenu") Then
    response.write JSONStringify(getMenu())
End If

Function getMenu()
    Dim rst, responseJSON, sql
    Set rst = CreateObject("ADODB.RecordSet")
    Set responseJSON = CreateObject("Scripting.Dictionary")
    sql = ""
    sql = sql & " SELECT FoodRecipeName  "
    sql = sql & " FROM FoodMenu "
    sql = sql & " INNER JOIN FoodRecipe ON FoodRecipe.FoodRecipeID = FoodMenu.FoodRecipeID "
    sql = sql & " WHERE appointdayid = '" & GetWorkingDay(Date) & "' "
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount = 1 Then
        responseJSON.add "success", True
        responseJSON.add "data", rst.fields("FoodRecipeName").value
    Else
        responseJSON.add "success", False
    End If
    rst.Close
    Set rst = Nothing
    Set getMenu = responseJSON
End Function
Function GetWorkingDay(d)
    GetWorkingDay = "DAY" & year(d) & Right("00" & Month(d), 2) & Right("00" & Day(d), 2)
End Function
Function getItems()
  Dim JSONArray, jsonObj, rst, sn, sql, billgroup, billGroupID
  Set JSONArray = CreateObject("System.Collections.ArrayList")
  Set rst = CreateObject("ADODB.RecordSet")
  billgroup = Trim(Request("store"))
  sn = Trim(Request("sn"))

  If UCase(billgroup) = "STAFF" Then
    billGroupID = "LSHHI172"
  Else
    billGroupID = "LSHHI170"
  End If
  sql = ""
  sql = sql & "SELECT * FROM Treatment WHERE billgroupid = '" & billGroupID & "' And TreatInfo1 = 'NO'"
  If sn <> "" Then
      sql = sql & " AND (treatmentid = '" & sn & "' "
      For Each word In Split(sn, " ")
          If word <> "" Then
              sql = sql & " OR treatmentname like '%" & word & "%' "
          End If
      Next
      sql = sql & " ) "
  End If
  rst.open qryPro.FltQry(sql), conn, 3, 4
  If rst.RecordCount > 0 Then
      rst.MoveFirst
      Do While Not rst.EOF
          Set jsonObj = CreateObject("Scripting.Dictionary")
          jsonObj.add "itemid", rst.fields("TreatmentId").value
          jsonObj.add "itemname", rst.fields("TreatmentName").value
          jsonObj.add "itemprice", Round(CDbl(rst.fields("UnitCost")), 2)
          JSONArray.add jsonObj
          rst.MoveNext
      Loop
  End If
  Set rst = Nothing
  Set getItems = JSONArray
End Function

Function getReceipts()
    Dim rst, sql, JSONArray, jsonObj, searchParameter, filterDate
    searchParameter = Trim(Request("search"))
    filterDate = Trim(Request("date"))
    Set rst = CreateObject("ADODB.RecordSet")
    Set JSONArray = CreateObject("System.Collections.ArrayList")

    If filterDate = "" Then
      filterDate = year(Date) & "-" & Right("00" & Month(Date), 2) & "-" & Right("00" & Day(Date), 2)
    End If

    sql = ""
    sql = sql & " SELECT ConsultReview.ConsultReviewID,ConsultReview.ConsultReviewName "
    sql = sql & "     ,ConsultReview.ConsultReviewDate, COUNT(Treatment.TreatmentID) as count "
    sql = sql & "     ,SUM(TreatCharges.Qty) AS qty, SUM(TreatCharges.FinalAmt) AS finalamt "
    sql = sql & " FROM ConsultReview "
    sql = sql & " INNER JOIN TreatCharges ON TreatCharges.ConsultReviewID = ConsultReview.ConsultReviewID "
    sql = sql & " INNER JOIN Treatment ON Treatment.TreatmentID = TreatCharges.TreatmentID "
    sql = sql & " WHERE Treatment.BillGroupID in ('LSHHI170','LSHHI172')  "
    If searchParameter <> "" Then
        sql = sql & " AND (ConsultReview.ConsultReviewID = '" & searchParameter & "' "
        For Each word In Split(searchParameter, " ")
            sql = sql & " OR ConsultReview.ConsultReviewName LIKE '%" & word & "%' "
        Next
        sql = sql & " ) "
    End If
    If UCase(jschd) <> "M26" Then
        sql = sql & "AND ConsultReview.systemuserid = '" & uName & "'"
    End If
    If filterDate <> "" Then
      sql = sql & " AND CAST(ConsultReview.ConsultReviewDate AS DATE) = '" & filterDate & "' "
    End If
    sql = sql & " GROUP BY ConsultReview.ConsultReviewID,ConsultReview.ConsultReviewName,ConsultReview.ConsultReviewDate "
    sql = sql & " ORDER BY ConsultReview.ConsultReviewDate desc "
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            Set jsonObj = CreateObject("Scripting.Dictionary")
            jsonObj.add "drugsaleid", rst.fields("ConsultReviewID").value
            jsonObj.add "drugsalename", rst.fields("ConsultReviewName").value
            jsonObj.add "itemno", rst.fields("Count").value
            jsonObj.add "itemqty", rst.fields("qty").value
            jsonObj.add "itemprice", Round(rst.fields("finalamt"), 2)
            JSONArray.add jsonObj
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    Set getReceipts = JSONArray
End Function

Function CompleteTransaction()
    Dim reference, rst, consultReviewID, data
    Set rst = CreateObject("ADODB.RecordSet")
    Set CompleteTransaction = CreateObject("Scripting.Dictionary")
    reference = Trim(Request("reference"))
    sql = "SELECT * FROM PerformVar29 WHERE PerformVar29ID = '" & reference & "'"
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount = 1 Then
        url = "https://api.paystack.co/transaction/verify/" & reference

        Set headerJSON = CreateObject("Scripting.Dictionary")

        headerJSON.add "Authorization", "Bearer " & secret_key
        Set responseJSON = FetchHttp(url, "GET", headerJSON, "JSON", "")
        If Not responseJSON("status") Then
            CompleteTransaction.add "success", False
            CompleteTransaction.add "message", "Something went wrong, try again!"
            Set responseJSON = Nothing
            Exit Function
        End If

        Set responseData = responseJSON("data")
        If UCase(responseData("status")) <> "SUCCESS" Then
            CompleteTransaction.add "success", False
            CompleteTransaction.add "message", "Transaction was not successful"
            Set responseJSON = Nothing
            Exit Function
        End If

        If responseData("amount") < rst.fields("amount") Then
            CompleteTransaction.add "success", False
            CompleteTransaction.add "message", "Amount was not paid in full"
            Set responseJSON = Nothing
            Exit Function
        End If

        rst.fields("KeyPrefix") = "COMPLETED"
        rst.fields("amount") = responseData("amount") / 100

        Set data = JSONParse(rst.fields("Description"), False)

        rst.Updatebatch
        
        consultReviewID = AddConsultReview(reference)
        
        For Each obj In data
            AddTreatCharges consultReviewID, obj("itemid"), obj("issueqty")
        Next
    End If
    rst.Close

    CompleteTransaction.add "success", True
    CompleteTransaction.add "message", "Order processed!"
    CompleteTransaction.add "id", consultReviewID
    Set responseJSON = Nothing
    Set data = Nothing
    Set rst = Nothing
End Function

Function InitializePayment()
    Dim url, headerJSON, responseJSON, RequestJSON, isValid, jsonObj, amount, RequestData, uMail
    amount = 0
    url = "https://api.paystack.co/transaction/initialize"
    Set InitializePayment = CreateObject("Scripting.Dictionary")
    Set RequestJSON = GetRequestBodyAsJSON(False)

    For Each obj In RequestJSON("orderitems")
        isValid = obj("issueqty") > 0
        If Not isValid Then
            InitializePayment.add "success", False
            InitializePayment.add "message", "Quantity for " & obj("itemname") & " [" & obj("itemid") & "] cannot be less than 1"
            Set RequestJSON = Nothing
            Exit Function
        End If
        amount = amount + obj("issueqty") * GetComboNameFld("Treatment", obj("itemid"), "UnitCost")
    Next
        
    Set headerJSON = CreateObject("Scripting.Dictionary")
    Set data = CreateObject("Scripting.Dictionary")

    headerJSON.add "Authorization", "Bearer " & secret_key
    headerJSON.add "Content-Type", "application/json"

    uMail = uName & "_" & amount & "@gmail.com"
    data.add "email", uMail '"customer@gmail.com"
    data.add "amount", amount * 100
    data.add "currency", "GHS"
    data.add "reference", GetRandomUUID(32)
    
    Set responseJSON = FetchHttp(url, "POST", headerJSON, "JSON", data)
  
    If Not responseJSON("status") Then
        InitializePayment.add "success", False
        InitializePayment.add "message", "Something went wrong, try again!"
        Set RequestJSON = Nothing
        Exit Function
    End If
    
    Set RequestData = responseJSON("data")

    Set jsonObj = CreateObject("Scripting.Dictionary")
    jsonObj.add "access_code", RequestData("access_code")
    jsonObj.add "reference", RequestData("reference")

    AddTransaction RequestData("reference"), amount, RequestJSON("orderitems")

    InitializePayment.add "success", True
    InitializePayment.add "data", jsonObj

    Set jsonObj = Nothing
    Set responseJSON = Nothing
    Set RequestData = Nothing
    Set data = Nothing
    Set headerJSON = Nothing
End Function

Function AddConsultReview(reference)
    Dim rst, recKy, sql
    Set rst = CreateObject("ADODB.RecordSet")
    recKy = GetRecordKey("ConsultReview", "ConsultReviewID", "NONE")
    sql = "SELECT * FROM ConsultReview WHERE ConsultReviewID = '" & recKy & "'"
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("ConsultReviewID") = recKy
        rst.fields("ConsultReviewName") = GetComboName("Staff", GetComboNameFld("SystemUser", uName, "StaffID"))
        rst.fields("Comments") = ""
        rst.fields("VisitationID") = "E01"
        rst.fields("MedicalStaffID") = "M0304"
        rst.fields("MedicalServiceID") = "M001"
        rst.fields("AgeGroupID") = "A002"
        rst.fields("PatientID") = "P1"
        rst.fields("GenderID") = "G001"
        rst.fields("InsuredPatientID") = "E01"
        rst.fields("InsuranceTypeID") = "I100"
        rst.fields("InsuranceSchemeID") = "SELF-000032"
        rst.fields("InsuranceNo") = "-"
        rst.fields("ReceiptTypeID") = "PAYSTACK"
        rst.fields("JobScheduleID") = jschd
        rst.fields("SystemUserID") = uName
        rst.fields("BranchID") = brnch
        rst.fields("WorkingYearID") = "YRS" & year(Date)
        rst.fields("WorkingMonthID") = "MTH" & year(Date) & Right("00" & Month(Date), 2)
        rst.fields("WorkingDayID") = "DAY" & year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2)
        rst.fields("ConsultReviewDate") = Now
        rst.fields("KeyPrefix") = ""
        rst.fields("InsuranceZoneID") = "I100-CASH"
        rst.fields("TransProcessStatID") = "T001"
        rst.fields("TransProcessValID") = "ConsultReviewPro-T001"
        rst.fields("RevenueCenterID") = "R111400"
        rst.fields("BillYearID") = "NONE"
        rst.fields("BillMonthID") = "NONE"
        rst.fields("BillDayID") = "NONE"
        rst.fields("BillProcessDate") = Now
        rst.fields("SponsorID") = "SELF"
        rst.fields("BranchBatchID") = brnch & "-" & "DAY" & year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2)
        rst.fields("BranchSubBatchID") = brnch & "-" & "DAY" & year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "-ConsultReview"
        rst.fields("ServiceNo") = "-"
        rst.fields("MainValue1") = 0
        rst.fields("MainValue2") = 0
        rst.fields("MainValue3") = 0
        rst.fields("MainValue4") = 0
        rst.fields("MainDate1") = Now
        rst.fields("MainDate2") = Now
        rst.fields("MainDate3") = Now
        rst.fields("MainDate4") = Now
        rst.fields("MainInfo1") = reference
        rst.fields("MainInfo2") = ""
        rst.fields("MainInfo3") = "ConsultReview"
        rst.fields("MainInfo4") = ""
        rst.fields("MainInfo5") = ""
        rst.fields("MainInfo6") = ""
        rst.Updatebatch
    End If
    rst.Close
    Set rst = Nothing
    AddConsultReview = recKy
End Function

Sub AddTreatCharges(consultReviewID, treatmentid, qty)
    Dim rst, UnitCost
    Set rst = CreateObject("ADODB.RecordSet")
    UnitCost = GetComboNameFld("Treatment", treatmentid, "UnitCost")
    sql = "SELECT * FROM TreatCharges WHERE ConsultReviewID = '" & consultReviewID & "' AND TreatmentID = '" & treatmentid & "'"
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("ConsultReviewID") = consultReviewID
        rst.fields("TreatmentID") = treatmentid
        rst.fields("TreatCategoryID") = "LSHHI16"
        rst.fields("TreatTypeID") = "LSHHI170"
        rst.fields("PatientID") = "P1"
        rst.fields("VisitationID") = "E01"
        rst.fields("GenderID") = "G001"
        rst.fields("AgeGroupID") = "A001"
        rst.fields("InsuredPatientID") = "E01"
        rst.fields("InsuranceTypeID") = "I100"
        rst.fields("InsuranceSchemeID") = "SELF-000032"
        rst.fields("InsuranceNo") = "-"
        rst.fields("ReceiptTypeID") = "R001"
        rst.fields("MedicalServiceID") = "M001"
        rst.fields("ConsultReviewDate") = Now
        rst.fields("JobScheduleID") = jschd
        rst.fields("SystemUserID") = uName
        rst.fields("WorkingYearID") = "YRS" & year(Date)
        rst.fields("WorkingMonthID") = "MTH" & year(Date) & Right("00" & Month(Date), 2)
        rst.fields("WorkingDayID") = "DAY" & year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2)
        rst.fields("Qty") = qty
        rst.fields("UnitCost") = UnitCost
        rst.fields("InitAmt") = qty * UnitCost
        rst.fields("DiscAmt") = 0
        rst.fields("FinalAmt") = qty * UnitCost
        rst.fields("InsuranceZoneID") = "I100-CASH"
        rst.fields("TransProcessStatID") = "T001"
        rst.fields("TransProcessValID") = "ConsultReviewPro-T001"
        rst.fields("BillYearID") = "NONE"
        rst.fields("BillMonthID") = "NONE"
        rst.fields("BillDayID") = "NONE"
        rst.fields("BillProcessDate") = Now
        rst.fields("MedicalStaffID") = "001"
        rst.fields("TreatGroupID") = "LSHHI16"
        rst.fields("TreatClassID") = "T001"
        rst.fields("TreatModeID") = "LSHHI16"
        rst.fields("RevenueCenterID") = "R111400"
        rst.fields("BillGroupCatID") = "LSHHI16"
        rst.fields("BillGroupID") = "LSHHI170"
        rst.fields("SponsorID") = "SELF"
        rst.fields("BranchBatchID") = brnch & "-" & "DAY" & year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2)
        rst.fields("BranchSubBatchID") = brnch & "-" & "DAY" & year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "-ConsultReview"
        rst.fields("BranchID") = brnch
        rst.fields("MainInfo1") = ""
        rst.fields("MainInfo2") = ""
        rst.fields("MainValue1") = 0
        rst.fields("MainValue2") = 0
        rst.fields("MainDate1") = Now
        rst.fields("MainDate2") = Now

        rst.Updatebatch
    End If
    rst.Close
    Set rst = Nothing
End Sub

Sub AddTransaction(reference, amount, data)
    Dim rst
    Set rst = CreateObject("ADODB.RecordSet")

    sql = "SELECT * FROM PerformVar29 WHERE PerformVar29ID = '" & reference & "'"
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("PerformVar29ID") = reference
        rst.fields("PerformVar29Name") = GetComboName("Staff", GetComboNameFld("Systemuser", uName, "StaffID"))
        rst.fields("KeyPrefix") = "INITIATED"
        rst.fields("Description") = JSONStringify(data)
        rst.fields("Amount") = amount
        rst.fields("SystemUserID") = uName
        rst.fields("JobscheduleID") = jschd
        rst.fields("WorkingdayID") = "DAY" & year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2)
        rst.fields("WorkingMonthID") = "MTH" & year(Date) & Right("00" & Month(Date), 2)
        rst.fields("WorkingYearID") = "YRS" & year(Date)
        rst.fields("EntryDate") = Now
        rst.Updatebatch
    End If
    rst.Close
    Set rst = Nothing
End Sub

Function FetchHttp(url, method, headerJSON, responseType, data)
    Dim http

    Set http = CreateObject("MSXML2.XMLHTTP")

    http.open UCase(method), url, False

    For Each key In headerJSON.Keys
        http.setRequestHeader key, headerJSON(key)
    Next

    If UCase(method) = "POST" Or UCase(method) = "PUT" Or UCase(method) = "PATCH" Then
        If Not IsNull(data) Then
            http.send (JSONStringify(data))
        Else
            http.send
        End If
    Else
        http.send
    End If

    If UCase(responseType) = "JSON" Then
        Set FetchHttp = JSONParse(http.responseText, False)
    Else
        FetchHttp = http.responseText
    End If
    Set http = Nothing
End Function

Function GetRandomUUID(size)
    If size < 10 Then
        GetRandomUUID = ""
    End If
    Randomize
    str = ""
    For i = 0 To size
        num = Int(Rnd * 10000)
        If num <= 5000 Then
            num = (num Mod 10) + 48
        Else
            num = (num Mod 26) + 97
        End If
        If num Mod 3 = 0 Then
            str = str & UCase(Chr(num))
        ElseIf num Mod 32 = 1 Then
            str = str & "-"
        Else
            str = str & Chr(num)
        End If
    Next
    GetRandomUUID = str
End Function

Function GetRequestBodyAsJSON(caseSensitivity)
   Dim ot, Stream, bytes, str

   Set ot = Nothing
   If IsObject(Request) Then
      bytes = Request.BinaryRead(Request.TotalBytes)
      If Len(bytes) > 0 Then
            Set Stream = server.CreateObject("Adodb.Stream")
            With Stream
               .Type = 1 'adTypeBinary
               .open
               .write bytes
               .position = 0
               .Type = 2 'adTypeText
               .Charset = "utf-8"

               str = Stream.readtext
                    
               .Close
            End With
                    
            Set Stream = Nothing
            Set ot = JSONParse(str, caseSensitivity)
      End If
   Else
      Set ot = CreatObject("Scripting.Dictionary")
      ot.CompareMode = 1
   End If

   Set GetRequestBodyAsJSON = ot
End Function

Function JSONStringify(obj)
   Dim refList, sc
   Set refList = CreateObject("System.Collections.ArrayList")
        
   JSONStringify = JSONStringify_(obj, refList)
   Set refList = Nothing
End Function
Private Function JSONStringify_(ByRef obj, ByRef refList)
   Dim key, tmpKey, value, tmp, ot, field
   Dim objType: objType = typename(obj)
   tmp = ""


   If IsObject(obj) Then

      If refList.Contains(obj) Then
             
      Else
         refList.add obj

      End If
   End If
        
   If objType = "Dictionary" Then
      For Each key In obj.Keys()
            If tmp <> "" Then tmp = tmp & ", "
            tmpKey = """" & key & """"
            tmp = tmp & tmpKey & ":" & JSONStringify_(obj(key), refList)
      Next
      ot = "{" & tmp & "}"
   ElseIf objType = "Fields" Then 'ADODB.Fields
      For Each field In obj
            If tmp <> "" Then tmp = tmp & ", "
            tmpKey = """" & field.name & """"
            tmp = tmp & tmpKey & ":" & JSONStringify_(field.value, refList)
      Next
      ot = "{" & tmp & "}"
   ElseIf IsArray(obj) Or objType = "ArrayList" Then
      For Each value In obj
            If tmp <> "" Then tmp = tmp & ", "
            tmp = tmp & JSONStringify_(value, refList)
      Next
      ot = "[" & tmp & "]"
   ElseIf objType = "String" Then
      tmp = Replace(obj, "\", "\\")
      tmp = Replace(tmp, """", "\""")
      tmp = Replace(tmp, vbTab, "\t")
      tmp = Replace(tmp, vbCrLf, "\r\n")
      tmp = Replace(tmp, vbCr, "\r")
      tmp = Replace(tmp, vbLf, "\n")
      ot = """" & tmp & """"
   ElseIf objType = "Boolean" Then
      ot = "" & LCase(obj) & ""
   ElseIf objType = "Byte" Then
      ot = CDbl(obj) 'Compatible with JSON.parse
   ElseIf objType = "Integer" Or objType = "Double" Or objType = "Long" Or objType = "Single" Or objType = "Currency" Then
      ot = obj
   ElseIf objType = "Empty" Or objType = "Null" Then
      ot = "null"
   ElseIf objType = "Date" Then
      ot = """" & obj & """"
   Else
      ot = """[Object " & objType & "]"""
   End If

   JSONStringify_ = ot
End Function
Function JSONParse(jStr, caseSensitivity)
   'Bluefaces
   Dim scriptEngine, arr, item, hasErr

   Set scriptEngine = CreateObject("MSScriptControl.ScriptControl")
   scriptEngine.Language = "JavaScript"
   scriptEngine.AddCode "function isObject(jsonObj){return (typeof jsonObj === 'object' || typeof jsonObj === 'function') && (jsonObj !== null);}"
   scriptEngine.AddCode "function isArray(jsonObj){return Object.prototype.toString.call(jsonObj) === '[object Array]';}"
        
   On Error Resume Next
      scriptEngine.AddCode "var jsonObject = " & jStr
      hasErr = Err.number <> 0
   On Error GoTo 0
   If hasErr <> 0 Then
      'User defined error boundry, -2147221504
      Err.Raise (-2147221504 + 10900), "JSONParse", "Unexpected token '" & jStr & "', """ & jStr & """ is not valid JSON", 0
   Else
      If scriptEngine.Run("isObject", scriptEngine.CodeObject.JSONobject) Or scriptEngine.Run("isArray", scriptEngine.CodeObject.JSONobject) Then
            Set JSONParse = JSONParse_(scriptEngine.CodeObject.JSONobject, caseSensitivity)
      Else
            JSONParse = JSONParse_(scriptEngine.CodeObject.JSONobject, caseSensitivity)
      End If
   End If
        
End Function
Function JSONParse_(jScriptType, caseSensitivity)
   Dim scriptEngine, tmp, arr, item, key
   Set scriptEngine = CreateObject("MSScriptControl.ScriptControl")
   scriptEngine.Language = "JavaScript"

   scriptEngine.AddCode "function isArray(jsonObj){return Object.prototype.toString.call(jsonObj) === '[object Array]';}"
   scriptEngine.AddCode "function isObject(jsonObj){return (typeof jsonObj === 'object' || typeof jsonObj === 'function') && (jsonObj !== null);}"
   scriptEngine.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
   scriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
        
   If scriptEngine.Run("isArray", jScriptType) Then
      Set JSONParse_ = CreateObject("System.Collections.ArrayList")
      For Each item In jScriptType
            JSONParse_.add JSONParse_(item, caseSensitivity)
      Next
   ElseIf scriptEngine.Run("isObject", jScriptType) Then
      Set JSONParse_ = CreateObject("Scripting.Dictionary")
      If caseSensitivity = False Then
            JSONParse_.CompareMode = 1
      End If

      For Each key In scriptEngine.Run("getKeys", jScriptType)
            If JSONParse_.Exists(key) Then
               JSONParse_.Remove key
            End If
            JSONParse_.add key, JSONParse_(scriptEngine.Run("getProperty", jScriptType, key), caseSensitivity)
      Next
   Else
      JSONParse_ = jScriptType
   End If
End Function





