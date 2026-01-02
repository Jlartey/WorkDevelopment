
Dim operation

response.Clear
response.ContentType = "application/json"

operation = Trim(Request("operation"))

If UCase(operation) = "ADDRACK" Then
    response.write JSONStringify(AddRack())
ElseIf UCase(operation) = "LOADRACKS" Then
    response.write JSONStringify(LoadRacks())
ElseIf UCase(operation) = "CHANGESTATUS" Then
    response.write JSONStringify(DisableRack())
ElseIf UCase(operation) = "GETTRVRACKS" Then
    response.write JSONStringify(GetTrvRacks())
ElseIf UCase(operation) = "GETRECENTVALUE" Then
    response.write JSONStringify(MostRecentValue())
ElseIf UCase(operation) = "LOADDRUG" Then
    response.write JSONStringify(GetItemInRack())
ElseIf UCase(operation) = "ADDDRUG" Then
    response.write JSONStringify(AddItemToRack())
ElseIf UCase(operation) = "DELETEDRUG" Then
    response.write JSONStringify(DeleteItemInRack())
ElseIf UCase(operation) = "SEARCHDRUG" Then
    response.write JSONStringify(GetItemBySearch())
ElseIf UCase(operation) = "PARENTRACK" Then
    response.write JSONStringify(GetParentRack())
End If

Function GetParentRack()
    Dim JSONArray, jsonObj, rst, sql
    Set rst = CreateObject("ADODB.RecordSet")
    Set JSONArray = CreateObject("System.Collections.ArrayList")
    sql = "SELECT DISTINCT rack FROM PerformVar22 WHERE KeyPrefix='" & GetUserDrugStore(jSchd) & "' AND (rack IS NOT NULL OR rack <> '')"
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            Set jsonObj = CreateObject("Scripting.Dictionary")
            jsonObj.Add "id", rst.fields("rack").value
            jsonObj.Add "name", rst.fields("rack").value
            JSONArray.Add jsonObj
            rst.MoveNext
        Loop
    End If
    Set GetParentRack = JSONArray
End Function

Function GetItemBySearch()
    Dim tbl, sql, sn, rst, JSONArray, jsonObj
    Set rst = CreateObject("ADODB.RecordSet")
    Set JSONArray = CreateObject("System.Collections.ArrayList")
    sn = Trim(Request("sn"))
    tbl = Trim(Request("tbl"))
    If tbl = "S20" Then
        sql = " SELECT top 5 ItemId [DrugId],ItemName [DrugName] FROM Items "
        sql = sql & " WHERE ItemStatusID='IST001' AND (ItemId = '" & sn & "' OR ( "
        For Each word In Split(sn, " ")
            sql = sql & " ItemName like '%" & word & "%' AND "
        Next
        sql = sql & " 1=1 )) "
    Else
        sql = " SELECT top 5 DrugId,DrugName FROM Drug "
        sql = sql & " WHERE DrugStatusID='IST001' AND  (DrugId = '" & sn & "' OR ( "
        For Each word In Split(sn, " ")
            sql = sql & " DrugName like '%" & word & "%' AND "
        Next
        sql = sql & " 1=1 )) "
    End If
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            Set jsonObj = CreateObject("Scripting.Dictionary")
            jsonObj.Add "drugid", rst.fields("DrugID").value
            jsonObj.Add "drugname", rst.fields("DrugName").value
            JSONArray.Add jsonObj
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    Set GetItemBySearch = JSONArray
End Function
Function DeleteItemInRack()
    Dim storeID, rackList, rst, jsonObj, recky, drugId
    rackid = Trim(Request("rack-id"))
    store = Trim(Request("store"))
    drugId = Trim(Request("drugid"))
    Set jsonObj = CreateObject("Scripting.Dictionary")
    recky = drugId & "||" & rackid & "||" & store
    sql = "DELETE PerformVar24 FROM PerformVar24 WHERE PerformVar24ID = '" & recky & "'"
    conn.execute sql
    jsonObj.Add "success", True
    jsonObj.Add "message", "Deleted!"
    Set DeleteItemInRack = jsonObj
End Function
Function AddItemToRack()
    Dim storeID, rackList, rst, jsonObj, recky, drugId
    rackid = Trim(Request("rack-id"))
    store = Trim(Request("store"))
    drugId = Trim(Request("drugid"))
    Set jsonObj = CreateObject("Scripting.Dictionary")
    Set rst = CreateObject("ADODB.RecordSet")
    recky = drugId & "||" & rackid & "||" & store
    sql = "SELECT * FROM PerformVar24 WHERE PerformVar24ID = '" & recky & "'"
    rst.open sql, conn, 3, 4
    If rst.recordCount = 0 Then
        rst.AddNew
        rst.fields("PerformVar24ID") = recky
        rst.fields("PerformVar24Name") = drugId
        rst.fields("Description") = rackid
        rst.fields("KeyPrefix") = store
        rst.updatebatch
    Else
        jsonObj.Add "success", False
        jsonObj.Add "message", "Drug already exist in Rack!"
        Set AddItemToRack = jsonObj
        rst.Close
        Set rst = Nothing
        Exit Function
    End If
    jsonObj.Add "success", True
    jsonObj.Add "message", "Saved!"
    Set AddItemToRack = jsonObj
    rst.Close
    Set rst = Nothing
End Function
Function GetItemInRack()
    Dim rackid, storeID, rst, JSONArray, jsonObj, tbl
    Set JSONArray = CreateObject("System.Collections.ArrayList")
    Set rst = CreateObject("ADODB.RecordSet")
    rackid = Trim(Request("rack-id"))
    store = Trim(Request("store"))
    tbl = Trim(Request("tbl"))
    sql = ""
    If UCase(tbl) = "S20" Then
        sql = sql & " SELECT Items.ItemId [DrugID],Items.ItemName [DrugName]"
        sql = sql & " FROM PerformVar24  "
        sql = sql & " INNER JOIN Items ON Items.ItemId = PerformVar24.PerformVar24Name "
    Else
        sql = sql & " SELECT Drug.DrugID,Drug.DrugName "
        sql = sql & " FROM PerformVar24  "
        sql = sql & " INNER JOIN Drug ON Drug.DrugID = PerformVar24.PerformVar24Name "
    End If
    sql = sql & " WHERE PerformVar24.Description = '" & rackid & "' AND PerformVar24.KeyPrefix = '" & store & "' "
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            Set jsonObj = CreateObject("Scripting.Dictionary")
            jsonObj.Add "drugid", rst.fields("DrugID").value
            jsonObj.Add "drugname", rst.fields("DrugName").value
            JSONArray.Add jsonObj
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    Set GetItemInRack = JSONArray
End Function
Function MostRecentValue()
    Dim tbl, rec, rst, jsonObj, mnFld
    Set rst = CreateObject("ADODB.RecordSet")
    Set jsonObj = CreateObject("Scripting.Dictionary")
    tbl = Trim(Request("tbl"))
    rec = Trim(Request("recFld"))
    mnFld = Trim(Request("mainFld"))
    id = Trim(Request("id"))
    sql = "SELECT TOP 1 " & rec & " FROM " & tbl & " WHERE " & rec & " <> '' AND " & rec & " IS NOT NULL AND " & mnFld & " = '" & id & "' "
    sql = sql & " ORDER BY EntryDate Desc"
    rst.open sql, conn, 3, 4
    If rst.recordCount = 1 Then
        jsonObj.Add "success", True
        jsonObj.Add "value", rst.fields(rec).value
    Else
        jsonObj.Add "success", False
        jsonObj.Add "value", ""
    End If
    rst.Close
    Set rst = Nothing
    Set MostRecentValue = jsonObj
End Function


Function GetTrvRacks()
    Dim rst, store, JSONArray, jsonObj, tbl, rec, mnFld
    Set rst = CreateObject("ADODB.RecordSet")
    store = Trim(Request("store"))
    itemId = Trim(Request("itemid"))
    tbl = Trim(Request("tbl"))
    rec = Trim(Request("recFld"))
    mnFld = Trim(Request("mainFld"))

    Set JSONArray = CreateObject("System.Collections.ArrayList")
    sql = ""
    sql = sql & " SELECT PerformVar22.PerformVar22ID,PerformVar22.PerformVar22Name, "
    sql = sql & " CASE WHEN rcnt." & rec & " IS NULL Then 0 else 1 END [status] FROM PerformVar24 "
    sql = sql & " INNER JOIN PerformVar22 ON PerformVar22.PerformVar22ID = PerformVar24.Description "
    sql = sql & " LEFT JOIN (SELECT TOP 1 " & rec & " FROM " & tbl & " WHERE " & rec & " <> '' AND "
    sql = sql & rec & " IS NOT NULL AND " & mnFld & " = '" & itemId & "' ) rcnt ON rcnt." & rec & " = PerformVar22.PerformVar22ID "
    sql = sql & " WHERE PerformVar24.PerformVar24Name = '" & itemId & "' AND PerformVar24.KeyPrefix = '" & store & "' "
    sql = sql & " and PerformVar22.Description = 'active'"
    rst.open sql, conn, 3, 4
    If rst.recordCount = 0 Then
        rst.Close
        sql = "SELECT PerformVar22.PerformVar22ID,PerformVar22.PerformVar22Name, "
        sql = sql & " CASE WHEN rcnt." & rec & " IS NULL Then 0 else 1 END [status] FROM PerformVar22 "
        sql = sql & " LEFT JOIN (SELECT TOP 1 " & rec & " FROM " & tbl & " WHERE " & rec & " <> '' AND "
        sql = sql & rec & " IS NOT NULL AND " & mnFld & " = '" & itemId & "' ) rcnt ON rcnt." & rec & " = PerformVar22.PerformVar22ID "
        sql = sql & "WHERE KeyPrefix = '" & store & "' and Description = 'active' "
        rst.open sql, conn, 3, 4
    End If
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            Set jsonObj = CreateObject("Scripting.Dictionary")
            jsonObj.Add "rackid", rst.fields("PerformVar22ID").value
            jsonObj.Add "rackname", rst.fields("PerformVar22Name").value
            jsonObj.Add "selected", rst.fields("status") = 1
            JSONArray.Add jsonObj
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    Set GetTrvRacks = JSONArray
End Function

Function AddRack()
    Dim RequestJSON, rst, rackname, rackid, store, jsonObj, parent
    Set jsonObj = CreateObject("Scripting.Dictionary")
    Set RequestJSON = GetRequestBodyAsJSON(False)
    Set rst = CreateObject("ADODB.RecordSet")
    If RequestJSON.Exists("rackname") Then
        rackname = RequestJSON("rackname")
    End If
    If RequestJSON.Exists("rackid") Then
        rackid = RequestJSON("rackid")
    End If
    If RequestJSON.Exists("store") Then
        store = RequestJSON("store")
    End If
    If Not RequestJSON.Exists("parent") Then
        jsonObj.Add "success", False
        jsonObj.Add "message", "Rack not entered!"
        Set AddRack = jsonObj
        Exit Function
    End If
    parent = RequestJSON("parent")
    parent = UCase(Trim(parent))
    If RackNameExist(rackid, rackname, store) Then
        jsonObj.Add "success", False
        jsonObj.Add "message", "Shealth already exist in store!"
        Set AddRack = jsonObj
        Exit Function
    End If
    sql = "SELECT * FROM PerformVar22 WHERE PerformVar22ID = '" & rackid & "' and KeyPrefix = '" & store & "'"
    rst.open sql, conn, 3, 4
    If rst.recordCount = 1 Then
        rst.fields("PerformVar22Name") = rackname
    Else
        rst.AddNew
        recky = GetRecordKey("PerformVar22", "PerformVar22ID", "NONE")
        rst.fields("PerformVar22ID") = recky
        rst.fields("PerformVar22Name") = rackname
        rst.fields("Description") = "active"
        rst.fields("KeyPrefix") = store
    End If
    rst.fields("rack") = parent
    rst.updatebatch
    rst.Close
    Set rst = Nothing
    jsonObj.Add "success", True
    jsonObj.Add "message", "Saved!"
    Set AddRack = jsonObj
End Function

Function RackNameExist(id, rackname, store)
    Dim rst
    RackNameExist = False
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "SELECT * FROM PerformVar22 WHERE PerformVar22Name = '" & rackname & "' and KeyPrefix = '" & store & "'"
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        If rst.fields("PerformVar22Name") = rackname _
            And UCase(GetComboName("PerformVar22", id)) <> UCase(rackname) Then
            RackNameExist = True 'rst.fields("PerformVar22Name") = rackname
        End If
    Else
        RackNameExist = False
    End If
    rst.Close
    Set rst = Nothing
End Function



Function LoadRacks()
    Dim rst, store, JSONArray, jsonObj
    Set rst = CreateObject("ADODB.RecordSet")
    Set JSONArray = CreateObject("System.Collections.ArrayList")

    store = Trim(Request("store"))
    sql = "SELECT * FROM PerformVar22 WHERE KeyPrefix = '" & store & "' ORDER BY PerformVar22Name"
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            Set jsonObj = CreateObject("Scripting.Dictionary")
            jsonObj.Add "rackid", rst.fields("PerformVar22ID").value
            jsonObj.Add "rackname", rst.fields("PerformVar22Name").value
            jsonObj.Add "rackstatus", rst.fields("Description").value
            jsonObj.Add "store", rst.fields("KeyPrefix").value
            jsonObj.Add "parent", rst.fields("rack").value
            JSONArray.Add jsonObj
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    Set LoadRacks = JSONArray
End Function

Function DisableRack()
    Dim rackid, rst, jsonObj, store, status
    Set jsonObj = CreateObject("Scripting.Dictionary")
    rackid = Trim(Request("rack-id"))
    store = Trim(Request("store"))
    status = Trim(Request("status"))
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "SELECT * FROM PerformVar22 WHERE PerformVar22ID = '" & rackid & "' and KeyPrefix = '" & store & "'"
    rst.open sql, conn, 3, 4
    If rst.recordCount = 1 Then
        rst.fields("Description") = status
        rst.updatebatch
    End If
    rst.Close
    Set rst = Nothing
    jsonObj.Add "success", True
    jsonObj.Add "message", "Status Changed!"
    Set DisableRack = jsonObj
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
   Dim Key, tmpKey, value, tmp, ot, field
   Dim objType: objType = TypeName(obj)
   tmp = ""


   If IsObject(obj) Then

      If refList.Contains(obj) Then
             
      Else
         refList.Add obj

      End If
   End If
        
   If objType = "Dictionary" Then
      For Each Key In obj.Keys()
            If tmp <> "" Then tmp = tmp & ", "
            tmpKey = """" & Key & """"
            tmp = tmp & tmpKey & ":" & JSONStringify_(obj(Key), refList)
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
   Dim scriptEngine, tmp, arr, item, Key
   Set scriptEngine = CreateObject("MSScriptControl.ScriptControl")
   scriptEngine.Language = "JavaScript"

   scriptEngine.AddCode "function isArray(jsonObj){return Object.prototype.toString.call(jsonObj) === '[object Array]';}"
   scriptEngine.AddCode "function isObject(jsonObj){return (typeof jsonObj === 'object' || typeof jsonObj === 'function') && (jsonObj !== null);}"
   scriptEngine.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
   scriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
        
   If scriptEngine.Run("isArray", jScriptType) Then
      Set JSONParse_ = CreateObject("System.Collections.ArrayList")
      For Each item In jScriptType
            JSONParse_.Add JSONParse_(item, caseSensitivity)
      Next
   ElseIf scriptEngine.Run("isObject", jScriptType) Then
      Set JSONParse_ = CreateObject("Scripting.Dictionary")
      If caseSensitivity = False Then
            JSONParse_.CompareMode = 1
      End If

      For Each Key In scriptEngine.Run("getKeys", jScriptType)
            If JSONParse_.Exists(Key) Then
               JSONParse_.Remove Key
            End If
            JSONParse_.Add Key, JSONParse_(scriptEngine.Run("getProperty", jScriptType, Key), caseSensitivity)
      Next
   Else
      JSONParse_ = jScriptType
   End If
End Function

Function GetUserDrugStore(jSchd)
    Dim rst, sql, ot
    ot = ""
    Set rst = CreateObject("ADODB.Recordset")
    With rst
        sql = "WITH sto AS ( "
        sql = sql & " SELECT DrugStoreID, JobScheduleID AS [stoName] FROM DrugStore "
        sql = sql & " UNION ALL "
        sql = sql & " SELECT DrugStoreID, JobScheduleID AS [stoName] FROM DrugStore2 "
        sql = sql & ")"
        sql = sql & "SELECT top 1 DrugStoreID FROM sto WHERE stoName= '" & jSchd & "'"
        
        .maxrecords = 1
        .open qryPro.FltQry(sql), conn, 3, 4
        If .recordCount > 0 Then
            ot = rst.fields("DrugStoreID")
        End If
        .Close
    End With
    GetUserDrugStore = ot
    Set rst = Nothing
End Function








