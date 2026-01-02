Dim procedurename
procedurename = Trim(Request.QueryString("procedurename"))

If procedurename = "InsertSponsorPayment" Then
    On Error Resume Next
    
    response.ContentType = "application/json"
    
    Dim SponsorID, billMonthID, amount
    SponsorID = Trim(Request.QueryString("sponsorID"))
    billMonthID = Trim(Request.QueryString("billMonthID"))
    amount = Trim(Request.QueryString("amount"))
    
    
    If SponsorID = "" Or billMonthID = "" Or amount = "" Then
        response.write "{""success"": false, ""message"": ""Missing or invalid parameters (sponsorID, billMonthID, amount)""}"
        response.End
    End If
    
   
    insertPerformVar16Record SponsorID, billMonthID, amount
    
    If Err.number = 0 Then
        response.write "{""success"": true, ""message"": ""Record inserted successfully""}"
    Else
        response.write "{""success"": false, ""message"": ""Error: " & Replace(Err.description, """", "\""") & """}"
        Err.Clear
    End If
    
    response.End
End If


Sub insertPerformVar16Record(SponsorID, billMonthID, amount)
    Dim rst, rck, sql
    Set rst = CreateObject("ADODB.Recordset")
    rck = GetRecordKey("PerformVar16", "PerformVar16ID", "NONE")
    
    
    sql = "SELECT * FROM PerformVar16 WHERE PerformVar16ID = '" & rck & "'"
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    If rst.recordCount = 0 Then
        rst.AddNew
        rst.fields("PerformVar16ID") = rck
        rst.fields("PerformVar16Name") = SponsorID
        rst.fields("Description") = billMonthID
        rst.fields("KeyPrefix") = amount
        rst.updatebatch
    End If
    
    rst.Close
    Set rst = Nothing
End Sub


