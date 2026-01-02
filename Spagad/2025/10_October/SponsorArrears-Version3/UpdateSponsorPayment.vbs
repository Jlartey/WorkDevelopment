
Dim procedurename
procedurename = Trim(Request.QueryString("procedurename"))

If procedurename = "UpdateSponsorPayment" Then
    On Error Resume Next
    
    response.ContentType = "application/json"
    
    Dim recordId, amountDue, amountPaid
    recordId = Trim(Request.QueryString("id"))
    amountDue = Trim(Request.QueryString("amountDue"))
    amountPaid = Trim(Request.QueryString("amountPaid"))
    
    If recordId = "" Or amountPaid = "" Then
        response.write "{""success"": false, ""message"": ""Missing or invalid parameters (id, amountPaid)""}"
        response.End
    End If
    
    updatePerformVar16Record recordId, amountPaid
    
    If Err.number = 0 Then
        response.write "{""success"": true, ""message"": ""Record updated successfully""}"
    Else
        response.write "{""success"": false, ""message"": ""Error: " & Replace(Err.description, """", "\""") & """}"
        Err.Clear
    End If
    
    response.End
End If



Sub updatePerformVar16Record(recordId, amountPaid)
    Dim rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "SELECT * FROM PerformVar16 WHERE PerformVar16ID = '" & recordId & "'"
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    If rst.recordCount > 0 Then
        Dim currentAmount, newAmount
        currentAmount = CDbl(rst.fields("KeyPrefix").value)
        newAmount = currentAmount - CDbl(amountPaid)
        rst.fields("KeyPrefix") = CStr(newAmount)
        rst.UpdateBatch
    End If
    
    rst.Close
    Set rst = Nothing
End Sub

