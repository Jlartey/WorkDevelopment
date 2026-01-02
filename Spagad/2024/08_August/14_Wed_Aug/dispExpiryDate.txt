Dim drugID, batchNum
drugID = Trim(request.querystring("drugID"))
batchNum = Trim(request.querystring("batchNum"))

dispExpiryDate drugID, batchNum

Sub dispExpiryDate(drugID, batchNum)
    Dim sql, rst, expiryDate
    
    sql = "SELECT ExpiryDate FROM IncomingDrugItems WHERE DrugID = '" & drugID & "' AND PurchaseOrderInfo1 = '" & batchNum & "'"
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4

    If Not rst.EOF Then
        expiryDate = rst.fields("ExpiryDate")
        response.write expiryDate
    Else
        response.write "N/A" ' Return N/A if no expiry date is found
    End If
    
    rst.Close
    Set rst = Nothing
End Sub





'Sub dispExpiryDate(drugID, batchNum)
'    Dim sql, rst
'
'    sql = "SELECT ExpiryDate FROM IncomingDrugItems WHERE DrugID = '" & drugID & "' AND PurchaseOrderInfo1 = '" & batchNum & "'"
'
'    Set rst = CreateObject("ADODB.RecordSet")
'    rst.open sql, conn, 3, 4
'
'    With rst
'        If .RecordCount > 0 Then
'            .movefirst
'
'            Do While Not .EOF
'                response.write "<td align=""right"">" & .fields("ExpiryDate") & "</td>"
'                .MoveNext
'            Loop
'        End If
'    End With
'
'    rst.Close
'    Set rst = Nothing
'End Sub

