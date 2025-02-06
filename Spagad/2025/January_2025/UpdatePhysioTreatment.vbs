Dim treatmentDate, intervention, treatmentType, treatmentValue, resp, visitationId
treatmentDate = Trim(Request("treatmentDate"))
treatmentType = Trim(Request("treatmentType"))
intervention = Trim(Request("intervention"))
treatmentValue = Trim(Request("treatmentValue"))
visitationId = Trim(Request("VisitationID"))

Set resp = CreateObject("Scripting.Dictionary")
resp.Add "success", True
On Error Resume Next
    retValue = updatePerformVar11(treatmentDate, treatmentType, intervention, treatmentValue)
    If Err.number <> 0 Then
        resp("success") = False
    End If
On Error GoTo 0
response.Clear
response.contentType = "application/json"
response.write "{""success"":""" & resp("success") & """,""data"":" & retValue & "}"

Function updatePerformVar11(treatmentDate, treatmentType, intervention, treatmentValue)
    Dim rst, rck, json
    Set rst = CreateObject("ADODB.RecordSet")
    rck = GetRecordKey("PerformVar11", "PerformVar11ID", "NONE")
    sql = "SELECT * FROM PerformVar11 WHERE PerformVar11ID = '" & rck & "'"
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("PerformVar11ID") = rck
        rst.fields("PerformVar11Name") = treatmentType & "||" & visitationId
        rst.fields("Description") = intervention
        rst.fields("KeyPrefix") = treatmentDate & "||" & treatmentValue
        rst.Updatebatch
    End If
    rst.Close
    Set rst = Nothing
    json = "{""performvarID"":""" & rck & """,""performvarName"":""" & treatmentType & "||" & visitationId & """,""Description"":""" & intervention & """,""KeyPrefix"":""" & treatmentDate & "||" & treatmentValue & """}"
    updatePerformVar11 = json
End Function



