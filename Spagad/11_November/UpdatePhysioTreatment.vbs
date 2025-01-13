
Dim treatmentDate, intervention, treatmentType, treatmentValue
treatmentDate = Trim(Request("treatmentDate"))
treatmentType = Trim(Request("treatmentType"))
intervention = Trim(Request("intervention"))
treatmentValue = Trim(Request("treatmentValue"))

Sub updatePerformVar11(treatmentDate, treatmentType, intervention, treatmentValue)
    Dim rst
    Set rst = CreateObject("ADODB.RecordSet")
    rck = GetRecordKey("PerformVar11", "PerformVar11ID", "NONE")
    sql = "SELECT * FROM PerformVar11 where Performvar11ID = '" & rck & "'"
    rst.open sql, conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("PerformVar11ID") = rck
        rst.fields("PerformVar11Name") = treatmentType
        rst.fields("Description") = intervention
        rest.fields("KeyPrefix") = treatmentDate & "||" & treatmentValue
        rst.updatebatch
    End If
    rst.Close
    Set rst = Nothing
End Sub


A more elegant solutio

































Dim treatmentDate, treatmentType, intervention, treatmentValue

treatmentDate = Trim(Request("treatmentDate"))
treatmentType = Trim(Request("treatmentType"))
intervention = Trim(Request("intervention"))
treatmentValue = Trim(Request("treatmentValue"))

Sub updatePerformVar11
    Dim rst, rck, sql
    rst = CreateObject("ADODB.RecordSet")
    rst.open conn, 3, 4

    rck = GetRecordKey("PerformVar11", "PerformVar11ID", "NONE")

    sql = "Select * from PerformVar11 where "
    If rst.RecordCount = 0 Then

End Sub