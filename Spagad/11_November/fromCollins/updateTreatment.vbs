Dim treatDate,intervention,treatType,treatValue
treatDate = Trim(Request("treatment-date"))
treatType = Trim(Request("type"))
intervention = Trim(Request("intervention"))
treatValue = Trim(Request("treatValue"))
visitID = Trim(Request("VisitID"))

Sub updatePerformVar20(tDate,ttype,intervene,tValue,vID)
    Dim rst
    Set rst = CreateObject("ADODB.RecordSet")
    rck = GetRecordKey("PerformVar20","PerformVar20ID","NONE")
    sql = "SELECT * FROM PerformVar20 where Performvar20ID = '" & rck & "'"
    rst.open sql,conn,3,4
    if rst.RecordCount = 0 then
        rst.AddNew
        rst.fields("PerformVar20ID") = rck & "||" & vID
        rst.fields("PerformVar20Name") = intervene
        rst.fields("Description") = 
        rst.updatebatch
    end if
    rst.close
    set rst = nothing
End Sub