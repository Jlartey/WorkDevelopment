getEMRResult(EMRRequestID, "E000059", "E000193", "column2")
getEMRResult(EMRRequestID, "E000059", "E000194", "column2")
getEMRResult(EMRRequestID, "E000059", "E000195", "column2")
getEMRResult(EMRRequestID, "E000059", "E000196", "column2")

Function getEMRResult(EMRRequestID, emrDataID, CompID, column)

    Dim sql, rst
    Set rst = server.CreateObject("ADODB.Recordset")
    getEMRResult = ""
    
    sql = "SELECT * FROM emrresults WHERE emrrequestid ='" & EMRRequestID & "'"
    sql = sql & " AND emrdataid ='" & emrDataID & "' AND emrcomponentid='" & CompID & "'"
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            If Not IsNull(.fields(column)) Then
                getEMRResult = Trim(.fields(column))
            End If
            If IsNull(.fields(column)) Then
                getEMRResult = "Null"
            End If
        End If
    End With
    Set rst = Nothing
    
End Function