Function getEMRResult(EMRRequestID, emrDataID, CompID, column)
    Dim sql, rst,emrValue
    Set rst = server.CreateObject("ADODB.Recordset")
    emrValue = ""
    sql = "SELECT * FROM emrresults WHERE emrrequestid ='" & EMRRequestID & "'"
    sql = sql & " AND emrdataid ='" & emrDataID & "' AND emrcomponentid='" & CompID & "'"
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            If Not IsNull(.fields(column)) Then
                emrValue = Trim(.fields(column))
            End If
            If IsNull(.fields(column)) Then
                emrValue = "Null"
            End If
        End If
        .Close
    End With
    getEMRResult = emrValue
    Set rst = Nothing
End Function
 47