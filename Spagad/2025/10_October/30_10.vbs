Function GetConsultingDoctor(visitationID)
    ' Strip hyphen and anything after it if present
    If InStr(visitationID, "-") > 0 Then
        visitationID = Left(visitationID, InStr(visitationID, "-") - 1)
    End If
    
    Dim sql, rst
    Set rst = Server.CreateObject("ADODB.Recordset")

    sql = "WITH DoctorConsults AS ("
    sql = sql & " SELECT DISTINCT SystemUserID, VisitationID"
    sql = sql & " From EMRRequestItems"
    sql = sql & " WHERE EMRDataID IN ('TH060', 'IM051', 'TH082', 'TH069', 'TH067', 'IM048', 'IM011', 'TH088', 'TH088B', 'TH085', 'TH072')"
    sql = sql & " )"
    sql = sql & " SELECT  Staff.StaffName"
    sql = sql & " FROM DoctorConsults dc"
    sql = sql & " JOIN SystemUser ON SystemUser.SystemUserID = dc.SystemUserID"
    sql = sql & " JOIN Staff ON Staff.StaffID = SystemUser.StaffID"
    sql = sql & " WHERE dc.VisitationID = '" & visitationID & "'"
   
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .recordCount > 0 Then
            GetConsultingDoctor = .fields("StaffName")
         Else
         GetConsultingDoctor = " "
        End If
    End With
    Set rst = Nothing
End Function