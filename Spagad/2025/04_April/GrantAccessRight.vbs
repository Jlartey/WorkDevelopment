Config
Sub Config()
	GrantAccessToTable "S01", "Receipt", "new"
	GrantAccessToTable "S01", "Receipt", "view"
	GrantAccessToTable "S01", "Receipt", "edit"
End Sub

Sub GrantAccessToTable(profile, tableName, uRight)
    Dim sql, rst, accessRight, userRight
    accessRight = ""
    
    Select Case UCase(uRight)
        Case "VIEW"
            userRight = "URT001"
            accessRight = "frm" & tableName & userRight
        Case "NEW"
            userRight = "URT002"
            accessRight = "frm" & tableName & userRight
        Case "EDIT"
            userRight = "URT003"
            accessRight = "frm" & tableName & userRight
        Case "SAVE"
            userRight = "URT004"
            accessRight = "frm" & tableName & userRight
        Case "DELETE"
            userRight = "URT005"
            accessRight = "frm" & tableName & userRight
        Case "SEARCH"
            userRight = "URT006"
            accessRight = "frm" & tableName & userRight
    End Select
    
    If Len(accessRight) > 0 Then
        sql = "select * from AccessRightAlloc where UserRoleID='" & profile & "' and AccessRightID='" & accessRight & "' "
        Set rst = CreateObject("ADODB.RecordSet")
        rst.open sql, conn, 3, 4
        If rst.RecordCount = 0 Then
            rst.AddNew
            rst.fields("UserRoleID") = profile
            rst.fields("AccessRightID") = accessRight
            rst.fields("UserRightID") = userRight
            rst.fields("TableID") = tableName
            rst.fields("AccessDetail") = "YES"
            rst.UpdateBatch
        End If
        rst.Close
    End If
    
End Sub


