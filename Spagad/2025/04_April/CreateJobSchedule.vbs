'<<--CONFIGURATION SCRIPT-->>
Dim replicateTbl, repWhCls, repPriKyLst, newPriKyVlLst

Pushing_P

Sub Pushing_P()
    Dim sql, source, target, targetName
    
    source = "M13"
    target = "dpt040"
    targetName = "Audit Heads"
    ReplicateProfile source, target, targetName
    
    'copy cashier access to s01b
    ReplicateAccess "M13", "dpt040"
End Sub


Sub ReplicateProfile(source, target, targetName)
    RepJobSchedule source, target
    UpdJobSchedule target, "JobScheduleName", targetName
    RepSystemUser source, target
    UpdSystemUser target, "JobScheduleID", target
    ResetUserPwd target
    ReplicateAccess source, target
    ReplicateModuleManager source, target
    ConvertNavToMenu3 target
End Sub

Sub ReplicateModuleManager(source, target)
    Dim sql, rst, rst2
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    sql = "select * from ModuleManagerAlloc where JobScheduleID='" & source & "' "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
        
            sql = "select * from ModuleManagerAlloc where JobScheduleID='" & target & "' "
            rst2.open sql, conn, 3, 4
            If rst2.RecordCount = 0 Then
                rst2.AddNew
                For Each field In rst2.fields
                    If UCase(field.name) = UCase("JobScheduleID") Then
                        rst2.fields(field.name) = target
                    Else
                        rst2.fields(field.name) = rst.fields(field.name)
                    End If
                Next
                rst2.updatebatch
            End If
            
            rst2.Close
            rst.MoveNext
        Loop
        
        rst.Close
    End If
    
End Sub
Sub GrantTblAccessToPrintlayout(tbl, printlayoutName)
    Dim sql, rst
    
    sql = "select UserRoleID from AccessRightAlloc where TableID='" & tbl & "' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            RepAnyPrintOutAlloc printlayoutName, rst.fields("UserRoleID")
            rst.MoveNext
        Loop
        rst.Close
    End If
    Set rst = Nothing
End Sub
Sub CopyPrintOutAccess(sourcePrint, targetPrint)
    Dim sql, rst
    
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "select * from PrintOutAlloc where PrintLayoutID='" & sourcePrint & "' "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            RepAnyPrintOutAlloc targetPrint, rst.fields("JobScheduleID")
            rst.MoveNext
        Loop
    End If
End Sub
Sub ReplicateAccess(sourceJb, targetJb)
    Dim sql, rst
    'tables
    sql = " select * from AccessRightAlloc where UserRoleID='" & sourceJb & "' "
    Set rst = CreateObject("ADODB.Recordset")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            RepAccessRightAlloc rst.fields("UserRoleID"), rst.fields("AccessRightID"), targetJb, rst.fields("AccessRightID")
            rst.MoveNext
        Loop
        rst.Close
    End If
    
    'printouts
    sql = " select * from PrintOutAlloc where JobScheduleID='" & sourceJb & "' "
    Set rst = CreateObject("ADODB.Recordset")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            RepPrintOutAlloc rst.fields("PrintLayoutID"), rst.fields("JobScheduleID"), rst.fields("PrintLayoutID"), targetJb
            rst.MoveNext
        Loop
        rst.Close
    End If
    
    'userrole
    sql = " select * from UserRoleAlloc where JobScheduleID='" & sourceJb & "' "
    Set rst = CreateObject("ADODB.Recordset")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            RepUserRoleAlloc rst.fields("UserRoleID"), rst.fields("JobScheduleID"), rst.fields("UserRoleID"), targetJb
            rst.MoveNext
        Loop
        rst.Close
    End If
    
End Sub


