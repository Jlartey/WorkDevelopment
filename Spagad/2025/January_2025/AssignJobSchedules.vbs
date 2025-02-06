Dim replicateTbl, repWhCls, repPriKyLst, newPriKyVlLst
AssignJobSchedules

Sub AssignJobSchedules()
    Dim rst, sql
    
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "SELECT JobScheduleId FROM JobSchedule WHERE jobscheduleid IN ('S01', 'S02', 'S03')"
    
    With rst
        .Open sql, conn, 3, 4
        Do While Not .EOF
            RepAnyPrintOutAlloc "Greetings", .fields("JobScheduleId")
            .MoveNext
        Loop
        .Close
    End With
    
    Set rst = Nothing
    
End Sub

IMAH-23-002960