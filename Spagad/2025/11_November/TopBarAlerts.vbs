

GetAlerts

Function GetAlerts()
    Dim ot, tmp
    Set ot = CreateObject("System.Collections.ArrayList")
    
    '//global alerts
    'structure: [{url:''}, {url:'', message:'', title: '', count:number}]
    
    'my pending Transition actions
    ot.Add GetItemsPendingMyActions(uname)
     
    'profile-specific alerts
    Set tmp = CreateObject("Scripting.Dictionary")
    tmp.Add "url", "wpgXMLHttp.asp?ProcedureName=topBarAlerts" & jSchd
    ot.Add tmp
    
    If IsObject(response) Then
        response.Clear
        response.ContentType = "application/json"
        response.write JSONStringify(ot)
    End If
    
    Set ot = Nothing
End Function
Function GetItemsPendingMyActions(user)
    Dim sql, ot, rst
    
    Set ot = CreateObject("Scripting.Dictionary")
    
    sql = GetMyPendingActionListSQl(user, False, "")
    If Len(sql) > 0 Then
        sql = "select sum(total) as [total] from ( " & sql & ") as tmp "
        Set rst = CreateObject("ADODB.RecordSet")
        rst.open sql, conn, 3, 4
        If rst.recordCount > 0 Then
            rst.MoveFirst
            
            If IsNumeric(rst.fields("total")) Then
                ot.Add "title", "Approval / Action Required"
                ot.Add "count", rst.fields("total").value
                ot.Add "url", "wpgPrtPrintLayoutAll.asp?PrintLayoutName=MyPendingActions&PositionForTableName=WorkingDay"
                ot.Add "class", "danger"
            End If
        End If
        rst.Close
        Set rst = Nothing
    End If
    Set GetItemsPendingMyActions = ot
End Function
Function GetMyPendingActionListSQl(UserName, detailsQuery, whichTable)
    Dim ot, sql, sql0, tmpSQL, rst, r

    Set rst = CreateObject("ADODB.RecordSet")

    '''START TRANSITION
    sql = sql & " select TableID, SourceTableID, string_agg(TransProcessStatID, ',') as TransitionsFrom"
    sql = sql & " from ("
    sql = sql & "   select TransProcessorAcc2.TableID, TransProcessorAcc2.TransProcessStatID"
    sql = sql & "       , TransactionProcess.SourceTableID"
    sql = sql & "   from TransProcessorAcc2"
    sql = sql & "   inner join TransactionProcess on TransactionProcess.TableID=TransProcessorAcc2.TableID"
    sql = sql & "   inner join TransProcessVal on TransProcessVal.TransProcessValID=TransProcessorAcc2.TableID + '-' + TransProcessorAcc2.TransProcessStatID"
    sql = sql & "   where InitialScheduleID='" & GetComboNameFld("SystemUser", UserName, "JobScheduleID") & "' "
    sql = sql & "       and (TransProcessVal.KeyPrefix<>'No' or TransProcessVal.KeyPrefix is null ) " 'show in my actions
    If Len(whichTable) > 0 Then
        sql = sql & "   and TransactionProcess.SourceTableID='" & whichTable & "' "
    End If

    sql = sql & "   union "
    sql = sql & "   select TransProcessorAcc.TableID, TransProcessorAcc.TransProcessStatID"
    sql = sql & "       , TransactionProcess.SourceTableID"
    sql = sql & "   from TransProcessorAcc"
    sql = sql & "   inner join TransactionProcess on TransactionProcess.TableID=TransProcessorAcc.TableID"
    sql = sql & "   inner join TransProcessVal on TransProcessVal.TransProcessValID=TransProcessorAcc.TableID + '-' + TransProcessorAcc.TransProcessStatID"
    sql = sql & "   where InitialSystemUserID='" & UserName & "' "
    sql = sql & "       and (TransProcessVal.KeyPrefix<>'No' or TransProcessVal.KeyPrefix is null ) " 'show in my actions
    If Len(whichTable) > 0 Then
        sql = sql & "   and TransactionProcess.SourceTableID='" & whichTable & "' "
    End If
    sql = sql & " ) as UserTransAccess"
    sql = sql & "   group by TableID, SourceTableID"

    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            Set r = GetExtraSQL(rst.fields("SourceTableID"))
            If Len(r("SourceKey")) > 0 Then

                If detailsQuery = True Then
                    sql0 = ""
                    sql0 = sql0 & " select"
                    sql0 = sql0 & "     '@Table' as [TableID]"
                    sql0 = sql0 & "     , '@SourceTable' as SourceTableID"
                    sql0 = sql0 & "     , Tables.DisplayName as [SourceTableDisplayName] "
                    sql0 = sql0 & "     , " & r("SourceKey") & " as [SourceKey]"
                    sql0 = sql0 & "     , " & r("SourceKeyVal") & " as [SourceKeyVal]"
                    sql0 = sql0 & "     , " & r("SourceName") & " as [SourceName]"
                    sql0 = sql0 & "     , " & r("staffSelect")
                    sql0 = sql0 & "     , " & r("DaySelect")
                    sql0 = sql0 & "     , @SourceTable.TransProcessValID"
                    sql0 = sql0 & "     , @SourceTable.TransProcessStatID"
                    sql0 = sql0 & "     , TransProcessVal.TransProcessValName"
                    sql0 = sql0 & " from @SourceTable"
                    sql0 = sql0 & " outer apply ("
                    sql0 = sql0 & "     select top 1 @Table.* "
                    sql0 = sql0 & "     from @Table "
                    sql0 = sql0 & "     where @Table.TransProcessVal2ID=@SourceTable.TransProcessValID"
                    sql0 = sql0 & "         and " & r("ProTableWhere")
                    sql0 = sql0 & "     order by @Table.TransProcessDate1 desc "
                    sql0 = sql0 & " ) as @Table "
                    sql0 = sql0 & " inner join Tables on Tables.TableID='@SourceTable' "
                    sql0 = sql0 & " inner join TransProcessVal on TransProcessVal.TransProcessValID=@SourceTable.TransProcessValID"
                    sql0 = sql0 & r("dayJoin")
                    sql0 = sql0 & r("BranchJoin")
                    sql0 = sql0 & r("staffJoin")
                    sql0 = sql0 & " where @SourceTable.TransProcessStatID in ('" & Replace(rst.fields("TransitionsFrom"), ",", "', '") & "') "
                    sql0 = sql0 & r("dayWhere")
                    sql0 = sql0 & r("BranchWhere")
                    sql0 = sql0 & r("ExtraWhere")
                Else
                    sql0 = ""
                    sql0 = sql0 & " select "
                    sql0 = sql0 & "     '@Table' as [TableID]"
                    sql0 = sql0 & "     , '@SourceTable' as SourceTableID"
                    sql0 = sql0 & "     , Tables.DisplayName as [SourceTableDisplayName] "
                    sql0 = sql0 & "     , count(*) as [Total]"
                    sql0 = sql0 & " from @SourceTable"
                    sql0 = sql0 & " inner join Tables on Tables.TableID='@SourceTable' "
                    sql0 = sql0 & r("dayJoin")
                    sql0 = sql0 & r("BranchJoin")
                    sql0 = sql0 & " where @SourceTable.TransProcessStatID in ('" & Replace(rst.fields("TransitionsFrom"), ",", "', '") & "') "
                    sql0 = sql0 & r("dayWhere")
                    sql0 = sql0 & r("BranchWhere")
                    sql0 = sql0 & r("ExtraWhere")
                    sql0 = sql0 & " group by Tables.DisplayName"
                End If

                sql0 = Replace(sql0, "@SourceTable", rst.fields("SourceTableID"))
                sql0 = Replace(sql0, "@Table", rst.fields("TableID"))

                If Len(sql0) > 0 Then
                    If Len(ot) > 0 Then
                        ot = ot & vbCrLf & " union all " & vbCrLf
                    End If
                    ot = ot & sql0
                End If
            End If

            rst.MoveNext
        Loop
    End If
    rst.Close
    GetMyPendingActionListSQl = ot
End Function
Function GetExtraSQL(table)
    Dim rstFlds, ot, sql, wkDy

    Set ot = CreateObject("Scripting.Dictionary")
    ot.CompareMode = 1
    Set rstFlds = CreateObject("ADODB.RecordSet")

    sql = "select * from TableField where TableID='" & table & "' "
    rstFlds.open sql, conn, 3, 4

    If True Then
        ot("SourceKey") = ""
        rstFlds.filter = " PrimaryKey='Yes' "
        If rstFlds.recordCount > 0 Then
            Do While Not rstFlds.EOF
                If Len(ot("SourceKey")) > 0 Then
                    ot("SourceKey") = ot("SourceKey") & " + '<->' + "
                End If
                If Len(ot("SourceKeyVal")) > 0 Then
                    ot("SourceKeyVal") = ot("SourceKeyVal") & " + '&' + "
                End If
                If Len(ot("ProTableWhere")) > 0 Then
                    ot("ProTableWhere") = ot("ProTableWhere") & " and "
                End If
                ot("SourceKey") = ot("SourceKey") & " @SourceTable." & rstFlds.fields("TableFieldID")
                ot("SourceKeyVal") = ot("SourceKeyVal") & "'" & rstFlds.fields("TableFieldID") & "=' + @SourceTable." & rstFlds.fields("TableFieldID")
                ot("ProTableWhere") = ot("ProTableWhere") & " @SourceTable." & rstFlds.fields("TableFieldID") & "=@Table." & rstFlds.fields("TableFieldID")
                rstFlds.MoveNext
            Loop
        End If

        ot("SourceName") = ""
        rstFlds.filter = " TransientField='Yes' "
        If rstFlds.recordCount > 0 Then
            Do While Not rstFlds.EOF
                If Len(ot("SourceName")) > 0 Then
                    ot("SourceName") = ot("SourceName") & " + '<->' + "
                End If
                ot("SourceName") = ot("SourceName") & " @SourceTable." & rstFlds.fields("TableFieldID")
                rstFlds.MoveNext
            Loop
        End If
    End If

    If True Then
        ot("staffJoin") = ""
        rstFlds.filter = " TableFieldID='SystemUserID' "
        If rstFlds.recordCount > 0 Then
            ot("staffJoin") = ot("staffJoin") & " inner join SystemUser on SystemUser.SystemUserID=@SourceTable.SystemUserID"
            ot("staffJoin") = ot("staffJoin") & " inner join Staff on Staff.StaffID=SystemUser.StaffID"
            ot("staffSelect") = " Staff.StaffName as [EntryBy]"
        Else
            '??
        End If

        If Len(ot("staffJoin")) > 0 Then
        Else
            ot("staffSelect") = " '' as [EntryBy]"
        End If
    End If

    If True Then
        wkDy = FormatWorkingDay(DateAdd("d", -365, Now))
        ot("DayJoin") = ""
        rstFlds.filter = " TableFieldID='WorkingDayID' "
        If rstFlds.recordCount > 0 Then
            ot("DayJoin") = " inner join WorkingDay on WorkingDay.WorkingDayID=@SourceTable.WorkingDayID"
            ot("DaySelect") = " @SourceTable.WorkingDayID, WorkingDay.WorkingDayName"
            ot("DayWhere") = " and @SourceTable.WorkingDayID >= '" & wkDy & "' "
        Else
            rstFlds.filter = " TableFieldID='FirstDayID' "
            If rstFlds.recordCount > 0 Then
                ot("DayJoin") = " inner join FirstDay on FirstDay.FirstDayID=@SourceTable.FirstDayID"
                ot("DaySelect") = " @SourceTable.FirstDayID as WorkingDayID, FirstDay.FirstDayName as WorkingDayName"
                ot("DayWhere") = " and @SourceTable.FirstDayID >= '" & wkDy & "' "
            Else
                '??
            End If
        End If

        If Len(ot("DayJoin")) > 0 Then
        Else
            ot("DaySelect") = " '' as WorkingDayID, '' as WorkingDayName"
        End If
    End If

    If True Then
        ot("BranchWhere") = ""

        rstFlds.filter = " TableFieldID='TaxPayerBranchID'"
        If rstFlds.recordCount > 0 Then
            rstFlds.filter = " TableFieldID='TaxPayerID' "
            If rstFlds.recordCount > 0 Then
                ot("BranchJoin") = " inner join TaxPayer as txB on txB.TaxPayerBranchID=@SourceTable.TaxPayerBranchID and txB.TaxPayerID=@SourceTable.TaxPayerID"
                ot("BranchWhere") = " and txB.TaxPayerBranchID='" & brnch & "'"
            Else
                ot("BranchWhere") = " and @SourceTable.TaxPayerBranchID='" & brnch & "'"
            End If
        End If

        If Not (Len(ot("BranchWhere")) > 0) Then
            rstFlds.filter = " TableFieldID='BranchID'"
            If rstFlds.recordCount > 0 Then
                ot("BranchWhere") = " and @SourceTable.BranchID='" & brnch & "'"
            End If
        End If
    End If
    
    ''' SPECIAL CASES
    '''
    ot("ExtraWhere") = ""
    If UCase(table) = "DEBITADJUSTMENT" Then
        ot("ExtraWhere") = ot("ExtraWhere") & " and @SourceTable.DebitAdjustStatusID='D002' "
    ElseIf UCase(table) = "CREDITADJUSTMENT" Then
        ot("ExtraWhere") = ot("ExtraWhere") & " and @SourceTable.CreditAdjustStatusID='C002' "
    End If

    Set GetExtraSQL = ot
End Function




Function JSONStringify(obj)
        Dim refList, sc
        Set refList = CreateObject("Scripting.Dictionary")
        
        JSONStringify = JSONStringify_(obj, refList)
        Set refList = Nothing
    End Function
Function JSONStringify_(ByRef obj, ByRef refList)
    Dim Key, tmpKey, value, tmp, ot, field
    Dim objType: objType = TypeName(obj)

    tmp = ""
    If objType = "Dictionary" Then
        For Each Key In obj.Keys()
            If tmp <> "" Then tmp = tmp & ", "
            tmpKey = """" & Key & """"
            tmp = tmp & tmpKey & ":" & JSONStringify_(obj(Key), refList)
        Next
        ot = "{" & tmp & "}"
    ElseIf objType = "Fields" Then 'ADODB.Fields
        For Each field In obj
            If tmp <> "" Then tmp = tmp & ", "
            tmpKey = """" & field.name & """"
            tmp = tmp & tmpKey & ":" & JSONStringify_(field.value, refList)
        Next
        ot = "{" & tmp & "}"
    ElseIf IsArray(obj) Or objType = "ArrayList" Then
        For Each value In obj
            If tmp <> "" Then tmp = tmp & ", "
            tmp = tmp & JSONStringify_(value, refList)
        Next
        ot = "[" & tmp & "]"
    ElseIf objType = "String" Then
        tmp = Replace(obj, "\", "\\")
        tmp = Replace(tmp, """", "\""")
        tmp = Replace(tmp, vbTab, "\t")
        tmp = Replace(tmp, vbCrLf, "\r\n")
        tmp = Replace(tmp, vbCr, "\r")
        tmp = Replace(tmp, vbLf, "\n")
        ot = """" & tmp & """"
    ElseIf objType = "Boolean" Then
        ot = "" & LCase(obj) & ""
    ElseIf objType = "Byte" Then
        ot = CDbl(obj) 'Compatible with JSON.parse
    ElseIf objType = "Integer" Or objType = "Double" Or objType = "Long" Or objType = "Single" Or objType = "Currency" Then
        ot = obj
    ElseIf objType = "Empty" Or objType = "Null" Then
        ot = "null"
    ElseIf objType = "Date" Then
        ot = """" & obj & """"
    Else
        ot = """[Object " & objType & "]"""
    End If

    JSONStringify_ = ot
End Function
    





    


