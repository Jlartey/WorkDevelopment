'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim dateRange, WorkingMonthID, drugId, agegrp
Dim rptGen, sql, args

Set rptGen = New PRTGLO_RptGen

WorkingMonthID = Trim(Request.QueryString("PrintFilter0"))
dateRange = Split(Trim(Request.QueryString("PrintFilter1")), "||")
agegrp = Trim(Request.QueryString("PrintFilter2"))
drugId = Trim(Request.QueryString("PrintFilter3"))

sql = GetDispenseReport
args = "title=drug dispense report;"
args = args & ";ShowColumnTotal=No"
rptGen.AddReport sql, args

rptGen.ShowReport

'PrintReport dateRange, WorkingMonthID, drugID
Function GetDispenseReport()
    If IsArray(dateRange) Then
        If UBound(dateRange) > 0 Then
            whcls = "       AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
        End If
    End If
    If WorkingMonthID <> "" Then whcls = whcls & "   AND WorkingMonthID='" & WorkingMonthID & "'"
    If drugId <> "" Then whcls = whcls & "    AND DrugID='" & drugId & "'"
    If agegrp <> "" Then whcls = whcls & " AND AgeGroupID='" & agegrp & "'"
    
    sql = " SELECT * "
    sql = sql & " FROM ("
        sql = sql & "   SELECT "
        sql = sql & "   PatientID AS [Folder No.]"
        sql = sql & "   , (SELECT PatientName FROM Patient WHERE Patient.PatientID=DrugSaleItems.PatientID) AS [Patient Name] "
        sql = sql & "   , ( SELECT VisitInfo6 FROM Visitation WHERE Visitation.VisitationID=DrugsaleItems.VisitationID) AS [Patient Age] "
        sql = sql & "   , ('['+DrugID + '] ' + ( SELECT DrugName FROM Drug WHERE Drug.DrugID=DrugSaleItems.DrugID)) AS [Drug Name]"
        sql = sql & "   , CAST(DispenseDate AS VARCHAR) AS [Date/time]"
        sql = sql & "    FROM DrugSaleItems "
        sql = sql & "    WHERE 1=1 AND DrugCategoryID<>'D002' "
        sql = sql & whcls
        sql = sql & " UNION ALL "
        sql = sql & " SELECT "
        sql = sql & "   PatientID AS [Folder No.]"
        sql = sql & "   , (SELECT PatientName FROM Patient WHERE Patient.PatientID=DrugSaleItems2.PatientID) AS [Patient Name] "
        sql = sql & "   , ( SELECT VisitInfo6 FROM Visitation WHERE Visitation.VisitationID=DrugsaleItems2.VisitationID) AS [Patient Age] "
        sql = sql & "   , ('['+DrugID + '] ' + ( SELECT DrugName FROM Drug WHERE Drug.DrugID=DrugSaleItems2.DrugID)) AS [Drug Name]"
        sql = sql & "   , CAST(DispenseDate AS VARCHAR) AS [Date/time]"
        sql = sql & "    FROM DrugSaleItems2 "
        sql = sql & "    WHERE 1=1 AND DrugCategoryID<>'D002' "
        sql = sql & whcls
    sql = sql & "  ) AS [Report]"
    sql = sql & " ORDER BY [Patient Name] ASC, [Drug Name]  "
    
    GetDispenseReport = sql
End Function
Sub PrintReport(dateRange, WorkingMonthID, drugId)
    
    Dim sql, whcls
    
    '''''
    
    If IsArray(dateRange) Then
        If UBound(dateRange) > 0 Then
            whcls = "       AND DispenseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
        End If
    End If
    
    If WorkingMonthID <> "" Then
        whcls = whcls & "   AND WorkingMonthID='" & WorkingMonthID & "'"
    End If
    If drugId <> "" Then
        whcls = whcls & "    AND DrugID='" & drugId & "'"
    End If
    
    ''''''
    
    sql = " SELECT * FROM ("
        sql = " SELECT "
        sql = sql & "   PatientID AS [Folder No.]"
        sql = sql & "   , (SELECT PatientName FROM Patient WHERE Patient.PatientID=DrugSaleItems.PatientID) AS [Patient Name] "
        sql = sql & "   , ( SELECT PatientAge FROM Visitation WHERE Visitation.VisitationID=DrugsaleItems.VisitationID) AS [Patient Age] "
        sql = sql & "   , ('['+DrugID + ']' + ( SELECT DrugName FROM Drug WHERE Drug.DrugID=DrugSaleItems.DrugID)) AS [Drug Name]"
        sql = sql & "   , DispenseDate AS [Date /time]"
        sql = sql & "    FROM DrugSaleItems "
        sql = sql & "    WHERE 1=1 AND DrugCategoryID<>'D002' "
        sql = sql & whcls
        sql = sql & " UNION ALL "
        sql = sql & " SELECT "
        sql = sql & "   PatientID AS [Folder No.]"
        sql = sql & "   , (SELECT PatientName FROM Patient WHERE Patient.PatientID=DrugSaleItems2.PatientID) AS [Patient Name] "
        sql = sql & "   , ( SELECT PatientAge FROM Visitation WHERE Visitation.VisitationID=DrugsaleItems2.VisitationID) AS [Patient Age] "
        sql = sql & "   , ('['+DrugID + ']' + ( SELECT DrugName FROM Drug WHERE Drug.DrugID=DrugSaleItems2.DrugID)) AS [Drug Name]"
        sql = sql & "   , DispenseDate AS [Date /time]"
        sql = sql & "    FROM DrugSaleItems2 "
        sql = sql & "    WHERE 1=1 AND DrugCategoryID<>'D002' "
        sql = sql & whcls
        sql = sql & " ) AS MReport "
    sql = sql & " ORDER BY [Date /time], [Patient Name] ASC, [Drug Name]  "
    
    response.write GetReportHeader(dateRange, WorkingMonthID, "Drug Dispense Report")
    response.write GenerateSQLReport(sql)

End Sub

Function GenerateSQLReport(sql)
    Dim rst, str, cnt
    
    Set rst = CreateObject("ADODB.RecordSet")
    
    rst.open sql, conn, 3, 4
    
    cnt = 1
    
    If rst.recordCount > 0 Then
        str = str & "<table border='1' cellspacing='0'>"
            str = str & "<thead>"
                str = str & "<tr>"
                    str = str & "<th>No.</th>"
                    For Each field In rst.fields
                        If UCase(field.name) = "PATIENT AGE" Then
                            str = str & "<td style='padding: 5px; '>" & UCase(field.name) & "</td>"
                        Else
                            str = str & "<th style='padding: 5px; min-width: 4.5cm;'>" & UCase(field.name) & "</th>"
                        End If
                    Next
                str = str & "</tr>"
            str = str & "</thead>"
            str = str & "<tbody>"
            
        Do While Not rst.EOF
            str = str & "<tr>"
                str = str & "<th>" & cnt & "</th>"
            For Each field In rst.fields
                str = str & "<td>" & rst.fields(field.name) & "</td>"
            Next
            
            str = str & "</tr>"
            
            cnt = cnt + 1
            rst.MoveNext
        Loop
        
        rst.Close
        Set rst = Nothing
        
        str = str & "</tbody>"
        str = str & "</table>"
    End If
    
    GenerateSQLReport = str
    
End Function

Function GetReportHeader(dateRange, wrkMnth, title)
    Dim str
    str = str & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
    str = str & "<tbody><tr><td align=""Center"" valign=""top""><img border=""0"" src=""images/logo.jpg"""
    str = str & "align=""Center""> </td><td align=""center"" height=""20"" bgcolor=""white"" "
    str = str & "style=""font-size: 14pt; color: #CCCCFF;  font-family:Arial"">"
    str = str & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" "
    str = str & "style=""font-family: Arial; font-size: 10pt; color: #000000"">"
    str = str & "<tbody><tr><td valign=""top""></td><td valign=""top""><b>IMaH&nbsp;&nbsp;HOSPITAL</b>"
    str = str & "</td><td valign=""top""></td><td valign=""top""><b>/</b></td></tr><tr>"
    str = str & "<td valign=""top""></td><td valign=""top""><b>Hospital</b></td>"
    str = str & "<td valign=""top""></td><td valign=""top""><b>/</b></td></tr><tr>"
    str = str & "<td valign=""top""></td><td valign=""top""><b>/</b></td>"
    str = str & "<td valign=""top""></td><td valign=""top""><b>/</b></td>"
    str = str & "</tr><tr><td valign=""top""></td><td valign=""top""><b>/</b></td>"
    str = str & "<td valign=""top""></td></tr><tr></tr></tbody></table></td></tr>"
    str = str & "<tr><td colspan=""2""><hr/></td><tr/>"
    str = str & "<tr><td colspan=""2"" style=""text-align: center;""><h2>" & title & "</h2></td><tr/>"
    
    If IsArray(dateRange) Then
        If UBound(dateRange) > 0 Then
            str = str & "<tr><td colspan=""2"" style=""text-align: center;"">From <span>" & dateRange(0) & "</span> To <span>" & dateRange(1) & "</span> </td><tr/>"
        End If
    End If
    
    If wrkMnth <> "" Then
        str = str & "<tr><td colspan=""2"" style=""text-align: center;""><b>AND</b> FOR THE MONTH OF <span>" & GetWorkingMonthName(wrkMnth) & "</span> </td><tr/>"
    End If
    
    str = str & "<tr><td colspan=""2""><hr/></td><tr/>"
    str = str & "</tbody></table>"
    
    GetReportHeader = str
    
End Function

Function GetWorkingMonthName(mth)

  Dim ot, ky

  ky = Trim(mth)
  ot = ""

  If Len(ky) = 9 Then
    If (UCase(Left(ky, 3)) = "MTH") And IsNumeric(Right(ky, 6)) Then
      ot = UCase(monthName(CLng(Right(ky, 2)), False) & " " & Mid(ky, 4, 4))
    End If
  End If

  GetWorkingMonthName = ot
  
End Function


'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
