'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim workMonth, dateRange
''' Print Dialysis Report
workMonth = Request.QueryString("PrintFilter0")
dateRange = Request.QueryString("PrintFilter1")

If dateRange <> "" Then
    dateRange = Split(dateRange, "||")
End If

PrintReport dateRange, workMonth

Sub PrintReport(dateRange, WorkingMonthID)
    '
    Dim sql, str, rst, emrDat, cnt, rptObj, dictParams, args
    
    'Set rptObj = New PRTGLO_kenGlobal
    Set rptObj = New PRTGLO_RptGen
    Set dictParams = CreateObject("Scripting.Dictionary")
    
    emrDat = "IM005"
    
    sql = sql & " SELECT "
    sql = sql & " vst.VisitDate AS [Date/Time] "
    sql = sql & " , e.VisitationID AS [Visit No] "
    sql = sql & " , e.PatientID AS [Patient ID] "
    sql = sql & " , pt.PatientName AS [Patient Name] "
    sql = sql & " , vst.PatientAge AS [Age] "
    sql = sql & " , ( SELECT GenderName FROM Gender WHERE pt.GenderID=Gender.GenderID ) AS [Sex] "
    sql = sql & " , a1.Column2 AS [Programme] "
    sql = sql & " , ( SELECT EMRVar3BName FROM EMRVar3B WHERE EMRVar3B.EMRVar3BID=CAST(a2.Column2 AS NVARCHAR(MAX)) ) AS [Access] "
    sql = sql & " , a3.Column6 AS [Hours Spent] "
    sql = sql & " , a4.Column2 AS [KT/V] "
    sql = sql & " , a5.Column4 AS [Heparin: Init Dose] "
    sql = sql & " , a6.Column6 AS [Heparin: Mnt. Dose] "
    sql = sql & " , a7.Column2 AS [Pre Weight Diff.] "
    sql = sql & " , a8.Column2 AS [Post Weight Diff.] "
    sql = sql & " , a9.Column2 AS [Dry Weight] "
    sql = sql & " , a10.Column4 AS [UF Rate] "
    sql = sql & " , a11.Column2 AS [HB] "
    sql = sql & " , a12.Column2 AS [Diagnosis] "
    sql = sql & " , a13.Column5 AS [Client Status] "
    
    
    sql = sql & " FROM EMRResults AS a1 "
    sql = sql & "   INNER JOIN EMRRequest AS e ON e.EMRRequestID=a1.EMRRequestID "
    sql = sql & "    INNER JOIN EMRResults AS a2 ON a1.EMRDataID=a2.EMRDataID AND a1.EMRRequestID=a2.EMRRequestID "
    sql = sql & "       AND a1.EMRComponentID='IM00513' AND a2.EMRComponentID='IM00506' "
    sql = sql & "       AND a1.EMRDataID='" & emrDat & "' AND a2.EMRDataID='" & emrDat & "'"
    
    sql = sql & "   INNER JOIN EMRResults AS a3 ON a1.EMRDataID=a3.EMRDataID AND a1.EMRRequestID=a3.EMRRequestID "
    sql = sql & "       AND a3.EMRComponentID='IM00501' "
    sql = sql & "       AND a3.EMRDataID='" & emrDat & "'"
    
    sql = sql & "   INNER JOIN EMRResults AS a4 ON a1.EMRDataID=a4.EMRDataID AND a1.EMRRequestID=a4.EMRRequestID "
    sql = sql & "       AND a4.EMRComponentID='IM00509' "
    sql = sql & "       AND a4.EMRDataID='" & emrDat & "'"
    
    sql = sql & "   INNER JOIN EMRResults AS a5 ON a1.EMRDataID=a5.EMRDataID AND a1.EMRRequestID=a5.EMRRequestID "
    sql = sql & "       AND a5.EMRComponentID='IM00510' "
    sql = sql & "       AND a5.EMRDataID='" & emrDat & "'"
    
    sql = sql & "   INNER JOIN EMRResults AS a6 ON a1.EMRDataID=a6.EMRDataID AND a1.EMRRequestID=a6.EMRRequestID "
    sql = sql & "       AND a6.EMRComponentID='IM00510' "
    sql = sql & "       AND a6.EMRDataID='" & emrDat & "'"
    
    sql = sql & "   INNER JOIN EMRResults AS a7 ON a1.EMRDataID=a7.EMRDataID AND a1.EMRRequestID=a7.EMRRequestID "
    sql = sql & "       AND a7.EMRComponentID='IM00510' "
    sql = sql & "       AND a7.EMRDataID='" & emrDat & "'"
    
    sql = sql & "   INNER JOIN EMRResults AS a8 ON a1.EMRDataID=a8.EMRDataID AND a1.EMRRequestID=a8.EMRRequestID "
    sql = sql & "       AND a8.EMRComponentID='IM00510' "
    sql = sql & "       AND a8.EMRDataID='" & emrDat & "'"
    
    sql = sql & "   INNER JOIN EMRResults AS a9 ON a1.EMRDataID=a9.EMRDataID AND a1.EMRRequestID=a9.EMRRequestID "
    sql = sql & "       AND a9.EMRComponentID='IM00502' "
    sql = sql & "       AND a9.EMRDataID='" & emrDat & "'"
    
    sql = sql & "   INNER JOIN EMRResults AS a10 ON a1.EMRDataID=a10.EMRDataID AND a1.EMRRequestID=a10.EMRRequestID "
    sql = sql & "       AND a10.EMRComponentID='IM00504' "
    sql = sql & "       AND a10.EMRDataID='" & emrDat & "'"

    sql = sql & "   LEFT JOIN EMRResults AS a12 ON a1.EMRDataID=a12.EMRDataID AND a1.EMRRequestID=a12.EMRRequestID "
    sql = sql & "       AND a12.EMRComponentID='IM00517' "
    sql = sql & "       AND a12.EMRDataID='" & emrDat & "'"

    sql = sql & "   LEFT JOIN EMRResults AS a13 ON a1.EMRDataID=a13.EMRDataID AND a1.EMRRequestID=a13.EMRRequestID "
    sql = sql & "       AND a13.EMRComponentID='IM00517' "
    sql = sql & "       AND a13.EMRDataID='" & emrDat & "'"
    
    sql = sql & "   LEFT JOIN LabResults AS a11 "
    sql = sql & "       ON a11.LabRequestID=("
    sql = sql & "               SELECT TOP 1 LabRequestID FROM LabRequest WHERE "
    sql = sql & "               EXISTS( SELECT LabRequestID FROM Investigation WHERE LabRequest.LabRequestID=Investigation.LabRequestID AND Investigation.LabTestID='LT145' AND Investigation.VisitationID=e.VisitationID ) "
    sql = sql & "               OR EXISTS( SELECT LabRequestID FROM Investigation2 WHERE LabRequest.LabRequestID=Investigation2.LabRequestID AND Investigation2.LabTestID='LT145' AND Investigation2.VisitationID=e.VisitationID ) "
    sql = sql & "       ) "
    sql = sql & "       AND a11.TestComponentID='L0467' "
    sql = sql & "       AND a11.LabTestID='LT145'"
    sql = sql & "   INNER JOIN Visitation AS vst ON vst.VisitationID=e.VisitationID "
    sql = sql & "   INNER JOIN Patient AS pt ON pt.PatientID=e.PatientID"
    
    sql = sql & "  WHERE 1=1"
    If IsArray(dateRange) Then
        If UBound(dateRange) > 0 Then
            sql = sql & "       AND e.EMRDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
        End If
    End If
    
    If WorkingMonthID <> "" Then
        sql = sql & "       AND e.WorkingMonthID='" & WorkingMonthID & "'"
    End If
    
    sql = sql & " ORDER BY vst.VisitDate DESC, a1.EMRRequestID ASC, a1.CompPos ASC "
    
    response.write GetReportHeader(dateRange, WorkingMonthID, "Dialysis Observation")
    
'    dictParams.CompareMode = 1
'    dictParams.Add "ShowRowTotal", False
'    dictParams.Add "ShowColumnTotal", False
'    dictParams.Add "heading", "Dialysis Observation"
'    rptObj.PrintSQLReport sql, Nothing, dictParams
    args = "title=Dialysis Observation;ShowColumnTotal=False;"
    rptObj.PrintSQLReport sql, args
End Sub

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
