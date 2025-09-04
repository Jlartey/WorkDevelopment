'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim dateRange
dateRange = Request.QueryString("PrintFilter0")
If dateRange = "" Then
    dateRange = FormatDate(Now()) & " 00:00:00||" & FormatDate(Now()) & " 23:59:59"
End If
dateRange = Split(dateRange, "||")

PrintHeader dateRange
PrintMorgueReport dateRange


Function PrintMorgueReport(dateRange)
    Dim rst, sql, str
    
    sql = " SELECT PatientID, MortuaryName, DepositDate, ReleaseDate, NoOfDays "
    sql = sql & " FROM Mortuary WHERE 1=1 "
    sql = sql & "   AND ReleaseDate IS NOT NULL "
    sql = sql & "   AND DepositDate BETWEEN '" & dateRange(0) & "'  AND '" & dateRange(1) & "'"
    
    
    'response.write sql & "<br/>"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    
    If rst.recordCount > 0 Then
        rst.MoveFirst
        
        str = "<table border='1' cellspacing='0' style='width: 800px;'><thead>"
        str = str & "<tr>"
            str = str & "<th> PatientID </th>"
            str = str & "<th> Name </th>"
            str = str & "<th> Deposit Date </th>"
            str = str & "<th> Release Date </th>"
            str = str & "<th> No. Of Days spent </th>"
        str = str & "</tr>"
        
        Do While Not rst.EOF
            str = str & "<tr>"
                str = str & "<td style='text-align: right;' padding-left: 10px; padding-top: 5px>" & (rst.fields("PatientID")) & "</td>"
                str = str & "<td style='padding-left: 10px; padding-top: 5px;'>" & GetName(rst.fields("MortuaryName")) & "</td>"
                str = str & "<td style='text-align: center; padding-left: 10px; padding-top: 5px'>" & FormatDate(rst.fields("DepositDate")) & "</td>"
                str = str & "<td style='text-align: center; padding-left: 10px; padding-top: 5px'>" & FormatDate(rst.fields("ReleaseDate")) & "</td>"
                str = str & "<td style='text-align: right;' padding-left: 10px; padding-top: 5px>" & (rst.fields("NoOfDays")) & "</td>"
            str = str & "</tr>"
            rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing
        
        str = str & "</table>"
    End If
    
    response.write str
End Function

Sub PrintHeader(dateRange)
    response.write "<table style='width: 800px; '>"
        response.write "<tr>"
            AddReportHeader
        response.write "</tr>"
        response.write "<tr>"
            response.write "<td style='text-align: center; font-weight: bold; font-size: larger;'>Mortuary Report</td>"
        response.write "</tr>"
        response.write "<tr>"
             response.write "<td style='text-align: center;'><b>From</b> " & dateRange(0) & "  <b>to</b> " & dateRange(1) & "</td>"
        response.write "</tr>"
    response.write "</table>"
    response.write "<br/>"
    
End Sub

Function GetName(str)
    GetName = Left(str, InStr(str, "[") - 1)
    
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
