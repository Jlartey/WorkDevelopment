Sub ReportHeaderDetails()
    response.write "<div style=""text-align: center;"">"
    response.write "<div style=""width: " & (PrintWidth) & "px; margin: 0 auto;"">"
    Glob_AddReportHeader
    response.write "</div>"
    response.write "</div>"
    response.write "<div style=""text-align: center;""><hr color=""#999999"" size=""1""></div>"
    response.write "<div style=""text-align: center;""><hr color=""#999999"" size=""1""></div>"
    response.write "<div style=""display: flex; justify-content: center; align-items: center; gap: 30px; width: 100%; margin: 0 auto;"">"
    response.write "<span style=""font-family: Arial; color: #111111; font-size: 10pt;"">" & dateRange & "</span>"
    response.write "<span style=""font-family: Arial; color: #111111; font-size: 10pt;"">" & "Printed by " & GetStaff() & " on " & FormatDateDetail(Now) & "</span>"
    response.write "</div>"
    response.write "<div style=""text-align: center;""><hr color=""#999999"" size=""1""></div>"
End Sub