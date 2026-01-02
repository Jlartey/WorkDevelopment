'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
addCSS
Dim rst, sql, cnt, SponsorID, sponsorName, MonthID
Set rst = CreateObject("ADODB.Recordset")
MonthID = Trim(Request.QueryString("MonthID"))
sql = "SELECT * FROM WorkingMonth ORDER BY WorkingMonthID DESC"
rst.open qryPro.FltQry(sql), conn, 3, 4
    response.write vbCrLf & "<div id='nav-container'>"
    response.write vbCrLf & "   <select name='month' id='month' onchange='filterByMonth()'>"
    If MonthID = "" Then
        response.write vbCrLf & "       <option value=''>Select Month</option>"
    Else
        response.write vbCrLf & "       <option value='" & MonthID & "'>" & GetComboName("WorkingMonth", MonthID) & "</option>"
    End If
    Do While Not rst.EOF
    response.write vbCrLf & "       <option value='" & rst.fields("WorkingMonthID") & "'>" & rst.fields("WorkingMonthName") & "</option>"
        rst.MoveNext
    Loop
    response.write vbCrLf & "   </select>"
    response.write vbCrLf & "   <div>"
    response.write vbCrLf & "      <button id='btn-profoma' onclick=""openModal('wpgPerformVar22.asp?PageMode=AddNew')"">+ Company</button>"
    response.write vbCrLf & "   </div>"
    response.write vbCrLf & "</div>"
    response.write vbCrLf & "<div id='report-container'>"
        ProcessCode
    response.write vbCrLf & "</div>"
rst.Close
Set rst = Nothing

response.write vbCrLf & "<script>"
response.write vbCrLf & "   function filterByMonth(){"
response.write vbCrLf & "      const month = document.querySelector('#month').value;"
response.write vbCrLf & "      const currentUrl = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=ClaimsProformaSponsor&PositionForTableName=WorkingDay&WorkingDayID=';"
response.write vbCrLf & "       if (month) {"
response.write vbCrLf & "          window.location.href = currentUrl + '&monthID=' + month;"
response.write vbCrLf & "       };"
response.write vbCrLf & "   };"
response.write vbCrLf & "   function openModal(url){"
response.write vbCrLf & "      window.open(url, '_blank', 'width=800,height=600');"
response.write vbCrLf & "   };"
response.write vbCrLf & "</script>"

Sub ProcessCode()
    Dim rst, sql, ot, cnt
    cnt = 0
    Set rst = CreateObject("ADODB.Recordset")
    sql = "SELECT * FROM PerformVar22 ORDER BY PerformVar22Name ASC"
    rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
            response.write vbCrLf & "<table id='myTable'>"
            response.write vbCrLf & "   <thead>"
            response.write vbCrLf & "   <tr>"
            response.write vbCrLf & "      <td>No.</td>"
            response.write vbCrLf & "      <td>Date</td>"
            response.write vbCrLf & "      <td>Company</td>"
            response.write vbCrLf & "      <td>Prepare Invoice</td>"
            response.write vbCrLf & "      <td>View Invoice</td>"
            response.write vbCrLf & "   </tr>"
            response.write vbCrLf & "   </thead>"
            response.write vbCrLf & "   <tbody>"
            Do While Not rst.EOF
                SponsorID = rst.fields("PerformVar22ID")
                sponsorName = rst.fields("PerformVar22Name")
                dt = FormatWorkingMonth(rst.fields("KeyPrefix"))
                hrf = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=generateProformaSponsor&PositionForTableName=WorkingDay&PrintFilter0=" & SponsorID
                hrf1 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=ViewProformaSponsor&PositionForTableName=WorkingDay&PrintFilter1=" & SponsorID
                cnt = cnt + 1
                response.write vbCrLf & "   <tr>"
                response.write vbCrLf & "      <td>" & cnt & "</td>"
                response.write vbCrLf & "      <td>" & GetComboName("WorkingMonth", dt) & "</td>"
                response.write vbCrLf & "      <td>" & sponsorName & "</td>"
                response.write vbCrLf & "      <td><a href='" & hrf & "' target='_blank'>Prepare</a></td>"
                response.write vbCrLf & "      <td><a href='" & hrf1 & "' target='_blank'>View</a></td>"
                response.write vbCrLf & "   </tr>"
                rst.MoveNext
            Loop
            response.write vbCrLf & "</tbody>"
            response.write vbCrLf & "</table>"
        End If
    rst.Close
    Set rst = Nothing
End Sub

Sub addCSS()
    response.write vbCrLf & "<style>"
    response.write vbCrLf & "   #nav-container {"
    response.write vbCrLf & "       display: flex;"
    response.write vbCrLf & "       justify-content: space-between;"
    response.write vbCrLf & "       align-items: center;"
    response.write vbCrLf & "       gap: 10px;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   select {"
    response.write vbCrLf & "       width: fit-content;"
    response.write vbCrLf & "       max-width: 350px;"
    response.write vbCrLf & "       padding: 5px;"
    response.write vbCrLf & "       border: 1px solid #ccc;"
    response.write vbCrLf & "       border-radius: 5px;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #btn-profoma {"
    response.write vbCrLf & "       padding: 5px 10px;"
    response.write vbCrLf & "       border: 1px solid #ccc;"
    response.write vbCrLf & "       border-radius: 5px;"
    response.write vbCrLf & "       background-color: #f0f0f0;"
    response.write vbCrLf & "       cursor: pointer;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable {"
    response.write vbCrLf & "       width: 100%;"
    response.write vbCrLf & "       max-width: 1200px;"
    response.write vbCrLf & "       margin: 20px auto;"
    response.write vbCrLf & "       font-size: 14px;"
    response.write vbCrLf & "       font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;"
    response.write vbCrLf & "       box-sizing: border-box;"
    response.write vbCrLf & "       border-collapse: collapse;"
    response.write vbCrLf & "       box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);"
    response.write vbCrLf & "       border-radius: 8px;"
    response.write vbCrLf & "       overflow: hidden;"
    response.write vbCrLf & "       background-color: #fff;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable thead {"
    response.write vbCrLf & "       background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);"
    response.write vbCrLf & "       color: white;"
    response.write vbCrLf & "       text-align: center;"
    response.write vbCrLf & "       font-size: 16px;"
    response.write vbCrLf & "       font-weight: 600;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable thead td {"
    response.write vbCrLf & "       padding: 15px 12px;"
    response.write vbCrLf & "       border: none;"
    response.write vbCrLf & "       text-transform: uppercase;"
    response.write vbCrLf & "       letter-spacing: 0.5px;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable tbody tr {"
    response.write vbCrLf & "       transition: all 0.3s ease;"
    response.write vbCrLf & "       border-bottom: 1px solid #e0e0e0;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable tbody tr:hover {"
    response.write vbCrLf & "       background-color: #f8f9ff;"
    response.write vbCrLf & "       transform: translateY(-1px);"
    response.write vbCrLf & "       box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable tbody tr:nth-child(even) {"
    response.write vbCrLf & "       background-color: #fafafa;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable tbody tr:nth-child(even):hover {"
    response.write vbCrLf & "       background-color: #f0f2ff;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable tbody td {"
    response.write vbCrLf & "       padding: 12px;"
    response.write vbCrLf & "       text-align: center;"
    response.write vbCrLf & "       border: none;"
    response.write vbCrLf & "       vertical-align: middle;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable tbody td:first-child {"
    response.write vbCrLf & "       font-weight: 600;"
    response.write vbCrLf & "       color: #667eea;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable tbody td:nth-child(3) {"
    response.write vbCrLf & "       font-weight: 500;"
    response.write vbCrLf & "       color: #333;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable tbody a {"
    response.write vbCrLf & "       display: inline-block;"
    response.write vbCrLf & "       padding: 6px 12px;"
    response.write vbCrLf & "       margin: 2px;"
    response.write vbCrLf & "       text-decoration: none;"
    response.write vbCrLf & "       border-radius: 4px;"
    response.write vbCrLf & "       font-size: 12px;"
    response.write vbCrLf & "       font-weight: 500;"
    response.write vbCrLf & "       transition: all 0.3s ease;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable tbody a[href*='Prepare'] {"
    response.write vbCrLf & "       background-color: #28a745;"
    response.write vbCrLf & "       color: white;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable tbody a[href*='Prepare']:hover {"
    response.write vbCrLf & "       background-color: #218838;"
    response.write vbCrLf & "       transform: translateY(-1px);"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable tbody a[href*='View'] {"
    response.write vbCrLf & "       background-color: #007bff;"
    response.write vbCrLf & "       color: white;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable tbody a[href*='View']:hover {"
    response.write vbCrLf & "       background-color: #0056b3;"
    response.write vbCrLf & "       transform: translateY(-1px);"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   #myTable .last {"
    response.write vbCrLf & "       background-color: #3C8F6D;"
    response.write vbCrLf & "       color: #fff;"
    response.write vbCrLf & "       font-weight: 700;"
    response.write vbCrLf & "       text-align: center;"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "   @media (max-width: 768px) {"
    response.write vbCrLf & "       #myTable {"
    response.write vbCrLf & "           font-size: 12px;"
    response.write vbCrLf & "       }"
    response.write vbCrLf & "       #myTable thead td, #myTable tbody td {"
    response.write vbCrLf & "           padding: 8px 6px;"
    response.write vbCrLf & "       }"
    response.write vbCrLf & "   }"
    response.write vbCrLf & "</style>"
End Sub


'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
