
Sub Glob_AddReportHeader()
    Dim str
    str = ""
    str = str & "<table border=""0""cellspacing=""0"" cellpadding=""0"" width=""" & PrintWidth & """ style=""border-collapse:collapse; page-break-after:always"">"
    str = str & "<tr><td>"
    str = str & "<img src=""images/banner1.bmp"" height=""60"" width=""150"">"
    str = str & "</td>"
    str = str & "<td>"

    str = str & "<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"" style=""font-family:'Times New Roman' !important"">"
    str = str & "<tr>"
    str = str & "<td align=""center"" style=""font-size: 14pt; font-weight:bold"" colspan=""6"">MEDIFEM MULTI-SPECIALIST AND FERTILITY HOSPITAL</td>"
    str = str & "</tr>"
    str = str & "<tr>"
    str = str & "<td align=""center"" style=""font-size: 14pt"" colspan=""6"">" & UCase(GetComboName("Branch", brnch)) & "</td>"
    str = str & "</tr>"
    ''Address
    str = str & "<tr>"
    str = str & "<td align=""center"" style=""font-size: 10pt"" colspan=""6"">" & Trim(GetComboNameFld("Branch", brnch, "Address"))
    str = str & "</td>"
    str = str & "</tr>"
    ' ''Tel
    ' str = str & "<tr>"
    ' str = str & "<td align=""center"" style=""font-size: 10pt"" colspan=""6""> TEL&nbsp;:&nbsp;"
    ' branchTel = Trim(GetComboNameFld("Branch", brnch, "OfficePhone"))
    ' If branchTel <> "/" Then
    '     str = str & branchTel
    ' End If
    ' str = str & "</td>"
    ' str = str & "</tr>"
    ' ' str = str & "<tr>"
    ' ' str = str & "<td align=""center"" style=""font-size: 10pt"" colspan=""6""><nobr>TEL:&nbsp;+233 0302 682832-4, 689175&nbsp;&nbsp;&nbsp;&nbsp;FAX:&nbsp;(0302) 683298&nbsp;&nbsp;&nbsp;&nbsp;EMAIL:&nbsp;sicinfo@sic-gh.com&nbsp;&nbsp;&nbsp;&nbsp;WEB:&nbsp;www.sic-gh.com</nobr></td>"
    ' ' str = str & "</tr>"

    str = str & "</table>"
    str = str & "</td>"
    str = str & "</tr>"

    str = str & "</table>"
    response.write str
End Sub



Sub Glob_AddUrlLink(lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth)
    Dim plusMinus, imgName, lnkOpClNavPop, align, lnkText2
    plusMinus = ""
    imgName = ""
    align = ""
    lnkText2 = "<div class='btn_ mouseover'>" & lnkText & "</div>"
    lnkOpClNavPop = inout & "||" & navPop & "||800||600||CLOSE"
    AddPrtNavLink lnkID, plusMinus, imgName, lnkText2, lnkUrl, lnkOpClNavPop, fntSize, fntColor, bgColor, align, wdth
End Sub

Sub Glob_AddUrlLinkSM(lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth)
    Dim plusMinus, imgName, lnkOpClNavPop, align, lnkText2
    plusMinus = ""
    imgName = ""
    align = ""
    lnkText2 = "<div class='btn_ mouseover btn_sm'>" & lnkText & "</div>"
    lnkOpClNavPop = inout & "||" & navPop & "||800||600||CLOSE"
    AddPrtNavLink lnkID, plusMinus, imgName, lnkText2, lnkUrl, lnkOpClNavPop, fntSize, fntColor, bgColor, align, wdth
End Sub


Sub Glob_ExecJS(script)
    response.write "<script type=""text/javascript"">" & script & "</script>"
End Sub

Sub Glob_SetDocumentTitle(title)
    Dim title2
    title2 = title
    title2 = Replace(title2, "<br>", " | ")
    title2 = Replace(title2, "<b>", "")
    title2 = Replace(title2, "</b>", "")
    response.write "<script type=""text/javascript"">this.document.title='" & title2 & "';var elDocPage = this.parent.document; if (elDocPage){elDocPage.title='" & title2 & "';} </script>"
End Sub

Sub Glob_CursorLoading2()
  response.write "<style>table{cursor:progress !important;}</style>"
End Sub

Sub Glob_CursorLoading()
  response.write "<style>table{cursor:wait !important;}</style>"
End Sub

Sub Glob_CursorReady()
  response.write "<style>table{cursor:auto !important;}</style>"
End Sub

Sub Glob_ShowDashboardToastNotification()
    response.write "<script>"
    response.write "var script = document.createElement('script');" & vbNewLine
    response.write "script.src = 'Scripts/toastify-js/toastify.js';" & vbNewLine
    response.write "document.getElementsByTagName('head')[0].appendChild(script);" & vbNewLine
    response.write " " & vbNewLine
    response.write "var script2 = document.createElement('script');" & vbNewLine
    response.write "script2.src = 'Scripts/notificationbroadcast.js';" & vbNewLine
    response.write "document.getElementsByTagName('head')[0].appendChild(script2);" & vbNewLine
    response.write " " & vbNewLine
    response.write "var link = document.createElement('link');" & vbNewLine
    response.write "link.href = 'Scripts/toastify-js/toastify.css';" & vbNewLine
    response.write "link.rel = 'stylesheet';" & vbNewLine
    response.write "link.type = 'text/css';" & vbNewLine
    response.write "document.getElementsByTagName('head')[0].appendChild(link); " & vbNewLine
    response.write " " & vbNewLine
    response.write "</script>"
End Sub

Sub Glob_cCreateSystemAdmin(usr)
    sql = "Insert into SystemAdmin (SystemAdminID,SystemAdminName,UserPassword,UserStatusID,ConfirmPassword,KeyPrefix,MyReportGroupID,NavigationModeID,NavigationThemeID, DefaultUrl) "
    sql = sql & " values ('" & usr & "','" & usr & "','" & usr & "','UST001','" & usr & "','',  'M001', 'N003', 'N008', '-') "
    conn.execute qryPro.FltQry(sql)

    sql = "Insert into SysDeveloperType (SysDeveloperTypeID,SysDeveloperTypeName,Description,KeyPrefix) "
    sql = sql & " values ('" & usr & "','" & usr & "','','') "
    conn.execute qryPro.FltQry(sql)

    sql = "Insert into SysDeveloper (SysDeveloperID,SysDeveloperName,SysDeveloperTypeID,StaffID,UserStatusID,BranchID,DepartmentID,UnitID,KeyPrefix) "
    sql = sql & " values ('" & usr & "','" & usr & "','" & usr & "','STF001', 'UST001', 'B001', 'DPT001', 'UNT001', NULL) "
    conn.execute qryPro.FltQry(sql)
End Sub


