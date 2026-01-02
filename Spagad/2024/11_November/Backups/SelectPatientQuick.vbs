'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

AddPageJS
ShowSearchPage
ShowSearchResult
Sub ShowSearchResult()
    Dim sql, rst, html, insno, pat, phNo, patFld, tmp, tmp2
    Dim pBDt, pAge, pAdd, pTel, pOcc, hasPrt
    
    pat = Trim(request("PrintFilter0"))
    insno = Trim(request("PrintFilter1"))
    phNo = Trim(request("PrintFilter2"))
    
    sql = "select top 100 Patient.PatientID, Patient.PatientName, Patient.GenderID, Patient.BirthDate, Patient.ResidencePhone, Patient.ResidenceAddress"
    sql = sql & " , Patient.Occupation "
    sql = sql & " from Patient where 1=0 "
    
    If Len(pat) > 0 Then
        tmp2 = ""
        For Each tmp In Split(pat, " ")
            If tmp2 <> "" Then
                tmp2 = tmp2 & " and "
            End If
            tmp2 = tmp2 & " (PatientID like '%" & Trim(tmp) & "%' "
            tmp2 = tmp2 & " or PatientName like '%" & Trim(tmp) & "%' "
            tmp2 = tmp2 & " or ServiceNo like '%" & Trim(tmp) & "%')"
        Next
        If tmp2 <> "" Then
            sql = sql & " or (" & tmp2 & ") "
        End If
    End If
    If Len(insno) > 0 Then
        If Len(pat) > 0 Then
            sql = sql & " and "
        Else
            sql = sql & " or "
        End If
        sql = sql & " Patient.PatientID in (select PatientID from InsuredPatient where InsuredPatient.InsuranceNo='" & insno & "')"
    End If
    If Len(phNo) > 0 Then
        If Len(pat) > 0 Or Len(insno) > 0 Then
            sql = sql & " and "
        Else
            sql = sql & " or "
        End If
        sql = sql & " Patient.ResidencePhone like '%" & phNo & "%' "
    End If
    
    tbNm = "Patient"
    tb = "Patient"
    tbKy = "PatientID"
    pCnt = 0
    hasPrt = HasPrintOutAccess(jSchd, tb & "RCP")
    
    response.write "<table width=""100%"" border=""1"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse:collapse; font-size:12pt"" >"
    If HasAccessRight(uName, "frm" & tb, "New") Then
        response.write "<tr class=""cpHdrTr2"">"
            response.write "<td colspan='100' style='text-align:left;'>"
            'Clickable Url Link
            lnkCnt = lnkCnt + 1
            lnkID = "lnk" & CStr(lnkCnt)
            lnkText = "<b>&nbsp;&nbsp;Add New " & tbNm & "</b>"
            lnkUrl = "wpg" & tb & ".asp?PageMode=AddNew"
            navPop = "POP"
            inout = "IN"
            fntSize = ""
            fntColor = "#444488"
            bgColor = ""
            wdth = ""
            AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            response.write "</td>"
        response.write "</tr>"
    End If

    response.write "<tr class=""cpHdrTr2"">"
    response.write "<td><b>No.</b></td>"
    response.write "<td align=""center""><b>Patient No.</b></td>"
    response.write "<td align=""center""><b>Patient Name</b></td>"
    response.write "<td align=""center""><b>Gender</b></td>"
    response.write "<td align=""center""><b>Birth Date</b></td>"
    response.write "<td align=""center""><b>Age</b></td>"
    response.write "<td align=""center""><b>Phone</b></td>"
    response.write "<td align=""center""><b>Address</b></td>"
    response.write "<td align=""center""><b>Occupation</b></td>"
    response.write "</tr>"

    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
        
            recKy = rst.fields("PatientID")
            'Clickable Url Link
            lnkUrl = "wpg" & tb & ".asp?PageMode=ProcessSelect&" & tbKy & "=" & server.URLEncode(recKy)
            navPop = "POP"

            If hasPrt Then
                lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=" & tb & "RCP&PositionForTableName=" & tb & "&" & tb & "ID=" & server.URLEncode(recKy)
                navPop = "NAV"
            End If

            pCnt = pCnt + 1
            pBDt = ""
            pAge = ""
            pTel = ""
            pAdd = ""
            pOcc = ""
            
            If IsDate(rst.fields("BirthDate")) Then
                pBDt = FormatDate(rst.fields("BirthDate"))
                pAge = Round((DateDiff("d", CDate(pBDt), Now()) / 365.25), 0)
            End If
            pTel = rst.fields("ResidencePhone")
            pOcc = rst.fields("Occupation")
            pAdd = rst.fields("ResidenceAddress")
            pGen = rst.fields("GenderID")
            
            response.write "<tr>"
                response.write "<td>" & CStr(pCnt) & "</td>"
                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = rst.fields("PatientID")
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = rst.fields("PatientName")
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = GetComboName("Gender", pGen)
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = pBDt
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = pAge
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = pTel
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = pAdd
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = pOcc
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
            response.write "</tr>"
            rst.MoveNext
        Loop
        rst.Close
    End If
    response.write "</table></td>"
End Sub
Sub ShowSearchPage()
    SetPageVariable "AutoHidePrintControl", "Yes"
    
    response.write " <style id=""styMain"">"
    response.write "     .inpEleFnt {"
    response.write "         font-size: 12pt"
    response.write "     }"
    response.write " "
    response.write "     .pnlOn {"
    response.write "         font-family: Verdana;"
    response.write "         font-size: 10pt;"
    response.write "         color: #FF6699;"
    response.write "         background-color: #bbbbff;"
    response.write "         cursor: hand"
    response.write "     }"
    response.write " "
    response.write "     .pnlOff {"
    response.write "         font-family: Verdana;"
    response.write "         font-size: 10pt;"
    response.write "         color: #6699FF;"
    response.write "         background-color: #f2f2f2"
    response.write "     }"
    response.write " "
    response.write "     .pnlHead {"
    response.write "         font-family: Verdana;"
    response.write "         font-size: 11pt;"
    response.write "         color: #0a0a0a;"
    response.write "         background-color: #e2e2e2"
    response.write "     }"
    response.write " "
    response.write "     .imgPnl {"
    response.write "         width: 18;"
    response.write "         height: 18"
    response.write "     }"
    response.write " "
    response.write "     .ctxPnlOn {"
    response.write "         font-family: verdana;"
    response.write "         font-size: 8pt;"
    response.write "         color: #FF6699;"
    response.write "         background-color: #bbbbff;"
    response.write "         cursor: hand"
    response.write "     }"
    response.write " "
    response.write "     .ctxPnlOff {"
    response.write "         font-family: verdana;"
    response.write "         font-size: 8pt;"
    response.write "         color: #6699FF"
    response.write "     }"
    response.write " </style>"
    response.write " <form method=""POST"" action=""wpgPrtPrintInputFilter.asp?PositionForTableName=WorkingDay&amp;"" id=""form1"""
    response.write "     name=""form1"">"
    response.write "     <center>"
    response.write "         <table width=""100%"" border=""0"" cellspacing=""2"" cellpadding=""0"">"
    response.write "             <tbody>"
    response.write "                 <tr>"
    response.write "                     <td align=""center"">"
    response.write "                         <table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
    response.write "                             <tbody>"
    response.write "                                 <tr>"
    response.write "                                     <td width=""70"" bgcolor=""#fffffe"">"
    response.write "                                         <p align=""center"">&nbsp;</p>"
    response.write "                                     </td>"
    response.write "                                     <td width=""60"" bgcolor=""#fffffe"">"
    response.write "                                         <p align=""center"">&nbsp;</p>"
    response.write "                                     </td>"
    response.write "                                     <td align=""center"" height=""30"" bgcolor=""#fffffe"""
    response.write "                                         style=""font-size: 13pt; color: #606060;  font-family:Arial"">"
    response.write "                                         <table width=""100%"" border=""0"" style=""border-collapse: collapse"""
    response.write "                                             bordercolor=""#111111"" cellpadding=""0"" cellspacing=""0"">"
    response.write "                                             <tbody>"
    response.write "                                                 <tr>"
    response.write "                                                     <td align=""right"">"
    response.write "                                                         <table border=""0"" cellpadding=""0"" cellspacing=""0"""
    response.write "                                                             style=""border-collapse: collapse"" bordercolor=""#111111"">"
    response.write "                                                             <tbody>"
    response.write "                                                                 <tr>"
    response.write "                                                                     <td width=""0"" height=""0"" bgcolor=""sandybrown"">"
    response.write "                                                                     </td>"
    response.write "                                                                     <td width=""0"" height=""0"" bgcolor=""lime""></td>"
    response.write "                                                                 </tr>"
    response.write "                                                                 <tr>"
    response.write "                                                                     <td width=""0"" height=""0"" bgcolor=""gold""></td>"
    response.write "                                                                     <td width=""0"" height=""0"" bgcolor=""salmon""></td>"
    response.write "                                                                 </tr>"
    response.write "                                                             </tbody>"
    response.write "                                                         </table>"
    response.write "                                                     </td>"
    response.write "                                                     <td style=""font-size: 11pt; font-weight:bold;  color: #606060; font-family: verdana"""
    response.write "                                                         align=""center"">Select Patient</td>"
    response.write "                                                 </tr>"
    response.write "                                             </tbody>"
    response.write "                                         </table>"
    response.write "                                     </td>"
    response.write "                                     <td width=""70"" bgcolor=""#fffffe"">"
    response.write "                                         <p align=""center"">&nbsp;</p>"
    response.write "                                     </td>"
    response.write "                                     <td width=""60"" bgcolor=""#fffffe"">"
    response.write "                                         <p align=""center"">&nbsp;</p>"
    response.write "                                     </td>"
    response.write "                                 </tr>"
    response.write "                             </tbody>"
    response.write "                         </table>"
    response.write "                     </td>"
    response.write "                 </tr>"
    response.write "                 <tr>"
    response.write "                     <td align=""center"">"
    response.write "                         <table id=""tblPrintLayout"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"""
    response.write "                             bgcolor=""White"" style=""border-collapse: collapse"" bordercolor=""#111111"">"
    response.write "                             <tbody>"
    response.write "                                 <tr>"
    response.write "                                     <td>"
    response.write "                                         <table width=""100%"" height=""100%"" border=""1"" cellpadding=""5"" cellspacing=""0"""
    response.write "                                             style=""border-collapse: collapse"" bordercolor=""#ffffff"">"
    response.write "                                             <tbody>"
    response.write "                                                 <tr class=""pnlOff"">"
    response.write "                                                     <td name=""tdLabelInpFltInputText||0"""
    response.write "                                                         id=""tdLabelInpFltInputText||0"" align=""right"">Patient Name/Folder No."
    response.write "                                                     </td>"
    response.write "                                                     <td name=""tdInputInpFltInputText||0"""
    response.write "                                                         id=""tdInputInpFltInputText||0"""
    response.write "                                                         style=""font-size: 9pt; font-family: Arial""><input"
    response.write "                                                             type=""text"" name=""PrintFilter0"""
    response.write "                                                             id=""PrintFilter0"" size=""50"" value=""" & request("PrintFilter0") & """></td>"
    response.write "                                                     <td name=""tdLabelInpFltInputText||6"""
    response.write "                                                         id=""tdLabelInpFltInputText||6"" align=""right"">Insurance /"
    response.write "                                                         staff no.</td>"
    response.write "                                                     <td name=""tdInputInpFltInputText||6"""
    response.write "                                                         id=""tdInputInpFltInputText||6"""
    response.write "                                                         style=""font-size: 9pt; font-family: Arial""><input"
    response.write "                                                             type=""text"" name=""PrintFilter1"""
    response.write "                                                             id=""PrintFilter1"" size=""20"" value=""" & request("PrintFilter1") & """></td>"
    response.write "                                                     <td name=""tdLabelInpFltInputText||7"""
    response.write "                                                         id=""tdLabelInpFltInputText||7"" align=""right"">Phone No.</td>"
    response.write "                                                     <td name=""tdInputInpFltInputText||7"""
    response.write "                                                         id=""tdInputInpFltInputText||7"""
    response.write "                                                         style=""font-size: 9pt; font-family: Arial""><input"
    response.write "                                                             type=""text"" name=""PrintFilter2"""
    response.write "                                                             id=""PrintFilter2"" size=""30"" value=""" & request("PrintFilter2") & """></td>"
    response.write "                                                 </tr>"
    response.write "                                                 <tr> "
    
    response.write "                                                 </tr> "
    response.write "                                                 <tr class=""pnlOff"">"
    response.write "                                                     <td align=""center"" colspan=""6""><input type=""submit"""
    response.write "                                                             value=""Search"""
    response.write "                                                             onclick=""search(this)"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    response.write "                                                     </td>"
    response.write "                                                 </tr>"
    response.write "                                             </tbody>"
    response.write "                                         </table>"
    response.write "                                     </td>"
    response.write "                                 </tr>"
    response.write "                             </tbody>"
    response.write "                         </table>"
    response.write "                     </td>"
    response.write "                 </tr>"
    response.write "             </tbody>"
    response.write "         </table>"
    response.write "     </center>"
    response.write " </form>"
End Sub
Sub AddPageJS()
    Dim html
    
    html = "<script>"
    html = html & vbCrLf & " function search(inp){"
    html = html & vbCrLf & "    inp.disabled = true;"
    html = html & vbCrLf & "    form1.submit();"
    html = html & vbCrLf & " }"
    html = html & vbCrLf & " "
    html = html & vbCrLf & " form1.autocomplete='off';"
    html = html & vbCrLf & " window.resizeTo(1000, 800);"
    html = html & "</script>"
    response.write html
End Sub
Function HasPrintOutAccess(jb, prt)
    Dim rstTblSql
    Dim sql
    Dim ot
    ot = False
    Set rstTblSql = CreateObject("ADODB.Recordset")

    With rstTblSql
        sql = "select JobScheduleID from printoutalloc "
        sql = sql & " where printlayoutid='" & prt & "' and jobscheduleid='" & jb & "'"
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .MoveFirst
            ot = True
        End If

        .Close
    End With

    HasPrintOutAccess = ot
    Set rstTblSql = Nothing
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
