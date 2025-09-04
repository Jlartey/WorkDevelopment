'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim nm
Dim dur
Dim bDt
Dim bDt1
Dim bDt2
Dim gen
Dim pat
Dim spn
Dim sqlSelect
Dim sqlFrom
Dim sqlWhCls
Dim cnt
Dim vDt1
Dim vDt2
Dim rst
Dim pCnt
Dim pBDt
Dim pAge
Dim pGen
Dim pTel
Dim pAdd
Dim pOcc
Dim sql2
Dim lnkID
Dim lnkText
Dim lnkUrl
Dim navPop
Dim inout
Dim fntSize
Dim fntColor
Dim bgColor
Dim tb
Dim tbKy
Dim tbNm
Dim recKy
Dim hasPrt
Dim lnkCnt
Dim insno
Dim phoneNo


lnkCnt = 0
Set rst = CreateObject("ADODB.Recordset")

LoadCSS

nm = Trim(Request.querystring("printfilter0"))
dur = Trim(Request.querystring("printfilter1"))
bDt = Trim(Request.querystring("printfilter2"))
gen = Trim(Request.querystring("printfilter3"))
pat = Trim(Request.querystring("printfilter4"))
spn = Trim(Request.querystring("printfilter5"))
insno = Trim(Request.querystring("printfilter6"))
phoneNo = Trim(Request.querystring("printfilter7"))

bDt1 = ""
bDt2 = ""
sqlSelect = ""
sqlFrom = ""
sqlWhCls = ""
cnt = 0
'If Len(dur) = 0 Then
'  dur = "C860" ' "C990"
'End If
ExtractDates bDt, bDt1, bDt2
'Patient Name

If Len(nm) > 0 Then
    cnt = cnt + 1

    If cnt > 1 Then
        sqlWhCls = sqlWhCls & " and "
    End If

    sqlWhCls = sqlWhCls & " Patient.PatientName like '%" & nm & "%'"
End If

'Last Visit

If Len(dur) > 0 Then
    cnt = cnt + 1

    If cnt > 1 Then
        sqlWhCls = sqlWhCls & " and "
    End If

    vDt2 = CDate(FormatDate(Now()))

    Select Case UCase(dur)
        Case "C990" '1
            vDt1 = vDt2
        Case "C980" '2
            vDt1 = DateAdd("d", -1, vDt2)
        Case "C970" '3
            vDt1 = DateAdd("d", -3, vDt2)
        Case "C960" '7
            vDt1 = DateAdd("d", -7, vDt2)
        Case "C950" '14
            vDt1 = DateAdd("d", -14, vDt2)
        Case "C940" '31
            vDt1 = DateAdd("d", -31, vDt2)
        Case "C930" '62
            vDt1 = DateAdd("d", -62, vDt2)
        Case "C920" '93
            vDt1 = DateAdd("d", -93, vDt2)
        Case "C910" '186
            vDt1 = DateAdd("d", -186, vDt2)
        Case "C900" '279
            vDt1 = DateAdd("d", -279, vDt2)
        Case "C890" '366
            vDt1 = DateAdd("d", -366, vDt2)
        Case "C880" '2y
            vDt1 = DateAdd("d", -732, vDt2)
        Case "C870" '3 y
            vDt1 = DateAdd("d", -1008, vDt2)
        Case Else
            vDt1 = DateAdd("d", -10000, vDt2)
    End Select

    vDt1 = FormatDate(vDt1) & " 00:00:00"
    vDt2 = FormatDate(vDt2) & " 23:59:59"
    sqlFrom = sqlFrom & ",Visitation "
    sqlWhCls = sqlWhCls & " Visitation.VisitDate between '" & vDt1 & "' and '" & vDt2 & "'"
    sqlWhCls = sqlWhCls & " and Visitation.PatientID=Patient.PatientID "
End If

'BirthDate

If IsDate(bDt1) And IsDate(bDt2) Then
    cnt = cnt + 1

    If cnt > 1 Then
        sqlWhCls = sqlWhCls & " and "
    End If

    sqlWhCls = sqlWhCls & " Patient.BirthDate between '" & bDt1 & "' and '" & bDt2 & "'"
End If

'GenderID

If Len(gen) > 0 Then
    cnt = cnt + 1

    If cnt > 1 Then
        sqlWhCls = sqlWhCls & " and "
    End If

    sqlWhCls = sqlWhCls & " Patient.GenderID='" & gen & "'"
End If

'PatientID

If Len(pat) > 0 Then
    cnt = cnt + 1

    If cnt > 1 Then
        sqlWhCls = sqlWhCls & " and "
    End If

    sqlWhCls = sqlWhCls & " Patient.PatientID='" & pat & "'"
End If

'Sponsor

If Len(spn) > 0 Then
    cnt = cnt + 1

    If cnt > 1 Then
        sqlWhCls = sqlWhCls & " and "
    End If

    sqlFrom = sqlFrom & ",InsuredPatient,Sponsor "
    sqlWhCls = sqlWhCls & " InsuredPatient.SponsorID='" & spn & "'"
    sqlWhCls = sqlWhCls & " and InsuredPatient.SponsorID=Sponsor.SponsorID "
    sqlWhCls = sqlWhCls & " and InsuredPatient.PatientID=Patient.PatientID "
End If

If Len(insno) > 0 Then
    cnt = cnt + 1

    If cnt > 1 Then
        sqlWhCls = sqlWhCls & " and "
    End If
    
    If Len(spn) > 0 Then
       ' sqlWhCls = sqlWhCls & " InsuredPatient.InsuranceNo like '%" & insNo & "%'"
        sqlWhCls = sqlWhCls & " InsuredPatient.InsuranceNo = '" & insno & "'"
    Else
        sqlFrom = sqlFrom & ",InsuredPatient "
        
        sqlWhCls = sqlWhCls & " InsuredPatient.PatientID=Patient.PatientID "
        'sqlWhCls = sqlWhCls & " and InsuredPatient.InsuranceNo like '%" & insNo & "%'"
        sqlWhCls = sqlWhCls & " and InsuredPatient.InsuranceNo = '" & insno & "'"
    End If
End If

If Len(phoneNo) > 0 Then
    cnt = cnt + 1

    If cnt > 1 Then
        sqlWhCls = sqlWhCls & " and "
    End If
    sqlWhCls = sqlWhCls & " Patient.ResidencePhone like '%" & phoneNo & "%'"
End If

If cnt > 0 Then
    sql = "select distinct Patient.PatientID from Patient "
    sql = sql & " " & sqlFrom
    sql = sql & " where " & sqlWhCls

    sql2 = "select Patient.PatientID,Patient.PatientName,Patient.GenderID,Patient.BirthDate "
    sql2 = sql2 & " ,Patient.Occupation,Patient.ResidencePhone,Patient.ResidenceAddress from Patient "
    sql2 = sql2 & " where PatientID in (" & sql & ") order by Patient.PatientName"
    pCnt = 0
    tb = "Patient"
    tbKy = "PatientID"
    tbNm = "Patient"

    With rst
        .maxrecords = 100
        .open qryPro.FltQry(sql2), conn, 3, 4

        If .recordCount > 0 Then
            .MoveFirst
            hasPrt = HasPrintOutAccess(jSchd, tb & "RCP")

            If .recordCount = 1 Then
                recKy = .fields(tbKy)
                lnkUrl = "wpg" & tb & ".asp?PageMode=ProcessSelect&" & tbKy & "=" & recKy

                If hasPrt Then
                    lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=" & tb & "RCP&PositionForTableName=" & tb & "&" & tb & "ID=" & recKy
                End If

                response.redirect lnkUrl
            End If

            response.write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
            response.write "<tr class=""cpHdrTr"">"
            response.write "<td><table border=""0"" cellpadding=""3"" cellspacing=""0""><tr>"
            response.write "<td class=""cpHdrTd"">Patient Search Information&nbsp;&nbsp;</td>"

            If HasAccessRight(uName, "frm" & tb, "New") Then
                response.write "<td>"
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
            End If

            response.write "</tr></table></td>"
            response.write "</tr>"
            response.write "<tr>"
            response.write "<td>"
            response.write "<table width=""100%"" border=""1"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse:collapse; font-size:12pt"" >"
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

            inout = "IN"
            fntSize = ""
            fntColor = "#448844"
            bgColor = ""
            wdth = ""

            Do While Not .EOF
                recKy = .fields(tbKy)
                'Clickable Url Link
                lnkUrl = "wpg" & tb & ".asp?PageMode=ProcessSelect&" & tbKy & "=" & recKy
                navPop = "POP"

                If hasPrt Then
                    lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=" & tb & "RCP&PositionForTableName=" & tb & "&" & tb & "ID=" & recKy
                    navPop = "NAV"
                End If

                pCnt = pCnt + 1
                pBDt = ""
                pAge = ""
                pTel = ""
                pAdd = ""
                pOcc = ""

                If IsDate(.fields("BirthDate")) Then
                    pBDt = FormatDate(.fields("BirthDate"))
                    pAge = Round((DateDiff("d", CDate(pBDt), Now()) / 365.25), 0)
                End If

                pGen = .fields("GenderID")

                If Not IsNull(.fields("Occupation")) Then
                    pOcc = Trim(.fields("Occupation"))
                End If

                If Not IsNull(.fields("ResidenceAddress")) Then
                    pAdd = Trim(.fields("ResidenceAddress"))
                End If

                If Not IsNull(.fields("ResidencePhone")) Then
                    pTel = Trim(.fields("ResidencePhone"))
                End If
                
'                If Len(insNo) > 0 Then
'                    pins = Trim(.fields("InsuranceNo"))
'                End If
                
                response.write "<tr>"
                response.write "<td>" & CStr(pCnt) & "</td>"
                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = .fields("PatientID")
                AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
                'Clickable Url Link
                response.write "<td>"
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = .fields("PatientName")
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
                
                
                .MoveNext
            Loop

        End If

        .Close
    End With

    response.write "</table>"
    response.write "</td>"
    response.write "</tr>"
    response.write "</table>"
    Set rst = Nothing
End If

Sub AddUrlLink(lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth)
    Dim plusMinus
    Dim imgName
    Dim lnkOpClNavPop
    Dim align
    plusMinus = ""
    imgName = ""
    align = ""
    lnkOpClNavPop = inout & "||" & navPop & "||800||600||CLOSE"
    AddPrtNavLink lnkID, plusMinus, imgName, lnkText, lnkUrl, lnkOpClNavPop, fntSize, fntColor, bgColor, align, wdth
End Sub

'ExtractDates
Sub ExtractDates(inFlt, outDt1, outDt2)
    Dim arr
    Dim ul
    Dim num
    Dim dat1
    Dim dat2
    dat1 = ""
    dat2 = ""
    arr = Split(inFlt, "||")
    ul = UBound(arr)

    If ul >= 0 Then

        For num = 0 To ul

            If num = 0 Then
                dat1 = Trim(arr(0))
            ElseIf num = 1 Then
                dat2 = Trim(arr(1))
            End If

        Next

        If IsDate(dat1) Then

            If IsDate(dat2) Then
            Else 'No Dat2
                dat2 = FormatDate(CDate(dat1)) & " 23:59:59"
                dat1 = FormatDate(CDate(dat1)) & " 00:00:00"
            End If

        Else 'No Dat1

            If IsDate(dat2) Then
                dat1 = FormatDate(CDate(dat2)) & " 0:00:00"
                dat2 = FormatDate(CDate(dat2)) & " 23:59:59"
            Else 'No Dat2
            End If

        End If

    End If

    outDt1 = dat1
    outDt2 = dat2
End Sub

Sub LoadCSS()
    Dim str
    str = ""
    str = str & "<style type='text/css' id=""styPrt"">"
    str = str & ".cpHdrTd{font-size:14pt;font-weight:bold}"
    str = str & ".cpHdrTr{background-color:#eeeeee}"
    str = str & ".cpHdrTd2{font-size:12pt;font-weight:bold}"
    str = str & ".cpHdrTr2{background-color:#eeeeee}" 'fafafa
    str = str & "</style>"
    response.write str
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

        If .recordCount > 0 Then
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
