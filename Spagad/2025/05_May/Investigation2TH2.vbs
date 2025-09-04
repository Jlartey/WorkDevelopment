'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

ShowPage
Sub ShowPage()
    InitPrintLayout
    Dim LabRequestID, labTestID, targTbl, labbyDocID


    LabRequestID = Trim(Request.queryString("LabRequestID"))

    If Request.queryString("LabTestID").count > 0 Then
        labTestID = Trim(Request.queryString("LabTestID"))
       
    End If

    If Request.queryString("LabByDoctorID").count > 0 Then
        labbyDocID = Trim(Request.queryString("LabByDoctorID"))
       
    End If

    ShowLabReport LabRequestID, labTestID, labbyDocID
End Sub
Sub InitPrintLayout()
    Dim htStr
    LoadCSS
    AddPrintJS
    htStr = ""
    htStr = htStr & "<script id=""scptPrintLayoutExtraScript"" LANGUAGE=""javascript"">"
    htStr = htStr & vbCrLf
    htStr = htStr & "function PLExtraScriptOnLoad(){" & vbCrLf
    htStr = htStr & "window.scrollTo(0,0);" & vbCrLf
    htStr = htStr & "}" & vbCrLf
    htStr = htStr & "</script>"
    response.write htStr
   'SetPageVariable "AutoHidePrintControl", "1"
End Sub
Sub AddPrintJS()
    Dim str

    str = str & vbCrLf & "<script>"
    str = str & vbCrLf & "   window.onbeforeprint = addPageNumbers; "
    str = str & vbCrLf & "    "
    str = str & vbCrLf & "   function addPageNumbers() { "
    str = str & vbCrLf & "     if(addPageNumbers.pageNumberDivs){ "
    str = str & vbCrLf & "       for(x of addPageNumbers.pageNumberDivs){ "
    str = str & vbCrLf & "          x.parentNode.removeChild(x); "
    str = str & vbCrLf & "       } "
    str = str & vbCrLf & "     } "
    str = str & vbCrLf & "     addPageNumbers.pageNumberDivs = []; "
    str = str & vbCrLf & "     var totalPages = Math.ceil(document.body.scrollHeight / 1123);  //842px A4 pageheight for 72dpi, 1123px A4 pageheight for 96dpi, "
    str = str & vbCrLf & "     for (var i = 1; i <= totalPages; i++) { "
    str = str & vbCrLf & "       var pageNumberDiv = document.createElement(""div"");"
    str = str & vbCrLf & "       var pageNumber = document.createTextNode(""Page "" + i + "" of "" + totalPages);"
    str = str & vbCrLf & "       pageNumberDiv.style.position = ""absolute"";"
    str = str & vbCrLf & "       pageNumberDiv.style.top = ""calc(("" + i + "" * (297mm - 0.5px)) - 40px)""; //297mm A4 pageheight; 0,5px unknown needed necessary correction value; additional wanted 40px margin from bottom(own element height included)"
    str = str & vbCrLf & "       pageNumberDiv.style.height = ""15px"";"
    str = str & vbCrLf & "       pageNumberDiv.style.font = ""9px Arial"";"
    str = str & vbCrLf & "       pageNumberDiv.appendChild(pageNumber);"
    'str = str & vbCrLf & "       document.body.insertBefore(pageNumberDiv, document.getElementById(""paged-content""));"
    str = str & vbCrLf & "       let pgCn = document.getElementById(""paged-content"");"
    str = str & vbCrLf & "       if(pgCn){"
    str = str & vbCrLf & "          pgCn.insertAdjacentElement('beforebegin', pageNumberDiv);"
    str = str & vbCrLf & "       }"
    str = str & vbCrLf & "       pageNumberDiv.style.left = ""calc(100% - ("" + pageNumberDiv.offsetWidth + ""px + 10px))"";"
    str = str & vbCrLf & "       addPageNumbers.pageNumberDivs.push(pageNumberDiv);"
    str = str & vbCrLf & "     }"
    str = str & vbCrLf & "   }"
    str = str & vbCrLf & "</script>"

    response.write str
End Sub
Sub LoadCSS()
    Dim str
    str = ""
    str = str & "<style type='text/css' id=""styPrt"">"
    str = str & "   .cpHdrTd{font-size:14pt;font-weight:bold}"
    str = str & "   .cpHdrTr{background-color:#ddffdd}"
    str = str & "   .cpHdrTd2{font-size:12pt;font-weight:bold}"
    str = str & "   .cpHdrTr2{background-color:#ddffdd}" 'fafafa
    str = str & "   .report-header{margin:auto;width:177mm!important;font-family:Arial;/*filter:grayscale(1);*/ width:177mm;/*border-collapse:collapse;*/}"
    str = str & "   .report-table {margin:auto;margin-top:10px;font-family:Arial;filter:grayscale(1);width:177mm;/*border-collapse:collapse;*/}"
    str = str & "   .report-table thead tr th{margin:auto; background-color:#d0d0d0;/*rgb(240 230 140)*/;font-size:12px;border-top:1.5px solid #999999;border-bottom:1.5px solid #999999;}"
    str = str & "   .report-header th{margin:auto; background-color:rgb(240 230 140);font-size:12px;border:1px solid silver;}"
    str = str & "   .report-header tbody tr td{font-size:11px;}"
    str = str & "   .report-table tbody tr td{/*background-color:rgb(255 222 173);*/font-size:11px;}"
    str = str & "   .report-header tbody tr td{/*background-color:rgb(255 222 173);*/font-size:11px;}"
    str = str & "   .report-table>tbody>tr:nth-child(2n)>td{/*background-color:rgb(250 240 230);*/}"
    str = str & "   .report-table thead tr th{padding:3px 10px;}"
    str = str & "   .report-table tbody tr td{padding:3px 10px;}"
    str = str & "   .report-header td{padding:0px 10px;}"
    str = str & "   @page{size:a4;margin:0px;}"
    str = str & "   body{margin:10px;}"
    str = str & "  * {-webkit-print-color-adjust: exact;color-adjust: exact !important;}"
    str = str & "  body{counter-reset: inv-name-counter;}"
    str = str & "  .inv-name:before{counter-increment: inv-name-counter; content: counter(inv-name-counter) "". "";}"
    str = str & "  .e-sign{text-align:right;font-weight:bold;font-size:11px;margin-top:10px;font-family:Arial;border-top:1px solid silver;margin:10px auto;width:177mm;}"
    str = str & "     @media print {"
    str = str & "       #trPrintControl {"
    str = str & "         display:none"
    str = str & "       }"
    str = str & "     }"
    
    str = str & "</style>"

    response.write str
End Sub
Sub ShowLabReport(LabRequestID, labTestID, labbyDocID)
    Dim sql, rst, htmlStr, lastLabTest, ltName, requestedTests, lstLbTech, lstDate, dt


    sql = "select * from ("
    sql = sql & " select LabResults.LabRequestID,LabResults.LabTestID,LabResults.TestCompTabID,LabResults.TestComponentID"
    sql = sql & " ,LabRequest.RequestDate, Investigation.LabTechID, Investigation.RequestStatusID"
    sql = sql & " ,LabResults.CompPos"
    sql = sql & " ,LabResults.Column1,LabResults.Column2,LabResults.Column3"
    sql = sql & " ,LabResults.Column4,LabResults.Column5,LabResults.Column6"
    sql = sql & " ,  LabResults.LabRequestID + '||' + LabResults.LabTestID as targTbl"
    sql = sql & " , RequestDate1"
    sql = sql & " from LabResults "
    sql = sql & " inner join LabRequest on LabRequest.LabRequestID=LabResults.LabRequestID"
    sql = sql & " inner join Investigation on Investigation.LabRequestID=LabRequest.LabRequestID"
    sql = sql & " and LabResults.LabTestID=Investigation.LabTestID"
    sql = sql & " and LabResults.LabRequestID=Investigation.LabRequestID and Investigation.RequestStatusID='RRD002' "
    'If labTestID <> "ALL" Then
      '  sql = sql & " and labResults.LabTestID in ('" & Replace(labTestID, ", ", "','") & "')"
    'End If
    sql = sql & " and LabResults.LabRequestID='" & LabRequestID & "'"

    sql = sql & " union all "
    sql = sql & " select LabResults.LabRequestID,LabResults.LabTestID,LabResults.TestCompTabID,LabResults.TestComponentID"
    sql = sql & " ,LabRequest.RequestDate, Investigation2.LabTechID, Investigation2.RequestStatusID"
    sql = sql & " ,LabResults.CompPos"
    sql = sql & " ,LabResults.Column1,LabResults.Column2,LabResults.Column3"
    sql = sql & " ,LabResults.Column4,LabResults.Column5,LabResults.Column6"
    sql = sql & " ,  LabResults.LabRequestID + '||' + Investigation2.LabByDoctorID as targTbl"
    sql = sql & " , RequestDate1"
    sql = sql & " from LabResults "
    sql = sql & " inner join LabRequest on LabRequest.LabRequestID=LabResults.LabRequestID"
    sql = sql & " inner join Investigation2 on Investigation2.LabRequestID=LabRequest.LabRequestID"
    sql = sql & " and LabResults.LabTestID=Investigation2.LabTestID"
    sql = sql & " and LabResults.LabRequestID=Investigation2.LabRequestID and Investigation2.RequestStatusID='RRD002'"
    'If labbyDocID <> "ALL" Then
      '  sql = sql & " and Investigation2.LabByDoctorID in ('" & Replace(labbyDocID, ", ", "','") & "')"
    'End If
    sql = sql & " and LabResults.LabRequestID='" & LabRequestID & "'"
    sql = sql & " ) as LabResults where 1=1"
    sql = sql & " order by TargTbl asc, RequestDate, LabRequestID, LabTestID, TestCompTabID, len(CompPos) asc, CompPos asc"

    Set rst = CreateObject("ADODB.RecordSet")

    htmlStr = "<div id=""paged-content"" style=""width:210mm;margin:0 auto;"">"
    'DisplayHeadLab labRequestID
    htmlStr = htmlStr & DisplayHead(LabRequestID)
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        lastLabTest = ""
        rst.MoveFirst
        Do While Not rst.EOF
            If lastLabTest <> (rst.fields("targTbl")) Then
                If lastLabTest <> "" Then
                    htmlStr = htmlStr & "   </tbody>"
                    htmlStr = htmlStr & "</table>"
                    response.write htmlStr
                    response.flush
                    htmlStr = ""
                End If
                ltName = GetComboName("LabTest", rst.fields("LabTestID"))
                If requestedTests <> "" Then
                    requestedTests = requestedTests & ", "
                End If
                requestedTests = requestedTests & ltName
                ot2 = GetComboName("Labtech", rst.fields("LabTechID"))
                htmlStr = htmlStr & "<table class='report-table'>"
                htmlStr = htmlStr & "   <thead>"
                htmlStr = htmlStr & "       <tr>"
                htmlStr = htmlStr & "           <th class='inv-name' colspan=""100"" style=""text-align:left;"">" & ltName & " -> [" & GetComboName("TestCategory", GetComboNameFld("LabTest", rst.fields("LabTestID"), "TestCategoryID")) & "] Validated By : " & ot2 & "</th>"
                htmlStr = htmlStr & "       </tr>"
                htmlStr = htmlStr & "   </thead>"
                htmlStr = htmlStr & "   <tbody>"
            End If

            htmlStr = htmlStr & GetComponentRow(rst)
            lastLabTest = rst.fields("targTbl")

            If rst.fields("RequestDate1") > lstDate Then
                lstDate = rst.fields("RequestDate1")
                lstLbTech = rst.fields("LabTechID")
            End If

            rst.MoveNext
        Loop
        htmlStr = htmlStr & "   </tbody>"
        htmlStr = htmlStr & "</table>"
        htmlStr = htmlStr & "<div class='e-sign'>Electronically Signed By: " & GetComboName("LabTech", lstLbTech) & "</div>"
    End If
    Set rst = Nothing
    htmlStr = htmlStr & "</div>"

    htmlStr = htmlStr & "<script>let lt= document.getElementById('requestedTests');if(lt){lt.innerHTML = '" & requestedTests & "';}</script>"
    response.write htmlStr
    response.flush
    htmlStr = ""
End Sub
Function GetComponentRow(rst)
    Dim ot, sql, rstCompFld
    Dim arr, arrAll, ulAll, ulTd, arrTd, num, numTd
    Dim vars

    Set rstCompFld = CreateObject("ADODB.RecordSet")
    Set vars = CreateObject("Scripting.Dictionary")
    vars.CompareMode = 1

    ot = ot & "<tr>"
    ot = ot & "   <tr>"
    ot = ot & "   <td valign=""top"">" & GetComboName("TestComponent", rst.fields("TestComponentID")) & "</td>"

    sql = "select * from CompField where CompTableKeyID='TestComponentID' and RecordKey='" & rst.fields("TestComponentID") & "' order by CompFieldID"
    rstCompFld.open qryPro.FltQry(sql), conn, 3, 4
    If rstCompFld.RecordCount > 0 Then
        Do While Not rstCompFld.EOF
            vars.RemoveAll
            vars("testCmp") = rst.fields("TestComponentID").value
            vars("compKey") = "TestComponentID"

            vars("fd") = rstCompFld.fields("compfieldid").value
            vars("tdAttr") = ""
            If Not IsNull(rstCompFld.fields("SubTableFieldSource")) Then
                vars("tdAttr") = Trim(rstCompFld.fields("SubTableFieldSource"))
            End If

            vars("ul") = -1
            arrAll = Split(vars("tdAttr"), "%%")
            ulAll = UBound(arrAll)
            If ulAll = 0 Then
                arr = Split(arrAll(0), "**")
                vars("ul") = UBound(arr)
            ElseIf ulAll = 1 Then
                arr = Split(arrAll(1), "**")
                vars("ul") = UBound(arr)
                'TD
                arrTd = Split(arrAll(0), "**")
                ulTd = UBound(arrTd)
                vars("sTd0") = ""
                vars("sTd1") = ""
                vars("sTd2") = ""
                vars("sTd3") = ""
                If ulTd >= 0 Then
                    For numTd = 0 To ulTd
                        Select Case numTd
                            Case 0
                                vars("sTd0") = arrTd(0)
                            Case 1
                                vars("sTd1") = arrTd(1)
                            Case 2
                                vars("sTd2") = arrTd(2)
                            Case 3
                                vars("sTd3") = arrTd(3)
                        End Select
                    Next
                    Select Case UCase(Trim(vars("sTd0")))
                        Case "COLSPAN"
                            If IsNumeric(vars("sTd1")) Then
                                vars("tdColSp") = vars("sTd1")
                            End If
                        Case "COLSPANALIGN"
                            If IsNumeric(vars("sTd1")) Then
                                vars("tdColSp") = vars("sTd1")
                            End If
                            If Len(vars("sTd2")) > 0 Then
                                vars("tdAlign") = "align=""" & vars("sTd2") & """"
                            End If
                        Case "ALIGN"
                            If Len(vars("sTd1")) > 0 Then
                                vars("tdAlign") = "align=""" & vars("sTd1") & """"
                            End If
                    End Select
                End If
            End If
            vars("src0") = ""
            vars("src1") = ""
            vars("src2") = ""
            vars("src3") = ""
            vars("src4") = ""
            vars("src5") = ""
            If vars("ul") >= 0 Then
                For num = 0 To vars("ul")
                    Select Case num
                        Case 0
                            vars("src0") = arr(0)
                        Case 1
                            vars("src1") = arr(1)
                        Case 2
                            vars("src2") = arr(2)
                        Case 3
                            vars("src3") = arr(3)
                        Case 4
                            vars("src4") = arr(4)
                        Case 5
                            vars("src5") = arr(5)
                    End Select
                Next
            End If

            If Not IsNull(rst.fields(vars("fd"))) Then
                vars("fdVl") = Trim(rst.fields(vars("fd")))
            End If
            If UCase(Trim(vars("src0"))) = "USERWRITEPAD" Then
                vars("tdDim") = ""
                If IsNumeric(vars("src2")) Then
                    vars("tdDim") = vars("tdDim") & " width=""" & CStr(CInt(vars("src2")) + 30) & """ "
                End If
                If IsNumeric(vars("src3")) Then
                    vars("tdDim") = vars("tdDim") & " height=""" & CStr(CInt(vars("src3")) + 140) & """ "
                End If
                iFUrl = "wpgWritingPadViewer.asp?PositionforTableName=LabTest&LabTestID=" & testDat & "&TestComponentID=" & vars("testCmp") & "&PageMode=ProcessSelect&LabRequest=" & testRq & "&StoreField=" & vars("fd") & "&PadWidth=" & vars("src2") & "&PadHeight=" & vars("src3")
                vars("fdVl2") = "<iframe name=""iFrm" & secCnt & """ height=""" & vars("src3") & """ width=""" & vars("src2") & """ frameborder=""0"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"" src=""" & iFUrl & """></iframe>"
            ElseIf UCase(Trim(vars("src0"))) = "USERDICOM" Then
                vars("tdDim") = ""
                If IsNumeric(vars("src2")) Then
                    vars("tdDim") = vars("tdDim") & " width=""" & CStr(CInt(vars("src2")) + 30) & """ "
                End If
                If IsNumeric(vars("src3")) Then
                    vars("tdDim") = vars("tdDim") & " height=""" & CStr(CInt(vars("src3")) + 140) & """ "
                End If
                iFUrl = "wpgPrtPrintLayoutAll.asp?PositionforTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=dwvDicomViewer&LabRequestID=" & testRq & "&LabTestID=" & testDat & "&TestComponentID=" & vars("testCmp") & "&PadWidth=" & vars("src2") & "&PadHeight=" & vars("src3")
                vars("fdVl2") = "<iframe name=""iFrm" & secCnt & """ height=""" & vars("src3") & """ width=""" & vars("src2") & """ frameborder=""0"" scrolling=""yes"" marginwidth=""0"" marginheight=""0"" src=""" & iFUrl & """></iframe>"
            ElseIf Len(vars("fdVl")) > 0 Then
                'vars("fdVl2") = GetCompFldVal2(compKey, vars("testCmp"), vars("fd"), vars("fdVl"), "10", "1")
                'vars("fdVl2") = GetCompFldVal3(compKey, vars("testCmp"), vars("fd"), vars("fdVl"), "11", "1", srcKy, "LabTestID", testDat)
                vars("fdVl2") = GetCompFldVal(vars("compKey"), vars("testCmp"), vars("fd"), vars("fdVl"))
'                GetCompFldVal("TestComponentID", cmp, "Column" & CStr(colm), .fields("column" & CStr(colm)))
            End If

            If IsNumeric(vars("tdColSp")) Then
                For num = 1 To (CInt(vars("tdColSp")) - 1)
                    If Not rstCompFld.EOF Then
                        rstCompFld.MoveNext
                    End If
                Next
                ot = ot & "<td valign=""top"" " & vars("tdDim") & " " & vars("tdAlign") & " colspan=""" & vars("tdColSp") & """>" & vars("fdVl2") & "</td>"
            Else
                ot = ot & "<td valign=""top"" " & vars("tdDim") & " " & vars("tdAlign") & ">" & vars("fdVl2") & "</td>"
            End If
            If Not rstCompFld.EOF Then
                rstCompFld.MoveNext
            End If
        Loop
    End If
    ot = ot & "</tr>"

    Set rstCompFld = Nothing
    GetComponentRow = ot
End Function
'GetRequestedTest
Function GetRequestedTest(rec)
    Dim ot, rst, sql, cnt, nm, dt, lbTch, dia
    ot = ""

    sql = "select investigation.labtestid,investigation.labtechid,investigation.requestdate1,investigation.clinicaldiagnosis,labtest.labtestname "
    sql = sql & " from investigation,labtest"
    sql = sql & " where investigation.labtestid=labtest.labtestid and investigation.labrequestid='" & rec & "'"
    sql = sql & " order by investigation.labtestid"
    Set rst = CreateObject("ADODB.Recordset")
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
            .MoveFirst
            cnt = 0
            Do While Not .EOF
                cnt = cnt + 1
                If cnt > 1 Then
                    ot = ot & ", "
                End If
                nm = .fields("labtestname")
                dt = .fields("RequestDate1")
                lbTch = .fields("labtechid")
                dia = .fields("clinicaldiagnosis")
                ot = ot & Replace(nm, " ", "&nbsp;")
                If dt > lstDate Then
                    lstDate = dt
                    lstLbTech = lbTch
                End If
                If Not IsNull(dia) Then
                    If Len(Trim(dia)) > 0 Then
                        cliDia = "" 'dia
                    End If
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With

    '///////////Investigation2 /////////////
    sql = "select investigation2.labtestid,investigation2.labtechid,investigation2.requestdate1,investigation2.clinicaldiagnosis,labtest.labtestname "
    sql = sql & " from investigation2,labtest"
    sql = sql & " where investigation2.labtestid=labtest.labtestid and investigation2.labrequestid='" & rec & "'"
    sql = sql & " order by investigation2.labtestid"
    Set rst = CreateObject("ADODB.Recordset")
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
            .MoveFirst
            'cnt = 0 continue from above
            Do While Not .EOF
                cnt = cnt + 1
                If cnt > 1 Then
                    ot = ot & ", "
                End If
                nm = .fields("labtestname")
                dt = .fields("RequestDate1")
                lbTch = .fields("labtechid")
                dia = .fields("clinicaldiagnosis")
                ot = ot & Replace(nm, " ", "&nbsp;")
                If dt > lstDate Then
                    lstDate = dt
                    lstLbTech = lbTch
                End If
                If Not IsNull(dia) Then
                    If Len(Trim(dia)) > 0 Then
                        cliDia = "" 'dia
                    End If
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    GetRequestedTest = ot
End Function
Function VisitNoExempt(vst)
    Dim arr, ul, num, lst, ot
    ot = False
    lst = "0054444||0079669||0185870||0208211||0158043||0099346||0186748||0207940||E01||P1||E02||P2||V02||V03||V04||V05||E02"
    arr = Split(lst, "||")
    ul = UBound(arr)
    For num = 0 To ul
        If UCase(Trim(arr(num))) = UCase(Trim(vst)) Then
            ot = True
            Exit For
        End If
    Next
    VisitNoExempt = ot
End Function
'GetPatientAge
Function GetPatientAge(vst)
    Dim ot, rst, sql
    sql = "select patientage from visitation where visitationid='" & vst & "'"
    Set rst = server.CreateObject("ADODB.Recordset")
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
            ot = .fields("patientage")
        End If
        .Close
        GetPatientAge = CInt(ot)
    End With
End Function
'GetPatientAge
Function GetPatientAge2(vst)
    Dim ot, rst, sql
    sql = "select VisitInfo6 from visitation where visitationid='" & vst & "'"
    Set rst = server.CreateObject("ADODB.Recordset")
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
            ot = .fields("VisitInfo6")
        End If
        .Close
        GetPatientAge2 = ot
    End With
End Function
'GetRequestAge
Function GetRequestAge(vst)
    Dim ot, rst, sql
    sql = "select requestage from labrequest where labrequestid='" & vst & "'"
    Set rst = server.CreateObject("ADODB.Recordset")
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
            ot = .fields("requestage")
        End If
        .Close
        GetRequestAge = CInt(ot)
    End With
End Function
'GetRequestName
Function GetRequestName(vst)
    Dim ot, rst, sql
    sql = "select requestname from labrequest where labrequestid='" & vst & "'"
    Set rst = server.CreateObject("ADODB.Recordset")
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
            ot = .fields("requestname")
        End If
        .Close
        GetRequestName = ot
    End With
End Function
'GetRequestGender
Function GetRequestGender(vst)
    Dim ot, rst, sql
    sql = "select requestgenderid from labrequest where labrequestid='" & vst & "'"
    Set rst = server.CreateObject("ADODB.Recordset")
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
            ot = GetComboName("RequestGender", .fields("RequestGenderid"))
        End If
        .Close
        GetRequestGender = ot
    End With
End Function
Function DisplayHead(LabRequestID)
    Dim str, sql, rst, lstDate, cliDia, reqTest

    sql = GetTableSql("LabRequest") & " and LabRequest.LabRequestID='" & LabRequestID & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        'AddReportHeader
        str = str & "<table class='report-header' style='margin:auto;'>"
           str = str & "<tr><tr><td colspan=""100"">" & Glob_DisplayHeader2(ourRef, yourRef, GetRecordField("RequestDate"), "OFFICIAL REPORT") & "</td></tr>"
            str = str & "<tr><td align=""center"" style=""font-weight:bold;font-size:14pt"">MEDICAL LABORATORY REPORT</td></tr>"
            str = str & "<tr><td align=""center""><hr color=""#999999"" size=""1""></td></tr>"
            str = str & "<tr><td align=""center"">"
                str = str & "<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial"">"
                    str = str & "<tr>"

                        str = str & "<td name=""tdLabelInpLabReceiptID"" id=""tdLabelInpLabReceiptID"" style=""font-weight: bold"">LAB No.</td>"
                        str = str & "<td width=""10""></td>"
                        str = str & "<td style=""font-size:11pt"" name=""tdInputInpLabReceiptID"" id=""tdInputInpLabReceiptID""><b>" & rst.fields("LabReQUESTID") & "</b></td>"

                        str = str & "<td width=""10""></td>"

                        str = str & "<td name=""tdLabelInpGenderID"" id=""tdLabelInpGenderID"" style=""font-weight: bold"">Gender</td>"
                        str = str & "<td width=""10""></td>"

                        If VisitNoExempt(rst.fields("visitationid")) Or VisitNoExempt(Trim(rst.fields("Patientid"))) Then
                            str = str & "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & GetRequestGender(rst.fields("LabRequestID")) & "</td>"
                        Else
                            str = str & "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & rst.fields("GenderName") & "</td>"
                        End If

                        str = str & "<td width=""10""></td>"
                    str = str & "</tr>"
                    str = str & "<tr>"
                        str = str & "<td name=""tdLabelInpLabReceiptName"" id=""tdLabelInpLabReceiptName"" style=""font-weight: bold"">Patient Name</td>"
                        str = str & "<td width=""10""></td>"

                        If VisitNoExempt(rst.fields("visitationid")) Or VisitNoExempt(Trim(rst.fields("Patientid"))) Then
                            str = str & "<td style=""font-size:11pt"" name=""tdInputInpLabReceiptName"" id=""tdInputInpLabReceiptName""><b>" & GetRequestName(rst.fields("LabRequestID")) & "</b></td>"
                        Else
                            str = str & "<td style=""font-size:11pt"" name=""tdInputInpLabReceiptName"" id=""tdInputInpLabReceiptName""><b>" & rst.fields("pATIENTName") & "</b></td>"
                        End If


                        str = str & "<td width=""10""></td>"
                        str = str & "<td name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"" style=""font-weight: bold"">Receipt Type</td>"
                        str = str & "<td width=""10""></td>"
                        str = str & "<td name=""tdInputInpReceiptTypeID"" id=""tdInputInpReceiptTypeID""></td>" ' & rst.Fields("ReceiptTypeName") & "</td>"
                        str = str & "<td width=""10""></td>"
                    str = str & "</tr>"

                    str = str & "<tr>"
                        str = str & "<td name=""tdLabelInpAge"" id=""tdLabelInpAge"" style=""font-weight: bold"">Age</td>"
                        str = str & "<td width=""10""></td>"

                        If VisitNoExempt(rst.fields("visitationid")) Or VisitNoExempt(Trim(rst.fields("Patientid"))) Then
                            str = str & "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & GetRequestAge(rst.fields("LabRequestID")) & "</td>"
                        Else
                            str = str & "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & GetPatientAge(rst.fields("VisitationID")) & "</td>"
                        End If

                        str = str & "<td width=""10""></td>"

                        str = str & "<td name=""tdLabelInpInsuranceSchemeID"" id=""tdLabelInpInsuranceSchemeID"" style=""font-weight: bold"">Organization</td>"
                        str = str & "<td width=""10""></td>"
                        str = str & "<td name=""tdInputInpInsuranceSchemeID"" id=""tdInputInpInsuranceSchemeID""></td>" ' & rst.Fields("InsuranceSchemeName") & "</td>"
                        str = str & "<td width=""10""></td>"
                    str = str & "</tr>"

                    str = str & "<tr>"
                        str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Episode No.</td>"
                        str = str & "<td width=""10""></td>"
                        str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (rst.fields("VisitationID")) & "</td>"
                        str = str & "<td width=""10""></td>"
                        str = str & "<td name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"" style=""font-weight: bold"">Patient Tel:</td>"
                        str = str & "<td width=""10""></td>"
                        str = str & "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & (rst.fields("ContactNo")) & "</td>"
                        str = str & "<td width=""10""></td>"
                    str = str & "</tr>"

                    str = str & "<tr>"
                        str = str & "<td name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"" style=""font-weight: bold"">Manual Path. No</td>"
                        str = str & "<td width=""10""></td>"
                        str = str & "<td name=""tdInputInpReceiptTypeID"" id=""tdInputInpReceiptTypeID"">" & rst.fields("ReceiptInfo2") & "</td>"
                        str = str & "<td width=""10""></td>"
                        str = str & "<td name=""tdLabelInpInsuranceSchemeID"" id=""tdLabelInpInsuranceSchemeID"" style=""font-weight: bold""></td>"
                        str = str & "<td width=""10""></td>"
                        str = str & "<td name=""tdInputInpInsuranceSchemeID"" id=""tdInputInpInsuranceSchemeID""></td>"
                        str = str & "<td width=""10""></td>"
                    str = str & "</tr>"
                str = str & "</table>"
        str = str & "</td></tr>"
        'Doctor Info
        str = str & "<tr><td align=""center""><hr color=""#999999"" size=""1""></td></tr>"
        str = str & "<tr><td align=""center"">"
        str = str & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial;width:177mm"">"
        str = str & "<tr>"
        str = str & "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Requested By</td>"
        str = str & "<td width=""10""></td>"
        str = str & "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">" & rst.fields("DoctorName") & "</td>"
        str = str & "<td width=""10""></td>"
        str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Sample Collection Date</td>"
        str = str & "<td width=""10""></td>"
        str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & FormatDateDetail(rst.fields("RequestDate")) & "</td>"
        str = str & "<td width=""10""></td></tr>"

        str = str & "<tr>"
        str = str & "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Requested From</td>"
        str = str & "<td width=""10""></td>"
        str = str & "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID""></td>"
        str = str & "<td width=""10""></td>"
        str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Received Date</td>"
        str = str & "<td width=""10""></td>"
        str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">&nbsp;&nbsp;&nbsp;&nbsp;</td>"
        str = str & "<td width=""10""></td></tr>"

        str = str & "<tr>"
        str = str & "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Diagnosis</td>"
        str = str & "<td width=""10""></td>"
        str = str & "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">&nbsp;" & cliDia & "</td>"
        str = str & "<td width=""10""></td>"
        str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Report Date</td>"
        str = str & "<td width=""10""></td>"
        If IsDate(lstDate) Then
            str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & FormatDateDetail(lstDate) & "</td>"
        Else
            str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">&nbsp;</td>"
        End If
        str = str & "<td width=""10""></td></tr>"

        str = str & "<tr style=""display:none"">"
        str = str & "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Sample Collected At</td>"
        str = str & "<td width=""10""></td>"
        str = str & "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">SSNIT:" & rst.fields("BranchName") & "</td>"
        str = str & "<td width=""10""></td>"
        str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">&nbsp;</td>"
        str = str & "<td width=""10""></td>"
        str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">&nbsp;</td>"
        str = str & "<td width=""10""></td></tr>"
        str = str & "</table>"
        str = str & "</td></tr>"
        'Requested
        str = str & "<tr><td align=""center""><hr color=""#999999"" size=""1""></td></tr>"
        str = str & "<tr><td align=""left"">"
        str = str & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial;width:177mm;"">"
        str = str & "<tr><td style='width:80px;'><b>REQUESTED:</b>&nbsp;&nbsp;</td><td id='requestedTests' style='text-align:left;'></td></tr></table>"
        str = str & "</td></tr>"
        str = str & "</table>"

        DisplayHead = str
    End If
End Function



'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

str = "<table><tbody>"
str = str & "<tr><td><img src=""images/letterfoot.jpg"" style=""width:100%""></td></tr>"
str = "</tbody></table>"
response.write str

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
