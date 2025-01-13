'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
' response.write Glob_GetBootstrap5()
' response.write Glob_GetIconFontAwesome()
Dim rptType
rptType = (Trim(Request.QueryString("PrintFilter0")))
addJS

response.write "<tr><td>"
response.write "<table class=""table table-striped cmpTdSty"" cellpadding=""2"" border=""1"" cellspacing=""0"" width=""100%"" style=""font-size:10pt"">"
' PrintReport
DisplayDiagnoses rptType
response.write "</table></td></tr>"
response.write "</td></tr>"


Sub printReport()


End Sub

Function GetReportCategoryList(flt)
        Dim rst, sql, ot
        Set rst = CreateObject("ADODB.Recordset")
        ot = "||**"
        sql = "select * from TestVar1B Where TestVar1AID='" & flt & "' "
        sql = sql & "  "
        With rst
            rst.open qryPro.FltQry(sql), conn, 3, 4
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    ot = ot & rst.fields("TestVar1BID") & "||" & rst.fields("TestVar1BName")
                    ot = ot & "**"
                    rst.MoveNext
                Loop
            End If
            rst.Close
        End With
        Set rst = Nothing
        GetReportCategoryList = ot
End Function


Sub DisplayDiagnoses(flt)
    Dim rst, sql, whcls
    Set rst = CreateObject("ADODB.Recordset")
    whcls = ""

    '''@ eben 22/11/2022
    ' sql = "select DISTINCT top 100 d.DiseaseName, dg.DiseaseID, d.DiseaseTypeID  "
    ' sql = "select DISTINCT top 5000 d.DiseaseName, dg.DiseaseID, d.DiseaseTypeID  ">>>>>>>>ORIGINAL
    sql = "select DISTINCT d.DiseaseName, dg.DiseaseID, d.DiseaseTypeID  "
    sql = sql & " from Diagnosis dg, Disease d where d.DiseaseID=dg.DiseaseID " '' " and dg.WorkingMonthID>='MTH202202' "
    'sql = sql & " and dg.workingyearid = '" & wrkYr & "' "
    ' sql = sql & " And dg.DiseaseID IN (Select DISTINCT PerformVar15Name from PerformVar15 Where PerformVar15ID Like '" & flt & "::%') "
    ' sql = sql & " And dg.DiseaseID IN (Select DISTINCT PerformVar15Name from PerformVar15 Where KeyPrefix='" & flt & "') "
    sql = sql & whcls & " order by d.DiseaseName "
    lstOpt = GetReportCategoryList(flt)
    ' response.write sql

    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            Do While Not rst.EOF
                ds = rst.fields("DiseaseID")
                ins = flt '' rst.fields("DiseaseTypeID")
                id = "||" & ins & "||" & ds
                opt = GetComboNameFld("PerformVar15", flt & "::" & ds, "KeyPrefix")
                pStat = "" '' rst.fields("PermuteStatusID")
                response.write "<tr class=""" & pStat & """>"
                response.write "<td>" & rst.AbsolutePosition & "</td>"
                ' response.write "<td>" & BuildInputElement("DiseaseID", "hidden", ds, id, "", "") & rst.fields("DiseaseName") & "</td>"
                response.write "<td>" & rst.fields("DiseaseID") & "</td>"
                response.write "<td>" & BuildInputElement("DiseaseID", "hidden", ds, id, "", "") & rst.fields("DiseaseName") & "</td>"
                response.write "<td>" & BuildComboInput("PerformVar15ID", opt, lstOpt, id, "") & "</td>"

                ' response.write "<td>" & BuildInputElement("DrugStoreTypeID", "hidden", ds, id, "", "") & rst.fields("DrugStoreTypeName") & "</td>"
                ' response.write "<td>" & BuildInputElement("InsuranceTypeID", "hidden", ins, id, "", "") & rst.fields("InsuranceTypeName") & "</td>"
                ' response.write "<td>" & BuildInputElement("ItemUnitCost", "number", rst.fields("ItemUnitCost"), id, "", 5) & "</td>"
                ' response.write "<td>" & BuildInputElement("Column1", "", rst.fields("Column1"), id, "", 10) & "</td>"
                ' response.write "<td>" & BuildInputElement("Column2", "", rst.fields("Column2"), id, "", 30) & "</td>"
                ' response.write "<td>" & BuildComboInput("PermuteStatusID", pStat, "P001||Valid**P002||Invalid", id, "") & "</td>"

                response.write "<td><button class=""btn btn-secondary"" onclick=""DoAjaxUpdate('" & ds & "', '" & ins & "', '" & flt & "', '" & id & "', this)"">Update</button></td>"
                response.write "</tr>"
                response.flush
                rst.MoveNext
            Loop
        Else

        End If
        rst.Close
    End With
    Set rst = Nothing
End Sub

Function BuildComboInput(fld, opt, lstOpt, id, evt)
    Dim sltHt
    sltHt = "<select size=""1"" name=""inp" & fld & id & """ id=""inp" & fld & id & """ " & evt & " class=""form-control selectpicker"" data-live-search=""true"">"
    arOpt = Split(lstOpt, "**")
    For Each lst In arOpt
        arLst = Split(lst, "||")
        Selected = ""
        If UBound(arLst) >= 1 Then
            If UCase(opt) = UCase(arLst(0)) Then
                Selected = " selected "
            End If
            ' sltHt = sltHt & "<option value=""" & arLst(0) & """ " & Selected & ">" & arLst(1) & "</option>"
            ' sltHt = sltHt & "<option data-tokens=""" & arLst(0) & """ " & Selected & ">" & arLst(1) & "</option>"
            sltHt = sltHt & "<option value=""" & arLst(0) & """ data-tokens=""" & arLst(0) & """ " & Selected & ">" & arLst(1) & "</option>"
        End If
    Next
    sltHt = sltHt & "</select>"

    BuildComboInput = sltHt
End Function

Function BuildInputElement(fld, typ, val, id, evt, size_)
    Dim html, size
    html = ""
    size = (size_)
    Select Case UCase(typ)
        Case UCase("Hidden")
            If Not IsNumeric(size_) Then size = 20
            ' html = "<input type=""hidden"" name=""" & id & """ id=""inp" & id & """ "
            html = "<input type=""hidden"" name=""PrintLayoutID"" id=""inpPrintLayoutID" & id & """ "
            html = html & " value=""" & val & """>"
        Case UCase("TextArea")
            If Not IsNumeric(size_) Then size = 3
            html = "<textarea name=""inp" & fld & id & """ id=""inp" & fld & id & """ col=""20"" class=""form-control"" "
            html = html & " rows=""" & size & """ " & evt & ">" & val & "</textarea>"
        Case UCase("Date")
            If Not IsNumeric(size_) Then size = 20
            html = "<input type=""date"" name=""inp" & fld & id & """ id=""inp" & fld & id & """ class=""form-control"" "
            html = html & " value=""" & val & """ size=""" & size & """ " & evt & ">"
        Case UCase("Time")
            If Not IsNumeric(size_) Then size = 20
            html = "<input type=""time"" name=""inp" & fld & id & """ id=""inp" & fld & id & """ class=""form-control"" "
            html = html & " value=""" & val & """ size=""" & size & """ " & evt & ">"
        Case UCase("Number")
            If Not IsNumeric(size_) Then size = 20
            html = "<input type=""number"" name=""inp" & fld & id & """ id=""inp" & fld & id & """ class=""form-control"" "
            html = html & " value=""" & val & """ size=""" & size & """ " & evt & ">"
        Case Else ''Text
            If Not IsNumeric(size_) Then size = 30
            html = "<input type=""text"" name=""inp" & fld & id & """ id=""inp" & fld & id & """ class=""form-control"" "
            html = html & " value=""" & val & """ size=""" & size & """ " & evt & ">"
    End Select
    BuildInputElement = html
End Function


Sub addJS()
    Dim js
    js = ""
    js = js & "<script>" & vbNewLine
    js = js & " function DoAjaxUpdate(ds, ins, drg, id, ele) { console.log(ele); " & vbNewLine
        js = js & "      UpdateDiseaseMappingInfo(ds, ins, id); " & vbNewLine
        js = js & "     function UpdateDiseaseMappingInfo(ds, ins, id) { " & vbNewLine
        ' js = js & "         alert('Here'); " & vbNewLine
        js = js & "         var url, getStr, fnd, inp, fd, vl; " & vbNewLine
        js = js & "         if (Helpers.len(drg) > 0) { " & vbNewLine
            js = js & "             getStr = 'ProcedureName=UpdateDiseaseMappingInfoClient&DiseaseID=' + ds; " & vbNewLine
        js = js & "             getStr = getStr + '&TestVar1AID=" & rptType & "'; " & vbNewLine
        js = js & "             getStr = getStr + '&TestVar1BID=' + document.getElementById('inpPerformVar15ID'+id).value; " & vbNewLine
        js = js & "             url = 'wpgXmlHttp.' + appfilext + '?' + getStr; " & vbNewLine
        js = js & " console.log(url); " & vbNewLine
        js = js & "             Helpers.xmlhttprequest(Helpers.ucase('GET'), url,UpdateDiseaseMappingInfoCont); " & vbNewLine
        js = js & "         } " & vbNewLine
        js = js & "     } " & vbNewLine
        js = js & "     function UpdateDiseaseMappingInfoCont(readyState, responseText) { " & vbNewLine
        js = js & "         var arr, ul, num, str, rec; " & vbNewLine
        js = js & "         if (readyState == 4) { " & vbNewLine
        js = js & "             str = ReplaceXmlHttpComment(responseText); " & vbNewLine
        js = js & "             arr = Split(str, delim(1)); " & vbNewLine
        js = js & "             ul = UBound(arr); " & vbNewLine
        js = js & "             if (ul >= 0) { " & vbNewLine
        js = js & "                 rec = Helpers.trim(arr(0)); console.log(rec); " & vbNewLine
        js = js & "                 if (Helpers.ucase(rec) == Helpers.ucase('True')) { " & vbNewLine
        js = js & "                     if (ele) { " & vbNewLine
        js = js & "                         ele.setAttribute('class', 'btn btn-light'); " & vbNewLine
        js = js & "                     } " & vbNewLine
        js = js & "                     ele = null; " & vbNewLine
        js = js & "                 } else { " & vbNewLine
        js = js & "                     ele.setAttribute('class', 'btn btn-warning'); " & vbNewLine
        js = js & "                 } " & vbNewLine
        js = js & "             } " & vbNewLine
        js = js & "         } " & vbNewLine
        js = js & "     } " & vbNewLine
        js = js & " } " & vbNewLine
    js = js & " ;" & vbNewLine
    js = js & " ;" & vbNewLine
    js = js & "</script>" & vbNewLine
    response.write js
End Sub
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
