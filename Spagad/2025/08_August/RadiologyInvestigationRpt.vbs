'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

If TestIsValidated() Then
        Response.write PageRules
        Response.write reportHeader
'    Response.write "<page>"
        Response.write "<table>"
            Response.write TestInfo
            Response.write TestResults
            Response.write PageFooter
        Response.write "</table>"
'    Response.write "</page>"
Else
    Response.write "<br/><br/><font color='red' style=""font-weight: bold; font-size: x-large;"">Investigation has not been validated.</font>"
End If

''''''''' HELPERS '''''''''
Function PageFooter()
    Dim str
    
    'str = "<hr style=""width: 100%;"" color=""#999999"" size=""1"" />"
    str = "<tfoot class=""pagefooter""><tr><td></td></tr>"
    str = str & "<tr><td colspan='8'><hr color=""#999999"" size=""1""></td></tr>"
    str = str & "<tr><td colspan='8'><div><h5 class=""pagefooter""></h6></div></tr></tfoot>"
    PageFooter = str
End Function

Function TestIsValidated()
    TestIsValidate = False
    If UCase(GetRecordField("requeststatusid")) = "RRD002" Then
        TestIsValidated = True
    End If
End Function
Function PageRules()
    Dim str
    str = str & "<style>"
        str = str & " html * { font-size: 12pt !important; }"
        str = str & " @media print{ .avoidpagebreak { clear: both; page-break-before: avoid; } html .rpthead * {filter: grayscale(0%);}}"
        str = str & " body{counter-reset: headingnumber;}"
        str = str & " @page{ size: A4; margin-bottom: 0.3in; }"
        'str = str & " @page::before{ content: ""help""; }"
        str = str & " @page :left{ margin-top: 1in;}"
        str = str & " h4{ /*page-break-after: avoid;*/ text-align: left; text-transform: capitalize; margin-bottom: 2px; margin-top: 5px;}"
        str = str & " h4::before{counter-increment: headingnumber; counter-reset: subheadingnumber; content: counter(headingnumber) "". "";}"
        str = str & " h5{ /*page-break-after: avoid;*/ text-align: left; margin-left: 0.3in; margin-bottom: 0px; margin-top: 0px; text-transform: capitalize; }"
        str = str & " h5.contentgroup{font-style: italic;}"
        'str = str & " h5.contentgroup::before{counter-increment: subheadingnumber; content: counter(headingnumber) "". "" counter(subheadingnumber) "" "" ;}"
        str = str & " p{ text-align: justify; margin-left: 0.5in; margin-top: 2px; margin-bottom: 2px; line-height: 20px;}"
        str = str & " table.testInfo{ width: 100%;} "
        'str = str & " img.rpthead{ max-width: 100%; max-height: 100%;} "
        str = str & " div.rpthead{ /*max-width: 100%; max-height: 100%;*/ padding-bottom: 5px;} "
        str = str & " .contentgroup{ /*page-break-inside: avoid; page-break-after: avoid; */}"
        str = str & " .sticktosibling{ /*page-break-before: avoid; page-break-after: avoid; */}"
        str = str & " h5.pagefooter::after{ content: 'Electronically Signed By: " & GetComboName("LabTech", GetRecordField("LabTechID")) & "'; }"
        str = str & " h5.pagefooter{ text-align: right;  }"
        'str = str & " h5.pagefooter::before{ content: none;  }"
        str = str & " .pagefooter{ text-align: right;}"
        str = str & " b u{ text-decoration: none;  text-transform: capitalize;}"
'        str = str & " .sameline h5{display: inline; vertical-align: top;}"
'        str = str & " .sameline p{display: inline; vertical-align: top;}"
    str = str & "</style>"
'    str = str & " body{"
'                    str = str & " <style>body{counter-reset: headingnumber;"
'                str = str & " }"
'                str = str & " @page{"
'                    str = str & " size: A4;"
'
'                str = str & " }"
'                str = str & " h4{"
'                    str = str & " page-break-before: always;"
'                str = str & " }"
'                str = str & " h4::before{"
'                    str = str & " counter-increment: headingnumber;"
'               str = str & "      content: counter(headingnumber) "" "";"
'               str = str & " }</style>"
    PageRules = str
End Function

Function reportHeader()
    Dim str
    str = str & "<div class=""rpthead"">"
    str = str & " <img class=""rpthead"" src=""images/IMaH_Letterhead6.png"">"
    str = str & " </div>"
    reportHeader = str
End Function

Function TestInfo()
    Dim str
    lbt = Trim(Request.QueryString("LabTestID"))
'    str = str & "<table class=""testinfo"">"
        str = str & "<thead>"
            'str = str & "<tr><td><img src=""images/IMaH_Letterhead6.png""></td><tr>"
            'str = str & "<tr><div style='height:150px;'></div></tr>"
            'str = str & "<tr><th colspan=""8"" align=""center"">RADIOLOGY REPORT</th></tr>"
            str = str & "<tr><td align=""center"" colspan=""8""><hr color=""#999999"" size=""1""></td></tr>"
'        str = str & "</thead>"
'        str = str & "<tbody> "
            
            ' name & age
            str = str & "<tr>"
                'name
                str = str & "<td name=""tdLabelInpLabReceiptName"" id=""tdLabelInpLabReceiptName"" style=""font-weight: bold"">NAME :</td>"
                str = str & "<td width=""10""></td>"
                If VisitNoExempt(GetRecordField("visitationid")) Or VisitNoExempt(Trim(GetRecordField("Patientid"))) Then
                    str = str & "<td style=""font-size:11pt"" name=""tdInputInpLabReceiptName"" id=""tdInputInpLabReceiptName""><b>" & GetRequestName(GetRecordField("LabRequestID")) & "</b></td>"
                Else
                    str = str & "<td style=""font-size:11pt"" name=""tdInputInpLabReceiptName"" id=""tdInputInpLabReceiptName""><b>" & GetRecordField("pATIENTName") & "</b></td>"
                End If
                str = str & "<td width=""10""></td>"
            
            'Age
                str = str & "<td name=""tdLabelInpAge"" id=""tdLabelInpAge"" style=""font-weight: bold"">AGE :</td>"
                str = str & "<td width=""10""></td>"
                If VisitNoExempt(GetRecordField("visitationid")) Or VisitNoExempt(Trim(GetRecordField("Patientid"))) Then
                    str = str & "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & GetRequestAge(GetRecordField("LabRequestID")) & "</td>"
                Else
                    str = str & "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & GetPatientAge(GetRecordField("VisitationID")) & "</td>"
                End If
                str = str & "<td width=""10""></td>"
            str = str & "</tr>"
            

            'gender & request no
            str = str & "<tr>"
                'Gender
                str = str & "<td name=""tdLabelInpGenderID"" id=""tdLabelInpGenderID"" style=""font-weight: bold"">GENDER :</td>"
                str = str & "<td width=""10""></td>"
                If VisitNoExempt(GetRecordField("visitationid")) Or VisitNoExempt(Trim(GetRecordField("Patientid"))) Then
                    str = str & "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & GetRequestGender(GetRecordField("LabRequestID")) & "</td>"
                Else
                    str = str & "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & GetRecordField("GenderName") & "</td>"
                End If
                str = str & "<td width=""10""></td>"
                
                'Req No
                str = str & "<td name=""tdLabelInpLabReceiptID"" id=""tdLabelInpLabReceiptID"" style=""font-weight: bold"">REQUEST NO :</td>"
                str = str & "<td width=""10""></td>"
                str = str & "<td style=""font-size:11pt"" name=""tdInputInpLabReceiptID"" id=""tdInputInpLabReceiptID""><b>" & GetRecordField("LabReQUESTID") & "</b></td>"
                str = str & "<td width=""10""></td>"
            str = str & "</tr>"

            'radiologist & date
            
            str = str & "<tr>"
                'str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">" & TryGetRole(GetRecordField("LabTechID")) & ":</td>"
                If lbt <> "L0157" Then
                    str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">RADIOLOGISTS:</td>"
                    str = str & "<td width=""10""></td>"
                    'str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & GetComboName("LabTech", (GetRecordField("LabTechID"))) & "</td>"
                    'str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & GetComboName("Staff", "MH1708010") & "</td>"
                    str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">DR. BECKY .A. APPIAH <br> DR. ESI DE GRAFT-JOHNSON </td>"
                    str = str & "<td width=""10""></td>"
                End If
                str = str & "<td name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"" style=""font-weight: bold"">DATE :</td>"
                str = str & "<td width=""10""></td>"
                str = str & "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & FormatDate(GetRecordField("RequestDate")) & "</td>"
                str = str & "<td width=""10""></td>"
            str = str & "</tr>"
            
            'visit no. & phone no
            str = str & "<tr>"
                str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">VISIT NO. :</td>"
                str = str & "<td width=""10""></td>"
                str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (GetRecordField("VisitationID")) & "</td>"
                str = str & "<td width=""10""></td>"
                str = str & "<td name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"" style=""font-weight: bold"">PATIENT TEL :</td>"
                str = str & "<td width=""10""></td>"
                str = str & "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & (GetRecordField("ContactNo")) & "</td>"
                str = str & "<td width=""10""></td>"
            str = str & "</tr>"
            'investigation name
            str = str & "<tr><td align=""center"" colspan=""8""><hr color=""#999999"" size=""1""></td></tr>"
            str = str & "<tr>"
                'str = str & "<td colspan=""7"" align=""left"" style=""" & sTy & """><b>INVESTIGATION NAME</b> : " & GetComboName("labtest", GetRecordField("LabTestID")) & "</td>"
                str = str & "<td colspan=""7"" align=""left"" style=""" & sTy & """><b>PART EXAMINED</b> : " & GetComboName("labtest", GetRecordField("LabTestID")) & "</td>"
            str = str & "</tr>"
            str = str & "<tr><td align=""center"" colspan=""8""><hr color=""#999999"" size=""1""></td></tr>"
'        str = str & "</tbody> "
'    TestInfo = str & "</table>"
    TestInfo = str & "</thead>"
End Function

Function TestResults()
    Dim str, sql, rst, ot
    Dim outTest, outTestA
    Dim outText
    Dim testName
    str = ""
    
    sql = "SELECT Column1, Column2, Column3, Column4, Column5, Column6, TestComponentID FROM LabResults "
    sql = sql & "   WHERE 1=1"
    sql = sql & "   AND LabRequestID='" & GetRecordField("LabRequestID") & "' "
    sql = sql & "   AND LabTestID='" & GetRecordField("LabTestID") & "' "
    sql = sql & "   ORDER BY CompPos "
    
    Set rst = CreateObject("ADODB.RecordSet")
    
    
    Response.write "<tbody>"
    rst.open sql, conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF ' for each testcomponent
            
            'str = str & "<div class=""grouptogether"">"
            If InStr(rst.fields("TestComponentID"), GetRecordField("LabTestID")) Then 'is a sub test
                
                testName = Replace(rst.fields("Column1"), "<br/>", "")
                testName = Replace(rst.fields("Column1"), "<br>", "")
                str = str & "<tr><td td colspan='8'><h4 class=""contentgroup"">" & SentenseCase(testName) & "</h4></td></tr>"
            Else
                
                outTest = rst.fields("Column1")
                
                
                outTestA = GetComboName("TestVar3B", Replace(outTest, "'", "''"))
                If Err.number <> 0 Then
                    outTestA = ""
                End If
                
                If (outTestA) = "" Then
                    outTest = outTest
                Else
                    outTest = outTestA
                End If
'                'preserve column 1 as label or title
                str = str & "<tr><td td colspan='8'><h5 class=""contentgroup"" style=''>" & SentenseCase(outTest) & "</h5></td></tr>"
                'column 2 as main body

                For i = 2 To 6
                    outText = rst.fields("Column" & i)
                    
                    If Len(outText) > 0 Then
                        Dim outTextA
                        
                        outTextA = GetComboName("TestVar3B", Replace(outText, "'", "''"))
                        If (outTextA) = "" Then
                            outText = outText
                        Else
                            outText = outTextA
                        End If
                        str = str & "<tr><td td colspan='8'><p class=""sticktosibling"" style=''>" & _
                                Replace((outText), vbCrLf, "</p></td></tr><tr><td td colspan='8'><p>") _
                                & "</p></td></tr>"
                    End If
                Next
            End If
            'str = "<tr class='page-break-inside: auto;'><td colspan='8'>" & str & "</td></tr>"
            rst.MoveNext
            Response.write str
            'ot = ot & str
            
            'If UCase(uname) = UCase(jSchd) Then Response.Write "<pre>" & str & "</pre>"
    
            str = ""
        Loop
        rst.Close
        Set rst = Nothing
    End If
    'TestResults = "<tbody>" & ot & "</tbody>"
    Response.write "</tbody>"
End Function

Function SentenseCase(sentence)
    SentenseCase = CapFirstLetter(sentence)
End Function

Function Capitalize(name)
    Dim namesplit: namesplit = Split(name, " ")
    Dim output
    For x = 0 To UBound(namesplit)
        output = output & " " & CapFirstLetter(namesplit(x))
    Next
    
    If InStr(output, "-") Then
        Dim namesplit2: namesplit2 = Split(output, "-")
        output = ""
        For x = 0 To UBound(namesplit2)
            output = output & "-" & CapFirstLetterOnly(namesplit2(x))
        Next
        If Left(output, 1) = "-" Then
            output = Mid(output, 2)
        End If
    End If
    
    Capitalize = output
End Function

Function CapFirstLetter(word)
    'Response.Write word
    CapFirstLetter = UCase(Mid(word, 1, 1)) & "" & LCase(Mid(word, 2))
End Function
Function CapFirstLetterOnly(word)
    CapFirstLetterOnly = UCase(Mid(word, 1, 1)) & "" & (Mid(word, 2))
End Function


'GetPatientAge
Function GetPatientAge(vst)
    Dim ot, rst, sql
    sql = "select patientage from visitation where visitationid='" & vst & "'"
    Set rst = Server.CreateObject("ADODB.Recordset")
    With rst
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
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
    Set rst = Server.CreateObject("ADODB.Recordset")
    With rst
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
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
    Set rst = Server.CreateObject("ADODB.Recordset")
    With rst
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
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
    Set rst = Server.CreateObject("ADODB.Recordset")
    With rst
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
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
    Set rst = Server.CreateObject("ADODB.Recordset")
    With rst
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
            ot = GetComboName("RequestGender", .fields("RequestGenderid"))
        End If
        .Close
        GetRequestGender = ot
    End With
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

Function TryGetRole(staffID)
    role = UCase(GetComboNameFld("Staff", staffID, "Address"))
    If IsEmpty(role) Then
        role = "Radiologist"
    End If
    TryGetRole = role
End Function



'Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
'Empty
'Response.write "<tfoot class=""pagefooter""><tr><td></td></tr><tr><td><hr color=""#999999"" size=""1""></td></tr><tr><td><div><h5 class=""pagefooter""></h6></div></tr></tfoot>"
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
