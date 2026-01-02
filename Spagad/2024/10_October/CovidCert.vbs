'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim requestID, testID, patid

requestID = Trim(request.querystring("LabRequestID"))
testID = Trim(request.querystring("LabTestID"))
patid = Trim(request.querystring("PrintFilter0"))

If Len(requestID) > 0 And Len(testID) > 0 Then
    PrintCovidCert requestID, testID
Else
    AddPageJS
    ShowSearch patid
End If

Sub AddPageJS()
    Dim html
    
    html = "<script>"
    html = html & vbCrLf & " function upload_results(spn){"
    html = html & vbCrLf & "     let xmlHttp = new XMLHttpRequest();"
    html = html & vbCrLf & "     let url = spn.dataset.href;"
    html = html & vbCrLf & "     xmlHttp.onload = function(){"
    html = html & vbCrLf & "         let obj = JSON.parse(xmlHttp.responseText);"
    html = html & vbCrLf & "         if(obj['status'].toUpperCase() == 'OK'){"
    html = html & vbCrLf & "            spn.innerHTML = 'OK';"
    html = html & vbCrLf & "            spn.style.cssText = 'color:green';"
    html = html & vbCrLf & "            spn.onclick=function(){alert('done');};"
    html = html & vbCrLf & "         }"
    html = html & vbCrLf & "         "
    html = html & vbCrLf & "         "
    html = html & vbCrLf & "     };"
    html = html & vbCrLf & "     xmlHttp.open('GET', url);"
    html = html & vbCrLf & "     xmlHttp.send();"
    'html = html & vbCrLf & "     console.log(url);"
    html = html & vbCrLf & " }"
    html = html & vbCrLf & "</script>"
    
    response.write html
End Sub
Sub ShowSearch(patid)
    Dim sql, rst, html, rptGen
    
    sql = " select Patient.PatientID, Patient.PatientName as [Name], Patient.BirthDate as [Birth Date]"
    sql = sql & " , inv.LabRequestID as [Sample No.], LabTest.LabTestName as [Test Name], Patient.ResidencePhone as [Phone]"
    sql = sql & " , inv.LabTestID, inv.LabRequestID"
    sql = sql & " from Patient  "
    sql = sql & " inner join LabRequest on Patient.PatientID=LabRequest.PatientID and Patient.PatientID='" & patid & "'"
    sql = sql & " inner join ("
    sql = sql & "     select LabRequestID, LabTestID, VisitationID, PatientID from Investigation where PatientID='" & patid & "' and LabTestID like 'COV%' and Investigation.RequestStatusID='RRD002' "
    sql = sql & "     union select LabRequestID, LabTestID, VisitationID, PatientID from Investigation2 where PatientID='" & patid & "' and LabTestID like 'COV%' and Investigation2.RequestStatusID='RRD002'"
    sql = sql & " ) as inv on inv.PatientID=Patient.PatientID and inv.LabRequestID=LabRequest.LabRequestID "
    sql = sql & " left join LabTest on LabTest.LabTestID=inv.LabTestID"
    
    Set rptGen = New PRTGLO_RptGen
    args = "title=COVID Reports for [" & GetComboName("Patient", patid) & "]"
    args = args & ";ExtraFields=Actions"
    args = args & ";FieldFunctions=Actions:GetActions|Birth Date:GetDate"
    args = args & ";HiddenFields=LabTestIS|LabRequestID"
    
    rptGen.PrintSQLReport sql, args
    
End Sub
Function getDate(RECOBJ, fieldName)
    getDate = FormatDate(RECOBJ(fieldName))
End Function
Function GetActions(RECOBJ, fieldName)
    Dim ot, ltID, reqID, href
    
    ltID = RECOBJ("LabTestID")
    reqID = RECOBJ("LabRequestID")
    
    sql = "select Column2 from LabResults where LabRequestID='" & reqID & "' and LabTestID='" & ltID & "' and TestComponentID='COV001003' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        If Not (Len(rst.fields("Column2")) > 0) Then
            href = "wpgXMLHttp.asp?ProcedureName=SendCovidResultsToPanabios&LabTestID=" & ltID
            href = href & "&LabRequestID=" & reqID
        
            ot = ot & "<div><span data-href='" & href & "' style='color:red;text-transform:none;cursor:pointer;' onclick='upload_results(this)'>Upload to Panabios</a></div>"
        End If
    End If
    
    href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=COVIDCert&PositionForTableName=WorkingDay&LabRequestID=" & reqID
    href = href & "&LabTestID=" & ltID
    ot = ot & "<div><a href='" & href & "' style='text-decoration:none;text-transform:none;'>Print Cert.</a></div>"

    GetActions = ot
End Function
Sub PrintCovidCert(requestID, testID)

    Dim tmp, passportNo, tmp2
    
    Set tmp = GetResultsFieldValues(requestID, testID)
    
    response.write "    <script src=""scripts/qrcodejs/qrcode.min.js""></script>"
    response.write "    <style>"
    response.write "        .report {"
    response.write "            border-collapse: separate;"
    response.write "            width: 100%;"
    response.write "            border-spacing: 0px;"
    response.write "        }"
    response.write "        .report >tbody>tr>td {"
'    Response.Write "            font-size:13px;"
    response.write "            text-align: left;"
    response.write "            padding: 8px 10px;"
    response.write "            border-left: 1px solid silver;"
    response.write "            border-bottom: 1px solid silver;"
    response.write "        }"
    response.write "        .report>tbody>tr:first-child td {"
    response.write "            border-top: 1px solid silver;"
    response.write "        }"
    response.write "        .report>tbody>tr>td:last-child {"
    response.write "            border-right: 1px solid silver;"
    response.write "        }"
    response.write "        .report>tbody>tr>td.header {"
    response.write "            background-color: #f0f0f0;"
    response.write "            font-weight: bold;"
    response.write "            text-align: center;"
    response.write "        }"
    response.write ""
    response.write "        * {"
    response.write "            -webkit-print-color-adjust: exact;"
    response.write "            print-color-adjust: exact;"
    response.write "            color-adjust: exact !important;"
    response.write "        }"
    response.write "    </style>"
    response.write "    <div style=""margin: 0 auto;text-align: center;width: 180mm;font-family:Arial, Helvetica, sans-serif"">"
    response.write "        <div style=""visibility:hidden;"">"
    response.write "            <img src=""images/logo.jpg"" style=""width: 130px;height: 100px;"">"
    response.write "            <div style=""font-size: 16px;font-weight: bold;"">AIRPORT CLINIC LIMITED</div>"
    response.write "            <div style=""font-size: 11px;"">Private Mail Bag, Kotoka International Airport, Accra, Ghana</div>"
    response.write "            <div style=""font-size: 11px;"">Email: info@airportclinic.org || Website:www.airportclinic.org</div>"
    response.write "            <div style=""font-size: 11px;"">Tel: +233 302 764987</div>"
    response.write "        </div>"
    response.write "        <div style=""display: flex;justify-content: space-between;text-align: center;"">"
    response.write "            <div>"
    response.write "                <div id=""pana-qr-code""></div>"
    response.write "                <div id=""pana-user-code"" style=""margin-top:10px;font-size:12px;"">" & tmp("COV001003Column2") & "</div>"
    response.write "            </div>"
    response.write "            <div>"
    response.write "                <div id=""lab-qr-code""></div>"
    response.write "                <div id=""lab-sample-code"" style=""margin-top:10px;font-size:12px;"">" & tmp("LabRequestID") & "</div>"
    response.write "            </div>"
    response.write "        </div>"
    response.write "        <div style=""visibility:hidden;height: 70px;font-size: 16px;font-weight: bold;text-align: center;vertical-align: middle;line-height: 50px;"">COVID-19 Certificate</div>"
    response.write "        <table class=""report"">"
    response.write "            <tbody>"
    response.write "                <tr>"
    response.write "                    <td colspan=""2"" class=""header"">PATIENT INFORMATION</td>"
    response.write "                </tr>"
    response.write "                <tr>"
    response.write "                    <td style=""width: 30%;"">Name of Client</td>"
    response.write "                    <td style=""width:70%;"">" & tmp("PatientName") & "</td>"
    response.write "                </tr>"
    response.write "                <tr>"
    response.write "                    <td>Gender</td>"
    response.write "                    <td>" & tmp("GenderName") & "</td>"
    response.write "                </tr>"
    response.write "                <tr>"
    response.write "                    <td>Passport Number</td>"
    
    'response.write "                    <td>" & GetCompFldVal("TestComponentID", "COV001001", "Column2", tmp("COV001001Column2")) & "</td>"
    tmp2 = UCase(GetComboNameFld("Patient", tmp("PatientID"), "PatientInsInfo2"))
    If Len(tmp2) > 0 Then
        If InStr(tmp2, "PASSPORT: ") > 0 Then
            passportNo = Replace(tmp2, "PASSPORT: ", "")
        End If
    End If
    
    response.write "                    <td>" & passportNo & "</td>"
    response.write "                </tr>"
    response.write "                <tr>"
    response.write "                    <td>Date of Birth</td>"
    response.write "                    <td>" & FormatDate(GetComboNameFld("Patient", tmp("PatientID"), "BirthDate")) & "</td>"
    response.write "                </tr>"
    response.write "                <tr>"
    response.write "                    <td colspan=""2"" class=""header"">LABORATORY ASSESSMENT DATA</td>"
    response.write "                </tr>"
    response.write "                <tr>"
    response.write "                    <td>Test Requested</td>"
    response.write "                    <td>" & GetComboName("TestCompTab", testID) & "</td>"
    response.write "                </tr>"
    response.write "                <tr>"
    response.write "                    <td>Specimen Type</td>"
    response.write "                    <td>" & GetCompFldVal("TestComponentID", "COV001003", "Column5", tmp("COV001003Column5")) & "</td>"
    response.write "                </tr>"
    response.write "                <tr>"
    response.write "                    <td>Date & Time of Sample Collection</td>"
'    Response.Write "                    <td>" & GetCompFldVal("TestComponentID", "COV001004", "Column2", tmp("COV001004Column2")) & "</td>"
    
    sDt = Trim(tmp("COV001004Column2"))
    If IsDate(sDt) Then
        sDt = FormatDateDetail(CDate(sDt))
    End If
    response.write "                    <td>" & sDt & "</td>"
    response.write "                </tr>"
    response.write "                <tr>"
    response.write "                    <td>Result</td>"
    response.write "                    <td style='font-weight:bold;'>" & GetCompFldVal("TestComponentID", "COV001005", "Column2", tmp("COV001005Column2")) & "</td>"
    response.write "                </tr>"
    response.write "                <tr>"
    response.write "                    <td>Remarks</td>"
    response.write "                    <td>" & GetCompFldVal("TestComponentID", "COV001006", "Column2", tmp("COV001006Column2")) & "</td>"
    response.write "                </tr>"
    response.write "            </tbody>"
    response.write "        </table>"
    response.write "        <div style=""margin-top: 5px;font-size:11px;"">YOU MAY KINDLY CONTACT <b>0244-485843</b> FOR ANY ENQUIRIES ABOUT YOUR TEST RESULT.</div>"
    response.write "        <div style=""visibility:hidden;display: flex;justify-content:space-between;"">"
    response.write "            <div style=""float: left;"">"
    response.write "                <div style=""height: 100px;width:200px;border-bottom:1px dotted black;""></div>"
    response.write "                <div>Signature & Stamp</div>"
    response.write "            </div>"
    response.write "            <div style=""float: right;"">"
    response.write "                <div style=""height: 100px;width:200px;border-bottom:1px dotted black""></div>"
    response.write "                <div>Date</div>"
    response.write "            </div>"
    response.write "        </div>"
    response.write "        <div>"
    response.write ""
    response.write "        </div>"
    response.write "    </div>"
    response.write ""
    response.write "    <script>"
    response.write "        let panaqrcode = new QRCode(document.getElementById(""pana-qr-code""), {"
    response.write "            width: 70,"
    response.write "            height: 70"
    response.write "        });"
    response.write "        let elText = document.getElementById(""pana-user-code"");"
    response.write "        if(elText.innerHTML.length>0){panaqrcode.makeCode(elText.innerHTML);}"
    response.write ""
    response.write "        var labqrcode = new QRCode(document.getElementById(""lab-qr-code""), {"
    response.write "            width: 70,"
    response.write "            height: 70"
    response.write "        });"
    response.write "        elText = document.getElementById(""lab-sample-code"");"
    response.write "        if(elText.innerHTML.length>0){labqrcode.makeCode(elText.innerHTML);}"
    response.write "    </script>"

End Sub
Function GetTableSql2(TableID) 'override
    Dim otSql, sql, rst, rst2, jnSql, selSql
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    sql = "select TableFieldID, ForeignKeyField, TransientField from TableField where TableID='" & TableID & "' order by LabelPos asc"
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            If selSql <> "" Then selSql = selSql & ", "
            If UCase(rst.fields("ForeignKeyField")) <> "NONE" And Len(rst.fields("ForeignKeyField")) > 0 Then
                selSql = selSql & TableID & "." & rst.fields("TableFieldID")

                'TransientField
                sql = "select TableFieldID from TableField where TableID='" & rst.fields("ForeignKeyField") & "' and TransientField='yes' "
                rst2.open qryPro.FltQry(sql), conn, 3, 4
                If rst2.RecordCount > 0 Then
                    rst2.movefirst
                    selSql = selSql & ", " & rst.fields("ForeignKeyField") & "." & rst2.fields("TableFieldID")
                    rst2.Close
                End If

                jnSql = jnSql & " left join " & rst.fields("ForeignKeyField")
                jnSql = jnSql & " on " & rst.fields("ForeignKeyField") & "." & rst.fields("TableFieldID") & "=" & TableID & "." & rst.fields("TableFieldID")
            Else
                selSql = selSql & TableID & "." & rst.fields("TableFieldID")
            End If
            rst.MoveNext
        Loop
        otSql = " select " & selSql & " from " & TableID & " " & jnSql
        otSql = otSql & " where 1=1 "
        rst.Close
    End If
    Set rst = Nothing
    GetTableSql2 = otSql
End Function
Function GetResultsFieldValues(requestID, labTestID)
    Dim ot, tmp, rst, sql, field

    Set ot = CreateObject("Scripting.Dictionary")
    ot.CompareMode = 1 'Case Insensitive
    sql = GetTableSql2("LabResults") & " and (LabResults.LabRequestID='" & requestID & "' and LabResults.LabTestID= '" & labTestID & "' ) "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            For Each field In rst.fields
                If UCase(Left(field.name, 6)) = "COLUMN" Then
                    'tmp = Split(field.Value, "||")
                    'If Ubound(tmp) > 0 Then
                    '    ot(rst.Fields("TestComponentID") & field.Name) = tmp
                    'Else
                        ot(rst.fields("TestComponentID") & field.name) = field.value
                    'End If
                ElseIf Not ot.Exists(field.name) Then
                    ot(field.name) = field.value
                End If
            Next
            rst.MoveNext
        Loop
        rst.Close
        sql = GetTableSql2("LabRequest") & " and (LabRequest.LabRequestID='" & requestID & "' )  "
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount = 1 Then
            rst.movefirst
            For Each field In rst.fields
                ot(field.name) = field.value
            Next
        End If
    Else
    End If
    Set GetResultsFieldValues = ot
End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
