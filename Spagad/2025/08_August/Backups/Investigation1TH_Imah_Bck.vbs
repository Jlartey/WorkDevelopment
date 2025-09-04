'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Sub DisplayHead()
    'AddReportHeader
    Dim str
    str = ""
    'str = str & "<tr><td align=""center"" height=""100""></td></tr>" '"<tr><td align=""center"" height=""10""></td></tr>"
    str = str & "<tr><td align=""center"" style=""font-weight:bold;font-size:14pt; margin-top: 10px"">MEDICAL LABORATORY REPORT</td></tr>"
    str = str & "<tr><td align=""center""><hr color=""#999999"" size=""1""></td></tr>"
    str = str & "<tr><td align=""center"">"
    str = str & "<table border=""0"" width=""650"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial"">"
    str = str & "<tr>"
    
    str = str & "<td name=""tdLabelInpLabReceiptID"" id=""tdLabelInpLabReceiptID"" style=""font-weight: bold"">LAB No.</td>"
    str = str & "<td width=""10""></td>"
    str = str & "<td style=""font-size:11pt"" name=""tdInputInpLabReceiptID"" id=""tdInputInpLabReceiptID""><b>" & GetRecordField("LabReQUESTID") & "</b></td>"
    
    str = str & "<td width=""10""></td>"
    
    str = str & "<td name=""tdLabelInpGenderID"" id=""tdLabelInpGenderID"" style=""font-weight: bold"">Gender</td>"
    str = str & "<td width=""10""></td>"
    
    If VisitNoExempt(GetRecordField("visitationid")) Or VisitNoExempt(Trim(GetRecordField("Patientid"))) Then
        str = str & "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & GetRequestGender(GetRecordField("LabRequestID")) & "</td>"
    Else
        str = str & "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & GetRecordField("GenderName") & "</td>"
    End If
    
    str = str & "<td width=""10""></td>"
    str = str & "</tr>"
    str = str & "<td name=""tdLabelInpLabReceiptName"" id=""tdLabelInpLabReceiptName"" style=""font-weight: bold"">Patient Name</td>"
    str = str & "<td width=""10""></td>"
    
    If VisitNoExempt(GetRecordField("visitationid")) Or VisitNoExempt(Trim(GetRecordField("Patientid"))) Then
        str = str & "<td style=""font-size:11pt"" name=""tdInputInpLabReceiptName"" id=""tdInputInpLabReceiptName""><b>" & GetRequestName(GetRecordField("LabRequestID")) & "</b></td>"
    Else
        str = str & "<td style=""font-size:11pt"" name=""tdInputInpLabReceiptName"" id=""tdInputInpLabReceiptName""><b>" & GetRecordField("pATIENTName") & "</b></td>"
    End If
    
    
    str = str & "<td width=""10""></td>"
    str = str & "<td name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"" style=""font-weight: bold"">Receipt Type</td>"
    str = str & "<td width=""10""></td>"
    str = str & "<td name=""tdInputInpReceiptTypeID"" id=""tdInputInpReceiptTypeID""></td>" ' & GetRecordField("ReceiptTypeName") & "</td>"
    str = str & "<td width=""10""></td>"
    str = str & "</tr>"
    
    str = str & "<tr>"
    
    str = str & "<tr>"
    str = str & "<td name=""tdLabelInpAge"" id=""tdLabelInpAge"" style=""font-weight: bold"">Age</td>"
    str = str & "<td width=""10""></td>"
    
    If VisitNoExempt(GetRecordField("visitationid")) Or VisitNoExempt(Trim(GetRecordField("Patientid"))) Then
        str = str & "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & GetRequestAge(GetRecordField("LabRequestID")) & "</td>"
    Else
        str = str & "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & GetPatientAge2(GetRecordField("VisitationID")) & "</td>"
    End If
    
    str = str & "<td width=""10""></td>"
    
    str = str & "<td name=""tdLabelInpInsuranceSchemeID"" id=""tdLabelInpInsuranceSchemeID"" style=""font-weight: bold"">Organization</td>"
    str = str & "<td width=""10""></td>"
    str = str & "<td name=""tdInputInpInsuranceSchemeID"" id=""tdInputInpInsuranceSchemeID""></td>" ' & GetRecordField("InsuranceSchemeName") & "</td>"
    str = str & "<td width=""10""></td>"
    str = str & "</tr>"
    
    str = str & "<tr>"
    str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Episode No.</td>"
    str = str & "<td width=""10""></td>"
    str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (GetRecordField("VisitationID")) & "</td>"
    str = str & "<td width=""10""></td>"
    str = str & "<td name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"" style=""font-weight: bold"">Patient Tel:</td>"
    str = str & "<td width=""10""></td>"
    str = str & "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & (GetRecordField("ContactNo")) & "</td>"
    str = str & "<td width=""10""></td>"
    str = str & "</tr>"
    
    str = str & "<tr>"
    str = str & "<td name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"" style=""font-weight: bold"">Manual Path. No</td>"
    str = str & "<td width=""10""></td>"
    str = str & "<td name=""tdInputInpReceiptTypeID"" id=""tdInputInpReceiptTypeID"">" & GetRecordField("ReceiptInfo2") & "</td>"
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
    str = str & "<table border=""0"" width=""650"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial"">"
    str = str & "<tr>"
    str = str & "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Requested By</td>"
    str = str & "<td width=""10""></td>"
    str = str & "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">" & GetRecordField("DoctorName") & "</td>"
    str = str & "<td width=""10""></td>"
    str = str & "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Sample Collection Date</td>"
    str = str & "<td width=""10""></td>"
    str = str & "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & FormatDateDetail(GetRecordField("RequestDate")) & "</td>"
    str = str & "<td width=""10""></td></tr>"
    
    str = str & "<tr>"
    str = str & "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Requested From</td>"
    str = str & "<td width=""10""></td>"
    str = str & "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">" & GetRecordField("HospitalName") & "</td>"
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
    str = str & "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">SSNIT:" & GetRecordField("BranchName") & "</td>"
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
    str = str & "<table border=""0"" width=""650"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial"">"
    str = str & "<tr><td><b>REQUESTED:</b>&nbsp;&nbsp;</td><td>" & reqTest & "</td></tr></table>"
    str = str & "</td></tr>"
    response.write str
End Sub
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
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
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
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
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
    Set rst = server.CreateObject("ADODB.Recordset")
    With rst
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
            ot = .fields("VisitInfo6")
            If Len(ot) > 0 Then
                'remove days when not needed
                tmp = Split(ot, " ")
                If UBound(tmp) >= 1 Then
                    ot = tmp(0) & " " & tmp(1)
                End If
                
            End If
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
    Set rst = server.CreateObject("ADODB.Recordset")
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
    Set rst = server.CreateObject("ADODB.Recordset")
    With rst
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
            ot = GetComboName("RequestGender", .fields("RequestGenderid"))
        End If
        .Close
        GetRequestGender = ot
    End With
End Function

Sub DisplayFoot()
    Dim str
    str = ""
    str = str & "<tr>"
    str = str & " <td valign=""bottom"" align=""center"" height=""10"">"
    str = str & "<table height=""10"" id=""tblFooter"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">"
    str = str & "<tr><td width=""110""></td>"
    str = str & "<td colspan=""6"" bgcolor=""#FFFFFF"" height=""10"" style=""font-size: 8pt"" align=""right"">"
    str = str & "IMaH Laboratory.Copyright@2013<br>" 'Software by : Spagad Technologies : 020-9426112<br> "
    str = str & "</td>"
    str = str & "</tr>"
    str = str & "</table>"
    str = str & "</td>"
    str = str & "</tr>"
    response.write str
End Sub

'DisplayLabResults
Sub DisplayLabResults()
    Dim rst, sql, rec, lnPerPg, i, j, num, num2, fnd, rst2
    Dim lnLbt, lbt, tct, str, cmp, sTy, out, lth, colm, colCnt, tdAttr
    Dim arrLbt(27, 5), arrPg(27), numLbt, numPg, pos, pdCnt, cPdPos
    Dim blkVl, blkCmp, blkCnt
    Dim cval
    Dim grf, yGrf, xGrf, hasGrf
    Dim arrGrf(100, 10)
    hasGrf = False
    yGrf = 0
    xGrf = 0
    
    lnPerPg = 27
    numLbt = 0
    numPg = 0
    lth = 50
    colCnt = 6
    'Init array
    For i = 1 To lnPerPg
        For j = 1 To 5
            arrLbt(i, j) = ""
            arrPg(i) = 0
        Next
    Next
    Set rst = server.CreateObject("ADODB.Recordset")
    Set rst2 = server.CreateObject("ADODB.Recordset")
    rec = Trim(Request.QueryString("LABREQUESTID"))
    lbt = Trim(Request.QueryString("LabTestID"))
    If Len(rec) > 0 Then
        sql = "SELECT testcategoryid,labtestid from investigation"
        sql = sql & " where LABREQUESTid='" & rec & "' and labtestid='" & lbt & "' and requeststatusid='RRD002' "
        sql = sql & " group by testcategoryid,labtestid order by testcategoryid,labtestid"
        With rst
            .open sql, conn, 3, 4
            If .recordCount > 0 Then
                .MoveFirst
                reqTest = GetRequestedTest(rec)
                Do While Not .EOF
                    lbt = .fields("labtestid")
                    numLbt = numLbt + 1
                    arrLbt(numLbt, 1) = .fields("testcategoryid")
                    arrLbt(numLbt, 2) = .fields("labtestid")
                    arrLbt(numLbt, 3) = "0"
            
                    sql = "SELECT count(labtestid) as cnt from labresults"
                    sql = sql & " where labrequestid='" & rec & "' and labtestid='" & lbt & "'"
                    With rst2
                        .open sql, conn, 3, 4
                        If .recordCount > 0 Then
                            .MoveFirst
                            If IsNumeric(.fields("cnt")) Then
                                arrLbt(numLbt, 3) = CStr(.fields("cnt"))
                            End If
                        End If
                        .Close
                    End With
                    
                    .MoveNext
                Loop
            End If
            .Close
        End With
        'Place Lab Test on pages
        For num = 1 To numLbt
            fnd = False
            lnLbt = CInt(arrLbt(num, 3)) + 2
            If numPg = 0 Then
                numPg = numPg + 1
                arrPg(numPg) = lnLbt
                arrLbt(num, 4) = CStr(numPg)
                fnd = True
            Else
                For num2 = 1 To numPg
                    If (lnPerPg >= (arrPg(num2) + lnLbt)) Then
                        arrPg(num2) = arrPg(num2) + lnLbt
                        arrLbt(num, 4) = CStr(num2)
                        fnd = True
                        Exit For
                    End If
                Next
            End If
            If Not fnd Then
                numPg = numPg + 1
                arrPg(numPg) = lnLbt
                arrLbt(num, 4) = CStr(numPg)
                fnd = True
            End If
        Next
        'Display Pages
        response.write "<tr><td width=""100%"" bgcolor=""white"" align=""right"">"
        response.write "<p align=""center"">"
        response.write "</td></tr>"
        For num = 1 To numPg
            DisplayHead
            response.write "<tr><td valign=""top"" align=""center"" height=""81"" width=""100%""><table width=""100%"" border=""0"" cellpadding=""0"" cellSpacing=""0""><tbody>"
            For num2 = 1 To numLbt
                If CInt(arrLbt(num2, 4)) = num Then
                    tct = arrLbt(num2, 1)
                    lbt = arrLbt(num2, 2)
    
                    'Block C/S sensitivity when no isolate
                    blkCmp = False
                    blkCnt = 0
                    
                    sql = "select column1,column2,column3,column4,column5,column6,testcomponentid from labresults "
                    sql = sql & " where LABREQUESTid='" & rec & "' and labtestid='" & lbt & "' order by comppos"
                    str = ""
                    With rst
                        .open sql, conn, 3, 4
                        If .recordCount > 0 Then
                            .MoveFirst
                            
                            sTy = "font-size:10pt; border-collapse:collapse"
                            response.write "<tr><td>"
                            response.write "<table border=""0"" style=""" & sTy & """ valign=""top"" width=""100%"" cellpadding=""2"" cellSpacing=""0"">"
                            response.write "<tr>"
                            response.write "<td height=""10"" colspan=""7"" align=""center""></td>"
                            response.write "</tr>"
                            response.write "<tr>"
                            sTy = "font-weight:bold; font-size:10pt; border-bottom-style: solid; border-bottom-width: 1px"
                            response.write "<td colspan=""7"" align=""left"" style=""" & sTy & """>" & GetComboName("labtest", lbt) & " -> [" & GetComboName("TestCategory", tct) & "]</td>"
                            response.write "</tr>"
                    
                            hasGrf = False
                            yGrf = 0
                            xGrf = 0
                            '///////////GRAPH - 27 Mar 2018 //////////
                            If UCase(lbt) = "LL0001" Then
                              hasGrf = True
                              For yGrf = 1 To 99
                                For xGrf = 1 To 9
                                  arrGrf(yGrf, xGrf) = ""
                                Next
                              Next
                              yGrf = 0
                              xGrf = 0
                              response.write "<tr><td colspan=""7"">"
                              grf = SetupUserGraph(500, 900, "Line", "Vertical", "Yes", "Yes", "")
                              response.write "</td></tr>"
                            End If
                            '/////////END GRAPH/////////
                        
                            cPdPos = 0
                            Do While Not .EOF
                                cmp = .fields("testcomponentid")
                                out = GetComboName("Testcomponent", cmp)
                                
                                yGrf = yGrf + 1
                                xGrf = 1
                                arrGrf(yGrf, xGrf) = cmp
                                
                                If Len(out) < lth Then
                                    out = Replace(out, " ", " ")
                                End If
                                sTy = ""
                                pos = InStr(1, UCase(out), "PARAMETER")
                                If (pos > 0) Or (UCase(cmp) = "L0516") Then
                                  cPdPos = 0
                                  sTy = "style=""font-weight:bold; font-size:10pt; border-bottom-style: solid; border-bottom-width: 1px"""
                                Else
                                  cPdPos = cPdPos + 1
                                  If cPdPos = 1 Then
                                    sTy = "style=""padding-top: 3px"""
                                  End If
                                End If
                                'Block C/S sensitivity when no isolate
                                blkCmp = False
                                If blkCnt > 0 Then
                                  blkCnt = blkCnt - 1
                                  blkCmp = True
                                End If
                                If (UCase(cmp) = "L0402") Or (UCase(cmp) = "L0410") Then
                                  blkVl = Trim(.fields("column1"))
                                  If UCase(blkVl) = "" Then
                                    blkCmp = True
                                    blkCnt = 6
                                  ElseIf (UCase(blkVl) = "12") Or (UCase(blkVl) = "71") Or (UCase(blkVl) = "72") Or (UCase(blkVl) = "75") Or (UCase(blkVl) = "77") Or (UCase(blkVl) = "80") Then
                                    blkCnt = 6
                                  ElseIf (UCase(blkVl) = "T006") Then
                                    blkCnt = 6
                                  End If
                              End If
                              If Not blkCmp Then
                              
                                If Utils.Contains(.fields("TestComponentID").value, "_interpretation") Then
                                    response.write "<tr>"
                                        response.write "<td style='text-transform:uppercase;font-weight:bold;text-decoration:underline;'>RESULTS INTERPRETATION</td>"
                                    response.write "</tr>"
                                    response.write "<tr>"
                                        cval = GetCompFldVal("TestComponentID", cmp, "Column1", .fields("column1"))
                                        response.write "<td colspan='6' style='text-align: justify;padding-left:20px;'>" & Replace(cval, vbCrLf, "<br/>") & "</td>"
                                    response.write "</tr>"
                                Else
                                
                                    response.write "<tr>"
                                    response.write "<td valign=""top""" & sTy & ">" & out & "</td>"
                                    For colm = 1 To colCnt
                                        tdColSp = "1"
                                        tdAlign = ""
                                        tdAttr = " valign=""top"""
                                        out = GetCompFldVal("TestComponentID", cmp, "Column" & CStr(colm), .fields("column" & CStr(colm)))
                                        
                                        xGrf = xGrf + 1
                                        arrGrf(yGrf, xGrf) = out
                                        
                                        If Len(tdAlign) > 0 Then
                                          tdAttr = tdAttr & " align=""" & tdAlign & """ "
                                        Else
                                             If Len(Trim(out)) >= 3 Then
                                              If UCase(Left(Trim(out), 3)) = "<B>" Then
                                                 tdAttr = tdAttr & " align=""left"" "
                                                Else
                                                  tdAttr = tdAttr & " align=""center"" "
                                              End If
                                             Else
                                                tdAttr = tdAttr & " align=""center"" "
                                             End If
                                        End If
                                        If Len(out) < lth Then
                                          out = Replace(out, " ", " ")
                                        End If
                                        If Len(Trim(out)) = 0 Then
                                           out = " "
                                        End If
                                        If CInt(tdColSp) = 1 Then
                                          response.write "<td " & sTy & " " & tdAttr & ">" & out & "</td>"
                                        Else
                                            response.write "<td " & sTy & " " & tdAttr & " colspan=""" & tdColSp & """>" & out & "</td>"
                                            colm = colm + CInt(tdColSp) - 1
                                        End If
                                    Next
                                    response.write "</tr>"
                                End If 'Utils.Contains
                              End If 'blkCmp
                             .MoveNext
                            Loop
                            response.write "</table>"
                            response.write "</td></tr>"
                            'response.write str
                            If hasGrf Then
                                ProcessGraphCoord lbt, grf, arrGrf, yGrf, xGrf
                            End If
                        End If
                        .Close
                    End With
                End If
            Next
                            
                                    
        'Fill Blank empty space
        'For pdCnt = arrPg(num) To lnPerPg
            ' Response.Write "<tr>"
            ' Response.Write "<td colspan=""7"" align=""center"">&nbsp;</td>"
            ' Response.Write "</tr>"
        'Next
        'If num = numPg Then
        '  Response.Write "<tr>"
        '  Response.Write "<td colspan=""7"" align=""right""><br>Electronically Signed By:  " & GetComboName("LabTech", lstLbTech) & "  <br> </td>"
        '  Response.Write "</tr>"
        'End If
        'Response.Write "<tr>"
        'Response.Write "<td colspan=""7"" align=""center""><br><u>Page " & CStr(num) & " of " & CStr(numPg) & "</u></td>"
        'Response.Write "</tr>"
        
        '' -- add footer  -- ''
        response.write "</tbody><tfoot>"
            response.write "<tr><td><hr color='#999999' size='1'/></td></tr>"
            response.write "<tr><td style='text-align: right; '>Electronically Signed By: " & GetComboName("LabTech", lstLbTech) & "</td></tr>"
        response.write "</tfoot>"
        response.write "</table></td></tr>"
        If num = numPg Then
            'DisplayFoot
        End If
        Next
    End If
    Set rst = Nothing
    Set rst2 = Nothing
End Sub

Sub ProcessGraphCoord(lbt, grf, arrGrf, yGrf, xGrf)
  Dim x, y, grfCat, pos, coordNm, coordVl
  Select Case UCase(lbt)
    Case "LL0001" 'GTT
      
      For y = 1 To yGrf
        If y = 1 Then 'Header
          pos = 3
          grfCat = AddGraphCategory(grf, "NORMAL", "Yes") 'Normal
          arrGrf(yGrf + 1, pos) = grfCat 'Store CatID in next row after last
          pos = 4
          grfCat = AddGraphCategory(grf, "PREDIABETES", "Yes") 'PreDiabetes
          arrGrf(yGrf + 1, pos) = grfCat 'Store CatID in next row after last
          pos = 5
          grfCat = AddGraphCategory(grf, "DIABETES", "Yes") 'Diabetes
          arrGrf(yGrf + 1, pos) = grfCat 'Store CatID in next row after last
          pos = 6
          grfCat = AddGraphCategory(grf, "PATIENT", "Yes") 'Patient
          arrGrf(yGrf + 1, pos) = grfCat 'Store CatID in next row after last
        Else
          For x = 3 To 6
            grfCat = Trim(arrGrf(yGrf + 1, x))
            If Len(grfCat) > 0 Then
              coordNm = arrGrf(y, 2) 'First Column of current row,y
              coordVl = Trim(arrGrf(y, x))
              If Not IsNumeric(coordVl) Then
                coordVl = "0"
              End If
              AddGraphCoord grf, grfCat, coordNm, coordVl
            End If
          Next
        End If
      Next

'       grfCat = AddGraphCategory(grf, "NORMAL2", "Yes")
'       AddGraphCoord grf, grfCat, "FBS", 5
'       AddGraphCoord grf, grfCat, "1", 6
'       AddGraphCoord grf, grfCat, "2", 7
'       AddGraphCoord grf, grfCat, "3", 8
'       AddGraphCoord grf, grfCat, "4", 9
'       grfCat = AddGraphCategory(grf, "PREDIABETES2", "Yes")
'       AddGraphCoord grf, grfCat, "FBS", 3
'       AddGraphCoord grf, grfCat, "1", 7
'       AddGraphCoord grf, grfCat, "2", 9
'       AddGraphCoord grf, grfCat, "3", 9
'       AddGraphCoord grf, grfCat, "4", 7
'       response.write "</td></tr>"
'        grfCat = AddGraphCategory(grf, "DIABETES2", "Yes")
'       AddGraphCoord grf, grfCat, "FBS", 3
'       AddGraphCoord grf, grfCat, "1", 9
'       AddGraphCoord grf, grfCat, "2", 11
'       AddGraphCoord grf, grfCat, "3", 11
'       AddGraphCoord grf, grfCat, "4", 7
  End Select
End Sub
'DisplayLabResults2
Sub DisplayLabResults2()
    Dim rst, sql, rec, lnPerPg, i, j, num, num2, fnd, rst2
    Dim lnLbt, lbt, tct, str, cmp, sTy, out, lth, colm, colCnt, tdAttr
    Dim arrLbt(27, 5), arrPg(27), numLbt, numPg, pos, pdCnt, cPdPos
    Dim blkVl, blkCmp, blkCnt
    lnPerPg = 27
    numLbt = 0
    numPg = 0
    lth = 50
    colCnt = 6
    'Init array
    For i = 1 To lnPerPg
        For j = 1 To 5
            arrLbt(i, j) = ""
            arrPg(i) = 0
        Next
    Next
    Set rst = server.CreateObject("ADODB.Recordset")
    Set rst2 = server.CreateObject("ADODB.Recordset")
    rec = Trim(Request.QueryString("LABREQUESTID"))
    lbt = Trim(Request.QueryString("LabTestID"))
    If Len(rec) > 0 Then
        sql = "SELECT testcategoryid,labtestid from investigation2"
        sql = sql & " where LABREQUESTid='" & rec & "' and labtestid='" & lbt & "'" ' and requeststatusid='RRD002' "
        sql = sql & " group by testcategoryid,labtestid order by testcategoryid,labtestid"
        With rst
            .open sql, conn, 3, 4
            If .recordCount > 0 Then
                .MoveFirst
                reqTest = GetRequestedTest(rec)
                Do While Not .EOF
                    lbt = .fields("labtestid")
                    numLbt = numLbt + 1
                    arrLbt(numLbt, 1) = .fields("testcategoryid")
                    arrLbt(numLbt, 2) = .fields("labtestid")
                    arrLbt(numLbt, 3) = "0"
                        
                    sql = "SELECT count(labtestid) as cnt from labresults"
                    sql = sql & " where labrequestid='" & rec & "' and labtestid='" & lbt & "'"
                    With rst2
                        .open sql, conn, 3, 4
                        If .recordCount > 0 Then
                            .MoveFirst
                            If IsNumeric(.fields("cnt")) Then
                                arrLbt(numLbt, 3) = CStr(.fields("cnt"))
                            End If
                        End If
                        .Close
                    End With
                    
                    .MoveNext
                Loop
            End If
                .Close
        End With
        'Place Lab Test on pages
        For num = 1 To numLbt
            fnd = False
            lnLbt = CInt(arrLbt(num, 3)) + 2
            If numPg = 0 Then
                numPg = numPg + 1
                arrPg(numPg) = lnLbt
                arrLbt(num, 4) = CStr(numPg)
                fnd = True
            Else
                For num2 = 1 To numPg
                    If (lnPerPg >= (arrPg(num2) + lnLbt)) Then
                        arrPg(num2) = arrPg(num2) + lnLbt
                        arrLbt(num, 4) = CStr(num2)
                        fnd = True
                        Exit For
                    End If
                Next
            End If
            If Not fnd Then
                numPg = numPg + 1
                arrPg(numPg) = lnLbt
                arrLbt(num, 4) = CStr(numPg)
                fnd = True
            End If
        Next
        'Display Pages
        response.write "<tr><td width=""100%"" bgcolor=""white"" align=""right"">"
        response.write "<p align=""center"">"
        response.write "</td></tr>"
        For num = 1 To numPg
            DisplayHead
            response.write "<tr><td valign=""top"" align=""center"" height=""81"" width=""100%""><table width=""100%"" border=""0"" cellpadding=""0"" cellSpacing=""0""></tbody>"
            For num2 = 1 To numLbt
                If CInt(arrLbt(num2, 4)) = num Then
                    tct = arrLbt(num2, 1)
                    lbt = arrLbt(num2, 2)
                    
                    'Block C/S sensitivity when no isolate
                    blkCmp = False
                    blkCnt = 0
                    
                    sql = "select column1,column2,column3,column4,column5,column6,testcomponentid from labresults "
                    sql = sql & " where LABREQUESTid='" & rec & "' and labtestid='" & lbt & "' order by comppos"
                    str = ""
                    With rst
                       .open sql, conn, 3, 4
                        If .recordCount > 0 Then
                            .MoveFirst
                            sTy = "font-size:10pt; border-collapse:collapse"
                            str = str & "<tr><td>"
                            str = str & "<table border=""0"" style=""" & sTy & """ valign=""top"" width=""100%"" cellpadding=""2"" cellSpacing=""0"">"
                            str = str & "<tr>"
                            str = str & "<td height=""10"" colspan=""7"" align=""center""></td>"
                            str = str & "</tr>"
                            str = str & "<tr>"
                            sTy = "font-weight:bold; font-size:10pt; border-bottom-style: solid; border-bottom-width: 1px"
                            str = str & "<td colspan=""7"" align=""left"" style=""" & sTy & """>" & GetComboName("labtest", lbt) & " -> [" & GetComboName("TestCategory", tct) & "]</td>"
                            str = str & "</tr>"
                            cPdPos = 0
                            Do While Not .EOF
                                cmp = .fields("testcomponentid")
                                out = GetComboName("Testcomponent", cmp)
                                If Len(out) < lth Then
                                    out = Replace(out, " ", " ")
                                End If
                                sTy = ""
                                pos = InStr(1, UCase(out), "PARAMETER")
                                If (pos > 0) Or (UCase(cmp) = "L0516") Then
                                  cPdPos = 0
                                  sTy = "style=""font-weight:bold; font-size:10pt; border-bottom-style: solid; border-bottom-width: 1px"""
                                Else
                                  cPdPos = cPdPos + 1
                                  If cPdPos = 1 Then
                                    sTy = "style=""padding-top: 3px"""
                                  End If
                                End If
                                'Block C/S sensitivity when no isolate
                                blkCmp = False
                                If blkCnt > 0 Then
                                  blkCnt = blkCnt - 1
                                  blkCmp = True
                                End If
                                If (UCase(cmp) = "L0402") Or (UCase(cmp) = "L0410") Then
                                  blkVl = Trim(.fields("column1"))
                                  If UCase(blkVl) = "" Then
                                    blkCmp = True
                                    blkCnt = 6
                                  ElseIf (UCase(blkVl) = "12") Or (UCase(blkVl) = "71") Or (UCase(blkVl) = "72") Or (UCase(blkVl) = "75") Or (UCase(blkVl) = "77") Or (UCase(blkVl) = "80") Then
                                    blkCnt = 6
                                  ElseIf (UCase(blkVl) = "T006") Then
                                    blkCnt = 6
                                  End If
                                End If
                                If Not blkCmp Then
                                    If Utils.Contains(.fields("TestComponentID").value, "_interpretation") And False Then
                                        str = str & "<tr>"
                                            str = str & "<td style='text-transform:uppercase;font-weight:bold;text-decoration:underline;'>RESULTS INTERPRETATION</td>"
                                        str = str & "</tr>"
                                        str = str & "<tr>"
                                            cval = GetCompFldVal("TestComponentID", cmp, "Column1", .fields("column1"))
                                            str = str & "<td colspan='6' style='text-align: justify;padding-left:20px;'>" & Replace(cval, vbCrLf, "<br/>") & "</td>"
                                        str = str & "</tr>"
                                    Else
                                      str = str & "<tr>"
                                      str = str & "<td valign=""top""" & sTy & ">" & out & "</td>"
                                      For colm = 1 To colCnt
                                        tdColSp = "1"
                                        tdAlign = ""
                                        tdAttr = " valign=""top"""
                                        out = GetCompFldVal("TestComponentID", cmp, "Column" & CStr(colm), .fields("column" & CStr(colm)))
                                        If Len(tdAlign) > 0 Then
                                            tdAttr = tdAttr & " align=""" & tdAlign & """ "
                                        Else
                                            If Len(Trim(out)) >= 3 Then
                                                If UCase(Left(Trim(out), 3)) = "<B>" Then
                                                    tdAttr = tdAttr & " align=""left"" "
                                                Else
                                                    tdAttr = tdAttr & " align=""center"" "
                                                End If
                                            Else
                                                tdAttr = tdAttr & " align=""center"" "
                                            End If
                                        End If
                                        If Len(out) < lth Then
                                            out = Replace(out, " ", " ")
                                        End If
                                        If Len(Trim(out)) = 0 Then
                                            out = " "
                                        End If
                                        If CInt(tdColSp) = 1 Then
                                          str = str & "<td " & sTy & " " & tdAttr & ">" & out & "</td>"
                                        Else
                                            str = str & "<td " & sTy & " " & tdAttr & " colspan=""" & tdColSp & """>" & out & "</td>"
                                            colm = colm + CInt(tdColSp) - 1
                                        End If
                                        Next
                                      str = str & "</tr>"
                                      End If
                                End If 'blkCmp
                                .MoveNext
                            Loop
                            str = str & "</table>"
                            str = str & "</td></tr>"
                            response.write str
                        End If
                        .Close
                    End With
               End If
            Next
                        
    
        'Fill Blank empty space
'           For pdCnt = arrPg(num) To lnPerPg
'                Response.Write "<tr>"
'                Response.Write "<td colspan=""7"" align=""center"">&nbsp;</td>"
'                Response.Write "</tr>"
'            Next
'            If num = numPg Then
'                Response.Write "<tr>"
'                Response.Write "<td colspan=""7"" align=""right""><br>Electronically Signed By:  " & GetComboName("LabTech", lstLbTech) & "  <br> </td>"
'                Response.Write "</tr>"
'            End If
            'Response.Write "<tr>"
            'Response.Write "<td colspan=""7"" align=""center""><br><u>Page " & CStr(num) & " of " & CStr(numPg) & "</u></td>"
            'Response.Write "</tr>"
'            Response.Write "<tfoot>"
'            Response.Write "</tfoot>"
            response.write "</tbody><tfoot>"
                response.write "<tr><td><hr color='#999999' size='1'/></td></tr>"
                response.write "<tr><td style='text-align: right; '>Electronically Signed By: " & GetComboName("LabTech", lstLbTech) & "</td></tr>"
            response.write "</tfoot>"
            response.write "</table></td></tr>"
            If num = numPg Then
                'DisplayFoot
            End If
        Next
    End If
    Set rst = Nothing
    Set rst2 = Nothing
    
End Sub

Function reportHeader()
    Dim str
    str = str & "<div class=""rpthead"">"
    str = str & " <img class=""rpthead"" src=""images/IMaH_Letterhead6.png"">"
    str = str & " </div>"
    reportHeader = str
End Function


Dim reqTest, lstLbTech, lstDate, cliDia
reqTest = ""
lstLbTech = ""
lstDate = CDate("1 Jan 2000")
cliDia = ""

response.write reportHeader
DisplayLabResults
DisplayLabResults2



'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
'Empty
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
