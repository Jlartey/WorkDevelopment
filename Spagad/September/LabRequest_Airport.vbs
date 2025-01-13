'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim sql
Dim rstPrn1
Dim rstPrn2
Dim cat
Dim catNm
Dim tot
Dim tDur
Dim tAmt
Dim rTyp
Dim recNo
Dim manNo
Dim sql2
Dim hasHdr
Set rstPrn1 = server.CreateObject("ADODB.Recordset")
Set rstPrn2 = server.CreateObject("ADODB.Recordset")
hasHdr = False
'Patient Copy
sql = "select distinct testcategoryid from investigation where labrequestid='" & Trim(request.querystring("labrequestid")) & "'"

With rstPrn1
    .open qryPro.FltQry(sql), conn, 3, 4

    If .RecordCount > 0 Then
        .movefirst
        'Do While Not .EOF
        cat = .fields("testcategoryid")
        tot = 0
        catNm = UCase(GetComboName("testcategory", cat))
        sql = GetTableSql("investigation")
        sql = sql & " and  investigation.labrequestid='" & Trim(request.querystring("labrequestid")) & "'"

        rTyp = Trim(GetRecordField("ReceiptTypeID"))
        recNo = Trim(GetRecordField("ReceiptInfo1"))
        manNo = Trim(GetRecordField("ReceiptInfo2"))

        If Len(recNo) > 0 Then
            recNo = "[" & recNo & "]"
        End If

        response.write "<tr>"
        response.write "<td align=""center"">"
        response.write "<table border=""0"" width=""600"" cellspacing=""0"" cellpadding=""0"">"
        response.write "<tr height=""60"">"
        response.write "<td align=""Center"" valign=""top""></td>"
        response.write "<td align=""center"" height=""20"" bgcolor=""white"" style=""font-size: 14pt; color: #CCCCFF;  font-family:Arial"">"
        AddReportHeader
        response.write "</td>"
        response.write "</tr>"
        response.write "</table>"
        response.write "</td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:10pt"">"
        response.write "RESULTS COLLECTION SLIP --- Collection Date:<u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center"">"
        response.write "<table border=""0"" width=""600"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial"">"
        response.write "<caption>"
        response.write "<tr>"
        response.write "<td name=""tdLabelInplabrequestID"" id=""tdLabelInplabrequestID"" style=""font-weight: bold"">Path. No.</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInplabrequestID"" id=""tdInputInplabrequestID"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"" >" & (GetRecordField("labrequestID")) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInplabrequestName"" id=""tdLabelInplabrequestName"" style=""font-weight: bold"">Patient Name</td>"
        response.write "<td width=""10""></td>"

        If VisitNoExempt(GetRecordField("visitationid")) Then
            response.write "<td name=""tdInputInplabrequestName"" id=""tdInputInplabrequestName"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"">" & (GetRecordField("RequestName")) & "</td>"
        Else
            response.write "<td name=""tdInputInplabrequestName"" id=""tdInputInplabrequestName"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"">" & (GetRecordField("PatientName")) & "</td>"
        End If

        response.write "<td width=""10""></td>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpAge"" id=""tdLabelInpAge"" style=""font-weight: bold"">Age</td>"
        response.write "<td width=""10""></td>"

        If VisitNoExempt(GetRecordField("visitationid")) Then
            response.write "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & (GetRecordField("RequestAge")) & "</td>"
        Else
            response.write "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & (GetPatientAge(GetRecordField("visitationid"))) & "</td>"
        End If

        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInpGenderID"" id=""tdLabelInpGenderID"" style=""font-weight: bold"">Gender</td>"
        response.write "<td width=""10""></td>"

        If VisitNoExempt(GetRecordField("visitationid")) Then
            response.write "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & (GetRecordField("RequestGenderName")) & "</td>"
        Else
            response.write "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & (GetRecordField("GenderName")) & "</td>"
        End If

        response.write "<td width=""10""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"" style=""font-weight: bold"">Request Type</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpReceiptTypeID"" id=""tdInputInpReceiptTypeID"">&nbsp;" & recNo & "</td>" '& (GetRecordField("ReceiptTypeName")) & "&nbsp;" & recNo & "</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Organization</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">&nbsp;</td>" '& (GetRecordField("InsuranceSchemeName")) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Request Day</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (GetRecordField("WorkingDayName")) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"" style=""font-weight: bold"">Request Time</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & (FormatDateDetail(GetRecordField("RequestDate"))) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Doctor</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (GetRecordField("DoctorName")) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"" style=""font-weight: bold"">Patient Tel:</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & (GetRecordField("ContactNo")) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpSystemUserID"" id=""tdLabelInpSystemUserID"" style=""font-weight: bold"">Lab Staff</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpSystemUserID"" id=""tdInputInpSystemUserID"">" & (GetRecordField("SystemUserName")) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Visit No.</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">" & (UCase(GetRecordField("VisitationID"))) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"" style=""font-weight: bold"">Manual Path. No</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpReceiptTypeID"" id=""tdInputInpReceiptTypeID"">" & GetRecordField("ReceiptInfo2") & "</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Requested From</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">" & GetRecordField("HospitalName") & "</td>"
        response.write "<td width=""10""></td>"
        response.write "</tr></caption>"
        response.write "</table>        </td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center"">"
        response.write "<table id=""tblMultiSelect"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""border-collapse: collapse"" bordercolor=""#111111"">"
        response.write "<tr>"
        response.write "<td align=""left"" colspan=""2"">"
        response.write "<table height=""10"" id=""tblMultiSelectSection"" cellspacing=""0"" cellpadding=""0"" border=""0"" style=""border-collapse: collapse; font-size:10pt"" bordercolor=""#999999"" width=""100%"">"
        response.write "<tr>"
        response.write "<td valign=""top"" width=""100%"">"
        response.write "<table style=""font-size: 9pt; font-family: Arial"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""2"">"
        hasHdr = True

        With rstPrn2
            .open qryPro.FltQry(sql), conn, 3, 4

            If .RecordCount > 0 Then
                .movefirst
                If UCase(rTyp) = "R001" Then 'CASH
                response.write "<tr><td><b>SERVICE    DESCRIPTION</b></td></tr>" '<td align=""right""><b>DUR.[TAT]</b></td><td align=""right""></td></tr>" '<b>AMOUNT</b></td></tr>"
                response.write "<tr><td colspan=""3"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"

                Do While Not .EOF
                    response.write "<tr><td>" & (.fields("Labtestname")) & "</td>"
                    tDur = ""
                    tAmt = .fields("TestAmt1")

                    If tAmt = 1 Then
                        tDur = FormatNumber(tAmt, 0, , , -1) & "&nbsp;Hr"
                    ElseIf tAmt < 24 Then
                        tDur = FormatNumber(tAmt, 0, , , -1) & "&nbsp;Hrs"
                    ElseIf tAmt = 24 Then
                        tDur = FormatNumber((tAmt / 24), 0, , , -1) & "&nbsp;Day"
                    ElseIf tAmt > 24 Then
                        tDur = FormatNumber((tAmt / 24), 1, , , -1) & "&nbsp;Day(s)"
                    End If

                    tDur = "&nbsp;"
                    response.write "<td align=""right"">" & tDur & "</td>"
                    response.write "<td align=""right"">&nbsp;</td></tr>" '& (FormatNumber(.Fields("finalamt"), 2, , , -1)) & "</td></tr>"
                    tot = tot + .fields("finalamt")
                    .MoveNext
                Loop

                response.write "<tr><td colspan=""3"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"
                '    response.write "<tr><td colspan=""2""><b>TOTAL AMOUNT</b></td>"
                '    response.write "<td align=""right""><b>" & (FormatNumber(tot, 2, , , -1)) & "</b></td></tr>"
                '    response.write "<tr><td colspan=""3"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"
            Else 'CREDIT
                response.write "<tr><td><b>SERVICE    DESCRIPTION</b></td></tr>" '<td align=""right""><b>DUR.[TAT]</b></td><td align=""right""><b>&nbsp;</b></td></tr>"
                response.write "<tr><td colspan=""3"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"

                Do While Not .EOF
                    response.write "<tr><td>" & (.fields("Labtestname")) & "</td>"
                    tDur = ""
                    tAmt = .fields("TestAmt1")

                    If tAmt = 1 Then
                        tDur = FormatNumber(tAmt, 0, , , -1) & "&nbsp;Hr"
                    ElseIf tAmt < 24 Then
                        tDur = FormatNumber(tAmt, 0, , , -1) & "&nbsp;Hrs"
                    ElseIf tAmt = 24 Then
                        tDur = FormatNumber((tAmt / 24), 0, , , -1) & "&nbsp;Day"
                    ElseIf tAmt > 24 Then
                        tDur = FormatNumber((tAmt / 24), 1, , , -1) & "&nbsp;Day(s)"
                    End If

                    tDur = "&nbsp;"
                    response.write "<td align=""right"">" & tDur & "</td>"
                    response.write "<td align=""right"">&nbsp;</td></tr>"
                    tot = tot + .fields("finalamt")
                    .MoveNext
                Loop

                response.write "<tr><td colspan=""3"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"
                'response.write "<tr><td colspan=""2""><b>TOTAL AMOUNT</b></td>"
                'response.write "<td align=""right""><b>" & (FormatNumber(tot, 2, , , -1)) & "</b></td></tr>"
                'response.write "<tr><td colspan=""3"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"
            End If

        End If

        .Close
    End With

    .MoveNext
    'Loop
End If

.Close
End With

'///////////////////////Investigation2//////////////////
'Patient Copy
sql = "select distinct testcategoryid from investigation2 where labrequestid='" & Trim(request.querystring("labrequestid")) & "'"

With rstPrn1
.open qryPro.FltQry(sql), conn, 3, 4

If .RecordCount > 0 Then
    .movefirst
    'Do While Not .EOF
    cat = .fields("testcategoryid")
    tot = 0
    catNm = UCase(GetComboName("testcategory", cat))
    sql = GetTableSql("investigation2")
    sql = sql & " and  investigation2.labrequestid='" & Trim(request.querystring("labrequestid")) & "'"

    rTyp = Trim(GetRecordField("ReceiptTypeID"))
    recNo = Trim(GetRecordField("ReceiptInfo1"))
    manNo = Trim(GetRecordField("ReceiptInfo2"))

    If Len(recNo) > 0 Then
        recNo = "[" & recNo & "]"
    End If

    If Not hasHdr Then
        response.write "<tr>"
        response.write "<td align=""center"">"
        response.write "<table border=""0"" width=""600"" cellspacing=""0"" cellpadding=""0"">"
        response.write "<tr height=""60"">"
        response.write "<td align=""Center"" valign=""top""></td>"
        response.write "<td align=""center"" height=""20"" bgcolor=""white"" style=""font-size: 14pt; color: #CCCCFF;  font-family:Arial"">"
        AddReportHeader
        response.write "</td>"
        response.write "</tr>"
        response.write "</table>"
        response.write "</td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:10pt"">"
        response.write "RESULTS COLLECTION SLIP --- Collection Date:<u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center"">"
        response.write "<table border=""0"" width=""600"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial"">"
        response.write "<caption>"
        response.write "<tr>"
        response.write "<td name=""tdLabelInplabrequestID"" id=""tdLabelInplabrequestID"" style=""font-weight: bold"">Path. No.</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInplabrequestID"" id=""tdInputInplabrequestID"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"" >" & (GetRecordField("labrequestID")) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInplabrequestName"" id=""tdLabelInplabrequestName"" style=""font-weight: bold"">Patient Name</td>"
        response.write "<td width=""10""></td>"

        If VisitNoExempt(GetRecordField("visitationid")) Then
            response.write "<td name=""tdInputInplabrequestName"" id=""tdInputInplabrequestName"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"">" & (GetRecordField("RequestName")) & "</td>"
        Else
            response.write "<td name=""tdInputInplabrequestName"" id=""tdInputInplabrequestName"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"">" & (GetRecordField("PatientName")) & "</td>"
        End If

        response.write "<td width=""10""></td>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpAge"" id=""tdLabelInpAge"" style=""font-weight: bold"">Age</td>"
        response.write "<td width=""10""></td>"

        If VisitNoExempt(GetRecordField("visitationid")) Then
            response.write "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & (GetRecordField("RequestAge")) & "</td>"
        Else
            response.write "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & (GetPatientAge(GetRecordField("visitationid"))) & "</td>"
        End If

        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInpGenderID"" id=""tdLabelInpGenderID"" style=""font-weight: bold"">Gender</td>"
        response.write "<td width=""10""></td>"

        If VisitNoExempt(GetRecordField("visitationid")) Then
            response.write "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & (GetRecordField("RequestGenderName")) & "</td>"
        Else
            response.write "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & (GetRecordField("GenderName")) & "</td>"
        End If

        response.write "<td width=""10""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"" style=""font-weight: bold"">Request Type</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpReceiptTypeID"" id=""tdInputInpReceiptTypeID"">&nbsp;" & recNo & "</td>" '& (GetRecordField("ReceiptTypeName")) & "&nbsp;" & recNo & "</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Organization</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">&nbsp;</td>" '& (GetRecordField("InsuranceSchemeName")) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Request Day</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (GetRecordField("WorkingDayName")) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"" style=""font-weight: bold"">Request Time</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & (FormatDateDetail(GetRecordField("RequestDate"))) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Doctor</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (GetRecordField("DoctorName")) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"" style=""font-weight: bold"">Patient Tel:</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & (GetRecordField("ContactNo")) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpSystemUserID"" id=""tdLabelInpSystemUserID"" style=""font-weight: bold"">Lab Staff</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpSystemUserID"" id=""tdInputInpSystemUserID"">" & (GetRecordField("SystemUserName")) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Visit No.</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">" & (UCase(GetRecordField("VisitationID"))) & "</td>"
        response.write "<td width=""10""></td>"
        response.write "</tr>"

        response.write "<tr>"
        response.write "<td name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"" style=""font-weight: bold"">Manual Path. No</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpReceiptTypeID"" id=""tdInputInpReceiptTypeID"">" & GetRecordField("ReceiptInfo2") & "</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Requested From</td>"
        response.write "<td width=""10""></td>"
        response.write "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">" & GetRecordField("HospitalName") & "</td>"
        response.write "<td width=""10""></td>"
        response.write "</tr></caption>"
        response.write "</table>        </td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align=""center"">"
        response.write "<table id=""tblMultiSelect"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""border-collapse: collapse"" bordercolor=""#111111"">"
        response.write "<tr>"
        response.write "<td align=""left"" colspan=""2"">"
        response.write "<table height=""10"" id=""tblMultiSelectSection"" cellspacing=""0"" cellpadding=""0"" border=""0"" style=""border-collapse: collapse; font-size:10pt"" bordercolor=""#999999"" width=""100%"">"
        response.write "<tr>"
        response.write "<td valign=""top"" width=""100%"">"
        response.write "<table style=""font-size: 9pt; font-family: Arial"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""2"">"
    End If

    'Investigation2

    With rstPrn2
        .open qryPro.FltQry(sql), conn, 3, 4

        If .RecordCount > 0 Then
            .movefirst
            If UCase(rTyp) = "R001" Then 'CASH
            response.write "<tr><td><b>SERVICE    DESCRIPTION</b></td></tr>" '<td align=""right""><b>DUR.[TAT]</b></td><td align=""right""></td></tr>" '<b>AMOUNT</b></td></tr>"
            response.write "<tr><td colspan=""3"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"

            Do While Not .EOF
                response.write "<tr><td>" & (.fields("Labtestname")) & "</td>"
                tDur = ""
                tAmt = .fields("TestAmt1")

                If tAmt = 1 Then
                    tDur = FormatNumber(tAmt, 0, , , -1) & "&nbsp;Hr"
                ElseIf tAmt < 24 Then
                    tDur = FormatNumber(tAmt, 0, , , -1) & "&nbsp;Hrs"
                ElseIf tAmt = 24 Then
                    tDur = FormatNumber((tAmt / 24), 0, , , -1) & "&nbsp;Day"
                ElseIf tAmt > 24 Then
                    tDur = FormatNumber((tAmt / 24), 1, , , -1) & "&nbsp;Day(s)"
                End If

                tDur = "&nbsp;"
                response.write "<td align=""right"">" & tDur & "</td>"
                response.write "<td align=""right"">&nbsp;</td></tr>" '& (FormatNumber(.Fields("finalamt"), 2, , , -1)) & "</td></tr>"
                tot = tot + .fields("finalamt")
                .MoveNext
            Loop

            response.write "<tr><td colspan=""3"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"
            '    response.write "<tr><td colspan=""2""><b>TOTAL AMOUNT</b></td>"
            '    response.write "<td align=""right""><b>" & (FormatNumber(tot, 2, , , -1)) & "</b></td></tr>"
            '    response.write "<tr><td colspan=""3"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"
        Else 'CREDIT
            response.write "<tr><td><b>SERVICE    DESCRIPTION</b></td></tr>" '<td align=""right""><b>DUR.[TAT]</b></td><td align=""right""><b>&nbsp;</b></td></tr>"
            response.write "<tr><td colspan=""3"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"

            Do While Not .EOF
                response.write "<tr><td>" & (.fields("Labtestname")) & "</td>"
                tDur = ""
                tAmt = .fields("TestAmt1")

                If tAmt = 1 Then
                    tDur = FormatNumber(tAmt, 0, , , -1) & "&nbsp;Hr"
                ElseIf tAmt < 24 Then
                    tDur = FormatNumber(tAmt, 0, , , -1) & "&nbsp;Hrs"
                ElseIf tAmt = 24 Then
                    tDur = FormatNumber((tAmt / 24), 0, , , -1) & "&nbsp;Day"
                ElseIf tAmt > 24 Then
                    tDur = FormatNumber((tAmt / 24), 1, , , -1) & "&nbsp;Day(s)"
                End If

                tDur = "&nbsp;"
                response.write "<td align=""right"">" & tDur & "</td>"
                response.write "<td align=""right"">&nbsp;</td></tr>"
                tot = tot + .fields("finalamt")
                .MoveNext
            Loop

            response.write "<tr><td colspan=""3"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"
            'response.write "<tr><td colspan=""2""><b>TOTAL AMOUNT</b></td>"
            'response.write "<td align=""right""><b>" & (FormatNumber(tot, 2, , , -1)) & "</b></td></tr>"
            'response.write "<tr><td colspan=""3"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"
        End If

    End If

    .Close
End With

.MoveNext

'Loop
End If

.Close
End With

response.write "</table>"
response.write "</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"
response.write "<tr><td valign=""bottom"" align=""center"">"
response.write "<table align=""center"" width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""3"">"
'Response.Write "<tr style=""font-size:8pt""><td></td><td align=""center"">Hospital Laboratory.Copyright@2013</td></tr>"
'response.write "<tr style=""font-size:8pt""><td></td><td align=""center"">Software By : Spagad Technologies : 020-9426112</td></tr>"
response.write "</table>"
response.write "</td></tr>"
'///////////////////////////////////////////////////
'response.write "<tr>"
'response.write "<td align=""center""><hr>"
'response.write "</td>"
'response.write "</tr>"
''LAB Copy
'sql = "select distinct testcategoryid from investigation where labrequestid='" & Trim(request.querystring("labrequestid")) & "'"
'With rstPrn1
'.open qryPro.FltQry(sql), conn, 3, 4
'If .RecordCount > 0 Then
'.MoveFirst
''Do While Not .EOF
'cat = .fields("testcategoryid")
'tot = 0
'catNm = UCase(GetComboName("testcategory", cat))
'sql = GetTableSql("investigation")
'sql = sql & " and  investigation.labrequestid='" & Trim(request.querystring("labrequestid")) & "'"
'response.write "<tr>"
'response.write "<td align=""center"">"
'response.write "<table border=""0"" width=""600"" cellspacing=""0"" cellpadding=""0"">"
'response.write "<tr height=""10"">"
'response.write "<td align=""Center"" valign=""top""></td>"
'response.write "<td align=""center"" height=""10"" bgcolor=""white"" style=""font-size: 10pt; color: #CCCCFF;  font-family:Arial"">"
'AddReportHeader
'response.write "</td>"
'response.write "</tr>"
'response.write "</table>"
'response.write "</td>"
'response.write "</tr>"
'response.write "<tr>"
'response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'response.write "</tr>"
'response.write "<tr>"
'response.write "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:10pt"">"
'response.write "TEST REQUEST [LAB COPY]</td>"
'response.write "</tr>"
'response.write "<tr>"
'response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'response.write "</tr>"
'response.write "<tr>"
'response.write "<td align=""center"">"
'response.write "<table border=""0"" width=""600"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 10pt; font-family: Arial"">"
'response.write "<caption>"
'response.write "<tr>"
'response.write "<td name=""tdLabelInplabrequestID"" id=""tdLabelInplabrequestID"" style=""font-weight: bold"">Path. No.</td>"
'response.write "<td width=""10""></td>"
'response.write "<td name=""tdInputInplabrequestID"" id=""tdInputInplabrequestID"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"" >" & (GetRecordField("labrequestID")) & "</td>"
'response.write "<td width=""10""></td>"
'response.write "<td name=""tdLabelInplabrequestName"" id=""tdLabelInplabrequestName"" style=""font-weight: bold"">Patient Name</td>"
'response.write "<td width=""10""></td>"
'
'If VisitNoExempt(GetRecordField("visitationid")) Then
'response.write "<td name=""tdInputInplabrequestName"" id=""tdInputInplabrequestName"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"">" & (GetRecordField("RequestName")) & "</td>"
'Else
'response.write "<td name=""tdInputInplabrequestName"" id=""tdInputInplabrequestName"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:11pt"">" & (GetRecordField("PatientName")) & "</td>"
'End If
'response.write "<td width=""10""></td>"
'
'response.write "<tr>"
'response.write "<td name=""tdLabelInpAge"" id=""tdLabelInpAge"" style=""font-weight: bold"">Age</td>"
'response.write "<td width=""10""></td>"
'
'If VisitNoExempt(GetRecordField("visitationid")) Then
'response.write "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & (GetRecordField("RequestAge")) & "</td>"
'Else
'response.write "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & (GetPatientAge(GetRecordField("visitationid"))) & "</td>"
'End If
'
'response.write "<td width=""10""></td>"
'response.write "<td name=""tdLabelInpGenderID"" id=""tdLabelInpGenderID"" style=""font-weight: bold"">Gender</td>"
'response.write "<td width=""10""></td>"
'
'If VisitNoExempt(GetRecordField("visitationid")) Then
'response.write "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & (GetRecordField("RequestGenderName")) & "</td>"
'Else
'response.write "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & (GetRecordField("GenderName")) & "</td>"
'End If
'
'response.write "<td width=""10""></td>"
'response.write "</tr>"
'response.write "<tr>"
'response.write "<td name=""tdLabelInpReceiptTypeID"" id=""tdLabelInpReceiptTypeID"" style=""font-weight: bold"">Request Type</td>"
'response.write "<td width=""10""></td>"
'response.write "<td name=""tdInputInpReceiptTypeID"" id=""tdInputInpReceiptTypeID"">" & (GetRecordField("ReceiptTypeName")) & "</td>"
'response.write "<td width=""10""></td>"
'response.write "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Organization</td>"
'response.write "<td width=""10""></td>"
'response.write "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">" & (GetRecordField("InsuranceSchemeName")) & "</td>"
'response.write "<td width=""10""></td>"
'response.write "</tr>"
'
'response.write "<tr>"
'response.write "<td name=""tdLabelInpWorkingDayID"" id=""tdLabelInpWorkingDayID"" style=""font-weight: bold"">Request Day</td>"
'response.write "<td width=""10""></td>"
'response.write "<td name=""tdInputInpWorkingDayID"" id=""tdInputInpWorkingDayID"">" & (GetRecordField("WorkingDayName")) & "</td>"
'response.write "<td width=""10""></td>"
'response.write "<td name=""tdLabelInpReceiptDate"" id=""tdLabelInpReceiptDate"" style=""font-weight: bold"">Request Time</td>"
'response.write "<td width=""10""></td>"
'response.write "<td name=""tdInputInpReceiptDate"" id=""tdInputInpReceiptDate"">" & (FormatDateDetail(GetRecordField("RequestDate"))) & "</td>"
'response.write "<td width=""10""></td>"
'response.write "</tr>"
'
'response.write "<tr>"
'response.write "<td name=""tdLabelInpSystemUserID"" id=""tdLabelInpSystemUserID"" style=""font-weight: bold"">Lab Staff</td>"
'response.write "<td width=""10""></td>"
'response.write "<td name=""tdInputInpSystemUserID"" id=""tdInputInpSystemUserID"">" & (GetRecordField("SystemUserName")) & "</td>"
'response.write "<td width=""10""></td>"
'response.write "<td name=""tdLabelInpInsuranceTypeID"" id=""tdLabelInpInsuranceTypeID"" style=""font-weight: bold"">Visit No.</td>"
'response.write "<td width=""10""></td>"
'response.write "<td name=""tdInputInpInsuranceTypeID"" id=""tdInputInpInsuranceTypeID"">" & (UCase(GetRecordField("VisitationID"))) & "</td>"
'response.write "<td width=""10""></td>"
'response.write "</tr> </caption>"
'response.write "</table>        </td>"
'response.write "</tr>"
'response.write "<tr>"
'response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
'response.write "</tr>"
'response.write "<tr>"
'response.write "<td align=""center"">"
'response.write "<table id=""tblMultiSelect"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""border-collapse: collapse"" bordercolor=""#111111"">"
'response.write "<tr>"
'response.write "<td align=""left"" colspan=""2"">"
'response.write "<table height=""10"" id=""tblMultiSelectSection"" cellspacing=""0"" cellpadding=""0"" border=""0"" style=""border-collapse: collapse; font-size:10pt"" bordercolor=""#999999"" width=""100%"">"
'response.write "<tr>"
'response.write "<td valign=""top"" width=""100%"">"
'response.write "<table style=""font-size: 9pt; font-family: Arial"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""2"">"
'With rstPrn2
'.open qryPro.FltQry(sql), conn, 3, 4
'If .RecordCount > 0 Then
'.MoveFirst
'response.write "<tr><td><b>SERVICE    DESCRIPTION</b></td><td align=""right""><b>AMOUNT</b></td></tr>"
'response.write "<tr><td colspan=""2"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"
'Do While Not .EOF
'response.write "<tr><td>" & (.fields("Labtestname")) & "</td><td align=""right"">" & (FormatNumber(.fields("finalamt"), 2, , , -1)) & "</td></tr>"
'tot = tot + .fields("finalamt")
'.MoveNext
'Loop
'response.write "<tr><td colspan=""2"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"
'response.write "<tr><td><b>TOTAL AMOUNT</b></td><td align=""right""><b>" & (FormatNumber(tot, 2, , , -1)) & "</b></td></tr>"
'response.write "<tr><td colspan=""2"" align=""center""><hr color=""#999999"" size=""1""></td></tr>"
'End If
'.Close
'End With
'.MoveNext
'response.write "</table>"
'response.write "</td>"
'response.write "</tr>"
'response.write "</table>"
'response.write "</td>"
'response.write "</tr>"
'response.write "</table>"
'response.write "</td>"
'response.write "</tr>"
'response.write "<tr><td valign=""bottom"" align=""center"">"
'response.write "<table align=""center"" width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""3"">"
'response.write "<tr style=""font-size:8pt""><td></td><td align=""center"">Hospital Laboratory.Copyright@2013</td></tr>"
'response.write "</table>"
'response.write "</td></tr>"
''Loop
'End If
'.Close
'End With
Function VisitNoExempt(vst)
    Dim arr
    Dim ul
    Dim num
    Dim lst
    Dim ot
    ot = False
    lst = "E01||V02||V03||V04||V05||E02"
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
    Dim ot
    Dim rst
    Dim sql
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

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
'Empty
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
