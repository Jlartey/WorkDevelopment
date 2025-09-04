'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim periodStart, periodEnd

datePeriod = Trim(Request.QueryString("PrintFilter1"))

If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    periodStart = dateArr(0)
    periodEnd = dateArr(1)
Else
    periodStart = DateAdd("d", -365, Now())
    periodEnd = Now()
End If


response.write "<tr>"
response.write "<td align=""center"" bgcolor=""#FFFFFF"" style=""font-family: Arial; color: #111111; font-weight:bold; font-size:12pt"">"
response.write "Patient</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align=""center"">"
response.write "<table border=""0"" width=""" & (PrintWidth) & """ cellspacing=""0"" cellpadding=""0"" bgcolor=""White"" style=""font-size: 9pt; font-family: Arial"">"
response.write "       <tr>"
response.write "<td name=""tdLabelInpPatientID"" id=""tdLabelInpPatientID"" style=""font-weight: bold"">Patient No.</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpPatientID"" id=""tdInputInpPatientID"">" & (GetRecordField("PatientID")) & "</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdLabelInpPatientName"" id=""tdLabelInpPatientName"" style=""font-weight: bold"">PatientName</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpPatientName"" id=""tdInputInpPatientName"">" & (GetRecordField("PatientName")) & "</td>"
response.write "<td width=""20""></td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td name=""tdLabelInpPatientTypeID"" id=""tdLabelInpPatientTypeID"" style=""font-weight: bold"">Type</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpPatientTypeID"" id=""tdInputInpPatientTypeID"">" ' & (GetRecordField("PatientTypeName")) & "</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdLabelInpGenderID"" id=""tdLabelInpGenderID"" style=""font-weight: bold"">Gender</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & (GetRecordField("GenderName")) & "</td>"
response.write "<td width=""20""></td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td name=""tdLabelInpBirthDate"" id=""tdLabelInpBirthDate"" style=""font-weight: bold"">BirthDate</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpBirthDate"" id=""tdInputInpBirthDate"">" & (FormatDateDetail(GetRecordField("BirthDate"))) & "</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdLabelInpAge"" id=""tdLabelInpAge"" style=""font-weight: bold"">Age</td>"
response.write "<td width=""20""></td>"

If IsDate(GetRecordField("BirthDate")) Then
  response.write "<td name=""tdInputInpAge"" id=""tdInputInpAge"">" & CStr(Round((DateDiff("d", CDate(GetRecordField("BirthDate")), Now())) / 365.25, 1)) & "</td>"
End If
response.write "<td width=""20""></td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td name=""tdLabelInpPatientTypeID"" id=""tdLabelInpPatientTypeID"" style=""font-weight: bold"">Telephone 1</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpPatientTypeID"" id=""tdInputInpPatientTypeID"">" & (GetRecordField("BusinessPhone")) & "</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdLabelInpGenderID"" id=""tdLabelInpGenderID"" style=""font-weight: bold"">Telephone 2</td>"
response.write "<td width=""20""></td>"
response.write "<td name=""tdInputInpGenderID"" id=""tdInputInpGenderID"">" & (GetRecordField("ResidencePhone")) & "</td>"
response.write "<td width=""20""></td>"
response.write "</tr>"

response.write "</table>"
response.write "</td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td align=""center""><hr color=""#999999"" size=""1""></td>"
response.write "</tr>"
response.write "<td style=""font-family: Arial; color: #111111; font-weight:bold; font-size:10pt"">"
response.write "Receipt History</td>"
response.write "</tr>"

response.write "<tr>"
response.write "<td align=""left"">"
  DisplayReceiptHist (GetRecordField("PatientID"))
response.write "</td></tr>"

Sub DisplayReceiptHist(patid)
  Dim rs, ot, sql, dt1, dt2, drg, sto, hrf, ky
    Set rs = CreateObject("ADODB.Recordset")
    ot = 0
    'dt2 = FormatDateDetail(Now())
   ' dt1 = FormatDateDetail(DateAdd("d", -365, Now()))
    sql = "select * from Receipt where patientid='" & patid & "' and receiptdate between '" & periodStart & "' and '" & periodEnd & "' order by receiptdate desc"
    
    With rs
      .open sql, conn, 3, 4
      If .recordCount > 0 Then
        .MoveFirst
        response.write "<table cellspacing=""0"" cellpadding=""3"" border=""1"" style=""border-collapse:collapse;font-size:9pt"">"
        
        response.write "<tr align=""center"">"
        response.write "<td><b>Date</b></td>"
        response.write "<td><b>Receipt No.</b></td>"
        response.write "<td style='white-space: nowrap;'><b>Receipt Name</b></td>"
        response.write "<td style='white-space: nowrap;'><b>Amount Paid</b></td>"
        response.write "<td style='white-space: nowrap;'><b>Amount Used</b></td>"
        response.write "<td style='white-space: nowrap;'><b>Balance Amount</b></td>"
        response.write "<td><b>Cashier</b></td>"
        response.write "<td><b>Clinic</b></td>"
        response.write "</tr>"
        Do While Not .EOF
          drg = .fields("ReceiptName")
          sto = .fields("BranchID")
          ky = .fields("ReceiptID")
          'hrf = "wpgDrugReturn.asp?PageMode=AddNew&pullupdata=DrugSaleID||" & ky
          response.write "<tr>"
          response.write "<td>" & FormatDateDetail(.fields("Receiptdate")) & "</td>"
          response.write "<td><b>" & ky & "</b></td>"
          response.write "<td>" & drg & "</td>"
          response.write "<td align=""right"">" & CStr(.fields("ReceiptAmount1")) & "</td>"
          response.write "<td align=""right"">" & CStr(.fields("ReceiptAmount2")) & "</td>"
          response.write "<td align=""right"">" & CStr(.fields("ReceiptAmount3")) & "</td>"
          response.write "<td>" & .fields("SystemUserID") & "</td>"
          response.write "<td>" & GetComboName("Branch", sto) & "</td>"
          response.write "</tr>"
          .MoveNext
        Loop
        response.write "</table>"
      End If
      .Close
    End With
    Set rs = Nothing
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
'None
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>

