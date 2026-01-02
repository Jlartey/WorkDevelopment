'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim dt, tmpDt

dt = Split(Request.QueryString("PrintFilter0"), "||")

If UBound(dt) < 0 Then
    tmpDt = FormatDate(Now)
    dt = Array(tmpDt & " 00:00:00", tmpDt & " 23:59:59")
End If

Call ShowReports(dt)

Sub ShowReports(dt)
    Dim sql, args, rptGen, tmp
    
    Set rptGen = New PRTGLO_RptGen
    
    tmp = GetInvoicesWithUnusedPayments(dt)
    rptGen.AddReport tmp(0), tmp(1)
    
    tmp = GetInvoicesWithoutPayments(dt)
    rptGen.AddReport tmp(0), tmp(1)
    
    rptGen.StyleAsDashboard = True
    rptGen.ShowReport
End Sub
Function GetInvoicesWithoutPayments(dt)
    Dim sql, args, rptGen, title
    
    sql = "select PatientFlag2.PatientFlag2ID as [Invoice No]"
    sql = sql & " , '[' + PatientFlag2.PatientID + '] ' + (case when PatientFlag2.PatientID='P1' then PatientFlag2.PatientFlag2Name else Patient.PatientName end) as [Patient]"
    sql = sql & " , PatientFlag2.FlagDetail2 as [Invoice Detail]"
    sql = sql & " , PatientFlag2.FlagValue1 as [Invoice Amount]"
    sql = sql & " , PatientFlag2.FlagInfo4 as [TableName]"
    sql = sql & " , Staff.StaffName as [Issued By] "
    sql = sql & " , PatientFlag2.EntryDate as [Date Issued] "
    sql = sql & " from PatientFlag2  "
    sql = sql & " left join Patient on Patient.PatientID=PatientFlag2.PatientID"
    sql = sql & " left join Receipt on PatientFlag2.PatientFlag2ID=Receipt.ReceiptInfo1"
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=PatientFlag2.SystemUserID"
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID "
    sql = sql & " where 1=1 "
    sql = sql & "   and Receipt.ReceiptID is null "
    
    If UBound(dt) > 0 Then
        If IsDate(dt(0)) Then
            sql = sql & " and  PatientFlag2.EntryDate >='" & dt(0) & "' "
        End If
        If IsDate(dt(0)) Then
            sql = sql & " and  PatientFlag2.EntryDate <='" & dt(1) & "' "
        End If
    End If
    
    title = "Invoices without payments<br> " & dt(0) & " - " & dt(1)
    
    args = args & "title=" & title
    args = args & ";HiddenFields=TableName"
    args = args & ";FieldFunctions=Invoice Detail:FormatInvoiceDetail|Date Issued:FormatInvoiceDate|Payment Date:FormatInvoiceDate"
    args = args & ";formatMoneyFields=Invoice Amount"
    args = args & ";IgnoreFromComputations=ReceiptID|Invoice No|Patient|Invoice Detail|Date Issued|Issued By|TableName"
    GetInvoicesWithoutPayments = Array(sql, args)
    
End Function
Function GetInvoicesWithUnusedPayments(dt)
    Dim sql, args, rptGen, title
    
    sql = "select Receipt.ReceiptID, Receipt.PatientID"
    sql = sql & " , PatientFlag2.PatientFlag2ID as [Invoice No]"
    sql = sql & " , '[' + Receipt.PatientID + '] ' + (case when Receipt.PatientID='P1' then Receipt.ReceiptName else Patient.PatientName end) as [Patient]"
    sql = sql & " , PatientFlag2.FlagDetail2 as [Invoice Detail]"
    sql = sql & " , Receipt.ReceiptAmount1 as [Paid]"
    sql = sql & " , Receipt.ReceiptAmount2 as [Used]"
    sql = sql & " , Receipt.ReceiptAmount3 as [Balance]"
    sql = sql & " , PatientFlag2.EntryDate as [Invoice Date]"
    sql = sql & " , Receipt.ReceiptDate as [Payment Date]"
    sql = sql & " , PatientFlag2.FlagInfo4 as [TableName]"
    sql = sql & " , Visitation.VisitationID as [Visit No]"
    sql = sql & " from Receipt "
    sql = sql & " inner join PatientFlag2 on PatientFlag2.PatientFlag2ID=Receipt.ReceiptInfo1"
    sql = sql & " left join Patient on Patient.PatientID=Receipt.PatientID"
    sql = sql & " left join Visitation on Visitation.PatientID=Patient.PatientID and Visitation.VisitationID=PatientFlag2.FlagInfo2"
    sql = sql & " where 1=1 "
    sql = sql & "   and ReceiptAmount3 > 0"
    sql = sql & "   and ((Receipt.ReceiptAmount1 - Receipt.ReceiptAmount2 - Receipt.PaidAmount)/Receipt.ReceiptAmount1*100)>=1" '1% margin of error
    If UBound(dt) > 0 Then
        If IsDate(dt(0)) Then
            sql = sql & " and  Receipt.ReceiptDate >='" & dt(0) & "' "
        End If
        If IsDate(dt(0)) Then
            sql = sql & " and  Receipt.ReceiptDate <='" & dt(1) & "' "
        End If
    End If
    sql = sql & " order by Receipt.ReceiptDate asc "
    
    title = "Invoices with unused payments<br> " & dt(0) & " - " & dt(1)
    
    args = args & "title=" & title
    args = args & ";ExtraFields=Actions"
'    args = args & ";HiddenFields=TableName|PatientID"
    args = args & ";FieldFunctions=Invoice Detail:FormatInvoiceDetail|Date Issued:FormatInvoiceDate|Payment Date:FormatInvoiceDate|Invoice Date:FormatInvoiceDate|Actions:GenerateBill"
    args = args & ";formatMoneyFields=Paid|used|balance"
    args = args & ";IgnoreFromComputations=ReceiptID|Invoice No|Patient|Invoice Detail|Date Issued|TableName|Payment Date|Visit No|Actions|Invoice Date"
    GetInvoicesWithUnusedPayments = Array(sql, args)
End Function
Function GenerateBill(RECOBJ, fieldNAme)
    Dim ot, lnk, vst, sql, rst, bNm
    
    vst = RECOBJ("Visit No")
    bNm = "[" & RECOBJ("SpecialisTypeName") & "]"

    If UCase(RECOBJ("PatientID")) = "P1" Then
        vst = "E01"
        bNm = ""
    ElseIf Not (Len(vst & "") > 0) Then
       'try grab some visits for the day
       sql = "select VisitationID, SpecialistTypeName "
       sql = sql & " from Visitation "
       sql = sql & " left join SpecialistType on SpecialistType.SpecialistTypeID=Visitation.SpecialistTypeID "
       sql = sql & " where 1=1 "
       sql = sql & "    and PatientID='" & RECOBJ("PatientID") & "' "
       sql = sql & "    and cast(VisitDate as date)= try_cast('" & RECOBJ("Invoice Date") & "' as date) "
      
       bNm = ""
       Set rst = CreateObject("ADODB.RecordSet")
       rst.open qryPro.FltQry(sql), conn, 3, 4
       If rst.RecordCount > 0 Then
            rst.MoveFirst
            Do While Not rst.EOF
                vst = vst & rst.fields("VisitationID")
                rst.MoveNext
            Loop
       End If
       rst.Close
    End If

    Select Case UCase(RECOBJ("TableName"))
        Case "CONSULTREVIEW"
            lnk = "wpgConsultReview.asp?PageMode=AddNew&PullUpData=VisitationID||" & vst
        Case "LABREQUEST"
            lnk = "wpgLabRequest.asp?PageMode=AddNew&PullUpData=VisitationID||" & vst
        Case "DRUGSALE"
            lnk = "wpgDrugSale.asp?PageMode=AddNew&PullUpData=VisitationID||" & vst
        Case Else
            lnk = ""
    End Select
    
    If Len(lnk) > 0 And Len(vst) > 0 Then
        ot = "<a href=""" & lnk & """ target=""_blank"">Generate Bill" & bNm & "</a>"
    End If
    
    GenerateBill = ot
End Function
Function FormatInvoiceDetail(RECOBJ, fieldNAme)
    Dim ot, tmp
    
    ot = ""
    If Not IsNull(RECOBJ(fieldNAme)) Then
        ot = "<ul>"
        For Each tmp In Split(RECOBJ(fieldNAme), vbCrLf)
            If Len(tmp) > 0 Then
                ot = ot & "<li>" & tmp & "</li>"
            End If
        Next
        ot = ot & "</ul>"
    End If
    FormatInvoiceDetail = ot
End Function
Function FormatInvoiceDate(RECOBJ, fieldNAme)
    Dim ot
    
    If IsDate(Trim(RECOBJ(fieldNAme))) Then
        ot = FormatDateDetail(Trim(RECOBJ(fieldNAme)))
    End If
    
    FormatInvoiceDate = ot
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
