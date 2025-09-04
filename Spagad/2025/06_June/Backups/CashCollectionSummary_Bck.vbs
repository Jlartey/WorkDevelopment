'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
response.Write Glob_GetBootstrap5()

Dim dat, periodStart, periodEnd
dat = Request.queryString("PrintFilter0")

If dat = "" Then
    response.Write "<h3>Please select a date range.</h3>"
    response.End
End If

Dim totalCollectionMedifem, totalRefundMedifem, totalBalanceMedifem
Dim totalCollectionAFC, totalRefundAFC, totalBalanceAFC

dt = Split(dat, "||")
periodStart = dt(0)
periodEnd = dt(1)

totalCollectionMedifem = 0
totalRefundMedifem = 0
totalBalanceMedifem = 0

totalCollectionAFC = 0
totalRefundAFC = 0
totalBalanceAFC = 0

'generateReportmedifem
'generateReportafc

Dim grandTotalCollection, grandTotalRefund, grandTotalBalance
grandTotalCollection = totalCollectionMedifem + totalCollectionAFC
grandTotalRefund = totalRefundMedifem + totalRefundAFC
grandTotalBalance = totalBalanceMedifem + totalBalanceAFC

response.Write "<h5><b>CASH COLLECTION SUMMARY BY FACILITY<b><h5> <br>"
response.Write "<div class='container'>"
response.Write "<div class='row'>"
response.Write "<div class='col-md-6'>"
medifemReportTotals = generateReportmedifem()
response.Write "</div>"

response.Write "<div class='col-md-6'>"
afcReportTotals = generateReportafc()
response.Write "</div>"
response.Write "</div>"
response.Write "</div>"

'Sub generateReportmedifem()
Function generateReportmedifem()
    Set rst = CreateObject("ADODB.Recordset")
    response.Write "<h5>CASH COLLECTION SUMMARY BY MEDIFEM </h5> "
    response.Write " <table class='table table-bordered table-hover' style='font-size:13px'>"
    response.Write "<thead>"
    response.Write "<th>#</th>"
    response.Write "<th>Date </th>"
    response.Write "<th>Mode</th>"
    response.Write "<th>Collection</th>"
    response.Write "<th>Refund</th>"
    response.Write "<th>Balance</th>"
    response.Write "</thead>"
    
sql = "SELECT "
sql = sql & "WorkingDay.WorkingDayName, receipt.PaymentModeid, "
sql = sql & "SUM(Receipt.ReceiptAmount1) AS collection, "
sql = sql & "SUM(Receipt.PaidAmount) AS refund, "
sql = sql & "SUM(Receipt.ReceiptAmount1 - Receipt.PaidAmount) AS Balance "
sql = sql & "FROM Receipt "
sql = sql & "JOIN WorkingDay ON Receipt.WorkingDayID = WorkingDay.WorkingDayID "
sql = sql & "JOIN CustomerType ON Receipt.CustomerTypeID = CustomerType.CustomerTypeID "
sql = sql & "WHERE CustomerType.CustomerTypeID NOT IN ('C111') AND Receipt.Receiptdate BETWEEN '" & periodStart & "' AND '" & periodEnd & "' "
sql = sql & "GROUP BY WorkingDay.WorkingDayName, receipt.paymentModeid "
sql = sql & "ORDER BY WorkingDay.WorkingDayName"

    rst.open qryPro.FltQry(sql), conn, 3, 4

    cnt = 0

    With rst
        If Not .EOF And Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                cnt = cnt + 1
                workinday = .fields("WorkingDayName").value
                ReceiptAmot = .fields("collection").value
                PaidAmot = .fields("refund").value
                bal = .fields("Balance").value
                paymentMode = .fields("PaymentModeid").value ' Added line for payment mode
                response.Write "<tr>"
                response.Write "<td>" & cnt & "</td>"
                response.Write "<td>" & workinday & "</td>"
                response.Write "<td>" & GetComboName("Paymentmode", paymentMode) & "</td>" ' Added payment mode column
                response.Write "<td>" & ReceiptAmot & "</td>"
                response.Write "<td>" & PaidAmot & "</td>"
                response.Write "<td>" & bal & "</td>"
                response.Write "</tr>"

                totalCollectionMedifem = totalCollectionMedifem + ReceiptAmot
                
                totalRefundMedifem = totalRefundMedifem + PaidAmot
                totalBalanceMedifem = totalBalanceMedifem + bal
            
                .MoveNext
                
            Loop

            response.Write "<tr style='font-weight:bold;'>"
                  response.Write "<td colspan='3'>Total</td>" ' Adjusted colspan to accommodate payment mode column
            response.Write "<td>" & FormatNumber(totalCollectionMedifem, 2, -1, 0, -1) & "</td>"
            response.Write "<td>" & FormatNumber(totalRefundMedifem, 2, -1, 0, -1) & "</td>"
            response.Write "<td>" & FormatNumber(totalBalanceMedifem, 2, -1, 0, -1) & "</td>"
            response.Write "</tr>"
        Else
            response.Write "<tr><td colspan='5' style='text-align:center; background-color:Red; font-weight:bold;'>NO DATA FOUND</td></tr>"
        End If
        .Close
    End With
   response.Write " </table>"
    Set rst = Nothing
    
    generateReportmedifem = Array(totalCollectionMedifem, totalRefundMedifem, totalBalanceMedifem)
End Function

'Sub generateReportafc()
Function generateReportafc()
    Set rst = CreateObject("ADODB.Recordset")
    response.Write "<h5>CASH COLLECTION SUMMARY BY AFC</h5> "
    response.Write " <table class='table table-bordered table-hover' style='font-size:13px'>"
    response.Write "<thead>"
    response.Write "<th>#</th>"
    response.Write "<th>Date </th>"
    response.Write "<th>Mode</th>"
    response.Write "<th>Collection</th>"
    response.Write "<th>Refund</th>"
    response.Write "<th>Balance</th>"
    response.Write "</thead>"

    sql = " SELECT"
    sql = sql & " WorkingDay.WorkingDayName,receipt.PaymentModeid,"
    sql = sql & " SUM(CASE WHEN CustomerType.CustomerTypeID = 'C111' THEN Receipt.ReceiptAmount1 ELSE 0 END) AS collection,"
    sql = sql & "  SUM(CASE WHEN CustomerType.CustomerTypeID = 'C111' THEN Receipt.PaidAmount ELSE 0 END) AS refund,"
    sql = sql & " SUM(CASE WHEN CustomerType.CustomerTypeID = 'C111' THEN (Receipt.ReceiptAmount1 - Receipt.PaidAmount) ELSE 0 END) AS Balance"
    sql = sql & " From Receipt"
    sql = sql & " LEFT JOIN WorkingDay ON Receipt.WorkingDayID = WorkingDay.WorkingDayID"
    sql = sql & " LEFT JOIN CustomerType ON Receipt.CustomerTypeID = CustomerType.CustomerTypeID"
    sql = sql & " where receipt.receiptdate Between  '" & periodStart & "' AND  '" & periodEnd & "'"

    sql = sql & " GROUP BY WorkingDay.WorkingDayName,receipt.paymentModeid"
    sql = sql & " ORDER BY WorkingDay.WorkingDayName"

    rst.open qryPro.FltQry(sql), conn, 3, 4
 
    cnt = 0

   With rst
    If Not .EOF And Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            cnt = cnt + 1
            workinday = .fields("WorkingDayName").value
            ReceiptAmot = .fields("collection").value
            PaidAmot = .fields("refund").value
            bal = .fields("Balance").value
            paymentMode = .fields("PaymentModeid").value ' Added line for payment mode

            response.Write "<tr>"
            response.Write "<td>" & cnt & "</td>"
            response.Write "<td>" & workinday & "</td>"
            response.Write "<td>" & GetComboName("Paymentmode", paymentMode) & "</td>" ' Added payment mode column
            response.Write "<td>" & ReceiptAmot & "</td>"
            response.Write "<td>" & PaidAmot & "</td>"
            response.Write "<td>" & bal & "</td>"
            response.Write "</tr>"

            totalCollectionAFC = totalCollectionAFC + ReceiptAmot
            totalRefundAFC = totalRefundAFC + PaidAmot
            totalBalanceAFC = totalBalanceAFC + bal

            .MoveNext
        Loop

        response.Write "<tr style='font-weight:bold;'>"
        response.Write "<td colspan='3'>Total</td>" ' Adjusted colspan to accommodate payment mode column
        response.Write "<td>" & FormatNumber(totalCollectionAFC, 2, -1, 0, -1) & "</td>"
        response.Write "<td>" & FormatNumber(totalRefundAFC, 2, -1, 0, -1) & "</td>"
        response.Write "<td>" & FormatNumber(totalBalanceAFC, 2, -1, 0, -1) & "</td>"
        response.Write "</tr>"
    Else
        response.Write "<tr><td colspan='6' style='text-align:center; background-color:Red; font-weight:bold;'>NO DATA FOUND</td></tr>" ' Adjusted colspan to accommodate payment mode column
    End If
    .Close
End With

    response.Write " </table>"
    
    Set rst = Nothing
    
    generateReportafc = Array(totalCollectionAFC, totalRefundAFC, totalBalanceAFC)
End Function

finalGrandTotal = medifemReportTotals(0) + afcReportTotals(0)
finalGrandTotalRefund = medifemReportTotals(1) + afcReportTotals(1)
finalGrandTotalBalance = medifemReportTotals(2) + afcReportTotals(2)

response.Write "<h5>Grand Total Cash Collection, Refund, and Balance</h5>"
response.Write "<div class='container'>"
response.Write "<div class='row'>"
response.Write "<div class='col-md-12'>"
response.Write "<b>Grand Total</b><br>"
response.Write "Total Collection: " & FormatNumber(finalGrandTotal, 2, -1, 0, -1) & "<br>"
response.Write "Total Refund: " & FormatNumber(finalGrandTotalRefund, 2, -1, 0, -1) & "<br>"
response.Write "Total Balance: " & FormatNumber(finalGrandTotalBalance, 2, -1, 0, -1) & "<br>"
response.Write "</div>"
response.Write "</div>"
response.Write "</div>"

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
