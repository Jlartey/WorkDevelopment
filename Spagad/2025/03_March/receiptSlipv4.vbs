'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim rDt, nMin, sDt, nMax
' sDt = GetRecordField("ServiceNo")
' rDt = GetRecordField("ReceiptDate")
' rcpID = GetRecordField("ReceiptID")

'If IsDate(sDt) Then
'    nMin = DateDiff("n", CDate(sDt), Now())
'    nMax = 15
'Else
'    nMin = DateDiff("n", CDate(rDt), Now())
'    nMax = 120
'End If
'If nMin < nMax Or True Then
    StylesAdded

    Printout

'End If


Sub Printout()
    Dim bal
    response.write " <div style='width: 74mm; margin: auto' >"
    response.write " <main class='main' style='width: 70mm; margin: auto'>"
    response.write " <div class='header'>"
    response.write "     <div style='padding: 0.5rem 0.5rem;' class='first-head'>"
    response.write "         <img src='images/banner1.bmp'/>"
    response.write "     </div>"
    response.write "     <div class='first-head'>"
    'response.write "         <img src='images/banner1.bmp'/>"
    response.write "         <div style='flex-direction: column; padding: 0.5rem 0.5rem;'>"
    response.write "             <h2>FOCOS Orthopaedic Hospital</h2>"
    response.write "             <p style='text-align: left;'>Tel: +233 59 692 0909/1 | Email: info@focosgh.com</p>"
    response.write "             <p style='text-align: left;'>Location: No.8 Teshie Street, Pantang, Accra | GPS: GM-109-8032</p>"
    'If UCase(GetRecordField("DrugSaleID")) <> UCase("P1") Then
    'response.write "             <p style='text-align: left;'>Customer Name: <b>" & (GetRecordField("PatientID")) & "</b></p>"
    'Else
    response.write "             <p style='text-align: left;'>Customer Name: <b>" & GetComboName("DrugSale", "D122070018") & "</b></p>"
    'End If
    response.write "         </div>"
    response.write "     </div>"
    response.write " </div>"
    response.write "     <div style='padding: 0.5rem 0.5rem;'>"
    response.write "         <p style='text-align: left;'>Receipt No: <span>" & (GetRecordField("ReceiptID")) & "</span></p>"
    response.write "         <p>Date / Time: <span>" & (FormatDateDetail(GetRecordField("ReceiptDate"))) & "</span></p>"
    response.write "     </div>"
    response.write "     <table class='table'>"
    response.write "         <tr>"
    response.write "             <td>Payment Mode: </td>"
    response.write "             <td><span><b>" & GetComboName("PaymentMode", GetRecordField("PaymentModeID")) & "</b></span></td>"
    response.write "         </tr>"
    If UCase(GetRecordField("PatientID")) <> UCase("P1") Then
    response.write "         <tr>"
    response.write "             <td>Received from: </td>"
    response.write "             <td><span>" & GetComboName("Patient", GetRecordField("PatientID")) & "</span></td>"
    response.write "         </tr>"
    Else
'    response.write "         <tr>"
'    response.write "             <td>Received from: </td>"
'    response.write "             <td><span>" & UCase(GetRecordField("ReceiptName")) & "</span></td>"
'    response.write "         </tr>"
    End If
    'response.write "         <tr>"
    'response.write "             <td>The sum of:</td>"
    'response.write "             <td>" & (UCase(GetPaymentWord(GetRecordField("ReceiptAmount1")))) & "</td>"
    'response.write "         </tr>"
'    response.write "         <tr>"
'    response.write "             <td>Being: </td>"
'    response.write "             <td><span>Payment For " & (GetRecordField("Remarks")) & "</span></td>"
'    response.write "         </tr>"

 ' Insert drug details table here
    response.write "         <tr>"
    response.write "             <td colspan='2'>"
    response.write "                 <table class='table' style='width: 100%;'>"
    response.write "                     <tr>"
    response.write "                         <td><b>DRUG NAME</b></td>"
    'response.write "                         <td><b>PRICE</b></td>"
    response.write "                         <td><b>QTY</b></td>"
    response.write "                         <td><b>TOTAL PRICE</b></td>"
    response.write "                     </tr>"
    GetDrugItems
    response.write "                 </table>"
    response.write "             </td>"
    response.write "         </tr>"
    
'  If GetRecordField("PaidAmount") > 0 Then
      response.write "         <tr>"
      response.write "             <td>AMOUNT: </td>"
'      response.write "             <td><span>" & FormatNumber(GetRecordField("ReceiptAmount1"), 2, , -1) & "</span></td>"
       response.write "             <td><span>" & GetTotalAmount & "</span></td>"
      response.write "         </tr>"
      ' response.write "         <tr>"
      ' response.write "             <td>REFUND: </td>"
      ' response.write "             <td><span>GH&#8373: " & FormatNumber(GetRecordField("PaidAmount"), 2, , -1) & "</span></td>"
      ' response.write "         </tr>"
      ' response.write "         <tr>"
      ' response.write "             <td>BAL. AMOUNT: </td>"
      ' response.write "             <td><span><b> GH&#8373: " & FormatNumber(GetRecordField("ReceiptAmount1") - GetRecordField("PaidAmount"), 2, , -1) & "</b></span></td>"
      ' response.write "         </tr>"
      ' bal = FormatNumber(GetRecordField("ReceiptAmount1") - GetRecordField("PaidAmount"), 2, , -1)
'      response.write "         <tr>"
'      response.write "             <td>The sum of:</td>"
'      response.write "             <td>" & (UCase(GetPaymentWord(bal))) & "</td>"
'      response.write "         </tr>"
'     Else
'       response.write "         <tr>"
'       response.write "             <td>AMOUNT: </td>"
' '      response.write "             <td><span><b>GH&#8373: " & FormatNumber(GetRecordField("ReceiptAmount1"), 2, , -1) & "</b></span></td>"
'        response.write "             <td><span><b>GH&#8373: " & FormatNumber(GetTotalAmount, 2, , -1) & "</b></span></td>"
'       response.write "         </tr>"
' '      response.write "         <tr>"
' '      response.write "             <td>The sum of:</td>"
' '      response.write "             <td>" & (UCase(GetPaymentWord(GetRecordField("ReceiptAmount1")))) & "</td>"
' '      response.write "         </tr>"
'     End If
    response.write "     </table>"
    'response.write "     <div style='padding: 0.5rem 0.5rem;'>"
    'response.write "         <p class='amount'>GH&#8373: " & FormatNumber(GetRecordField("ReceiptAmount1"), 2, , -1) & "</p>"
    'response.write "         <p class='signature'>Signature: ..............................................................................</p>"
    'response.write "     </div>"
    response.write " </main>"
    response.write " <h6 class='signature'><Label>cashier: &nbsp;</Label>" & (Replace(GetComboName("Staff", GetComboNameFld("SystemUser", GetRecordField("SystemUserID"), "StaffID")), " ", "&nbsp;")) & "</h6>"
    response.write " </div>"
End Sub

Sub pharmacySlip(rcpID)

    response.write " <header class='header'>"
    response.write "     <div class='first-head'>"
    response.write "         <img src='images/banner1.bmp'/>"
    response.write "         <div>"
    response.write "             <h2>FOCOS Orthopaedic Hospital</h2>"
    response.write "             <p style='text-align: left;>Tel: +233 59 692 0909/1 | Email: info@focosgh.com</p>"
    response.write "             <p style='text-align: left;>Location: No.8 Teshie Street, Pantang, Accra | GPS: GM-109-8032</p>"
    response.write "         </div>"
    response.write "     </div>"
    response.write " </header>"
    response.write " <main class='main'>"
    response.write "     <div>"
    response.write "         <p style='text-align: left;'>Receipt No: <span>" & (GetRecordField("ReceiptID")) & "</span></p>"
    response.write "         <p>Date / Time: <span>" & (FormatDateDetail(GetRecordField("ReceiptDate"))) & "</span></p>"
    response.write "     </div>"
    response.write "     <table class='table'>"
    response.write "         <tr>"
    response.write "             <td><b>DESCRIPTION</b></td>"
    response.write "             <td><b>PRICE</b></td>"
    response.write "             <td><b>QTY</b></td>"
    response.write "             <td><b>AMOUNT</b></td>"
    response.write "         </tr>"
    response.write GetDrugs(rcpID)
    response.write "     </table>"
    response.write "     <div>"
    response.write "         <p class='amount'>Amount Paid: GH&#8373 " & FormatNumber(GetRecordField("ReceiptAmount1"), 2, , -1) & "</p>"
    response.write "         <p class='signature'>Signature: ...........................</p>"
    response.write "     </div>"
    response.write " </main>"
    response.write " <h6 class='signature'><Label>Attendant:</Label>" & (Replace(GetComboName("Staff", GetComboNameFld("SystemUser", GetRecordField("SystemUserID"), "StaffID")), " ", "&nbsp;")) & "</h6>"

End Sub

Sub StylesAdded()

    response.write " <style>"
    response.write "     body{"
    response.write "         display: grid;"
    response.write "         justify-content: center;"
    response.write "         font-family:sans-serif;"
    response.write "     }"
    response.write "     .header{"
    response.write "         display: flex;"
    response.write "         flex-direction: column;"
    response.write "         justify-content: center;"
    response.write "     }"
    response.write "     .header h2{"
    response.write "         margin: 2px 0px;"
    response.write "         font-size: larger;"
    response.write "         color: black;"
    response.write "     }"
    response.write "     .first-head{"
    response.write "         display: flex;"
    response.write "         justify-content: center;"
    response.write "         align-items: center;"
    response.write "         font-size: small;"
    response.write "         padding-bottom: 10px;"
    response.write "     }"
    response.write "     .first-head p{"
    response.write "         margin: 2px 0px;"
    response.write "     }"
    response.write "     .first-head img{"
    response.write "         height: 8vh;"
    response.write "         width: auto;"
    response.write "     }"
    response.write "     .main{"
    response.write "         position: relative;;"
    response.write "     }"
    response.write "     .main::after{"
    response.write "         content: '';"
    response.write "         display: block;"
    response.write "         position: absolute;"
    response.write "         top: 0;"
    response.write "         right: 0;"
    response.write "         bottom: 0;"
    response.write "         left: 0;"
    response.write "         background-image: url(images/banner1.bmp);"
    response.write "         opacity: 0.06; "
    response.write "         z-index: -1;"
    response.write "         background-position: center;"
    response.write "         background-size: 80%;"
    response.write "         background-repeat: no-repeat;"
    response.write "     }"
    response.write "     .main p{"
    response.write "         text-align: end;"
    response.write "         font-size: small;"
    response.write "     }"
    response.write "     .main span{"
    response.write "         font-weight: bold;"
    response.write "         font-size: medium;"
    response.write "         font-family: monospace;"
    response.write "     }"
    response.write "     h4{"
    response.write "         font-size: medium;"
    response.write "         width: 100%;"
    response.write "         padding: 6px 0px 6px 0px;"
    response.write "         background-color: gainsboro;"
    response.write "         text-align: center;"
    response.write "     }"
    response.write "    .table{"
    response.write "         font-family: monospace;"
    'response.write "         padding-left: 20px;"
    response.write "         width: 100%;"
    response.write "         font-size: large;"
    response.write "     }"
    response.write "     .table span{"
    response.write "         font-size: large;"
    response.write "         font-weight: normal;"
    response.write "     }"
    response.write "     .amount{"
    response.write "         font-size: x-large;"
    response.write "         font-weight: bold;"
    response.write "     }"
    response.write "     .signature{"
    response.write "         font-size: small;"
    response.write "         padding-top: 5vh;"
    response.write "     }"
    response.write "     .main div{"
    response.write "         display: flex;"
    'response.write "         padding-left: 20px;"
    response.write "     }"
    response.write " </style>"

End Sub

Function GetPaymentWord(inAmt)
    Dim amt, fAmt, wAmt, ot
    ot = ""
    amt = Abs(CDbl(inAmt))
    wAmt = Int(amt)
    fAmt = Round(amt - wAmt, 2)
    ot = ot & GetAmountWord(wAmt) & " GHANA CEDI(S)"

    If fAmt > 0 Then
        ot = ot & " " & GetAmountWord(100 * fAmt) & " PESEWA(S)"
    End If
    GetPaymentWord = ot
End Function

Function GetAmountWord(inAmt)
    Dim amt, ot, amtRem, amtUnit
    amt = inAmt
    ot = ""
    If amt >= 1000000000 Then
        amtUnit = "Billion"
        ot = ot & " " & GetLess1000(Int(amt / 1000000000))
        ot = ot & " " & amtUnit
        amtRem = amt - (Int(amt / 1000000000) * 1000000000)
    ElseIf amt >= 1000000 Then
        amtUnit = "Million"
        ot = ot & " " & GetLess1000(Int(amt / 1000000))
        ot = ot & " " & amtUnit
        amtRem = amt Mod 1000000
    ElseIf amt >= 1000 Then
        amtUnit = "Thousand"
        ot = ot & " " & GetLess1000(Int(amt / 1000))
        ot = ot & " " & amtUnit
        amtRem = amt Mod 1000
    Else
        ot = ot & " " & GetLess1000(Int(amt / 1))
        amtRem = 0
    End If
    If amtRem > 0 Then
        ot = ot & " " & GetAmountWord(amtRem)
    End If
    GetAmountWord = ot
End Function

Function GetLess1000(Less1000)
    Dim ot, Less1000Rem
    ot = ""
    If Less1000 >= 100 Then
        ot = ot & " " & GetDigit(CStr(Int(Less1000 / 100)))
        ot = ot & " Hundred"
        Less1000Rem = Less1000 Mod 100
        If Less1000Rem > 0 Then
            ot = ot & " And"
        End If
    ElseIf Less1000 >= 10 Then
        If Less1000 >= 10 And Less1000 <= 19 Then
            Select Case Less1000
             Case 10
                ot = ot & "Ten"
             Case 11
                ot = ot & "Eleven"
             Case 12
                ot = ot & "Twelve"
             Case 13
                ot = ot & "Thirteen"
             Case 14
                ot = ot & "Fourteen"
             Case 15
                ot = ot & "Fifeteen"
             Case 16
                ot = ot & "Sixteen"
             Case 17
                ot = ot & "Seventeen"
             Case 18
                ot = ot & "Eighteen"
             Case 19
                ot = ot & "Nineteen"
             Case Else

            End Select
            Less1000Rem = 0
        Else
            ot = ot & " " & GetTens(Int(Less1000 / 10))
            Less1000Rem = Less1000 Mod 10
        End If
    ElseIf Less1000 < 10 Then
        ot = ot & " " & GetDigit(CStr(Less1000))
        Less1000Rem = 0
    End If

    If Less1000Rem > 0 Then
        ot = ot & " " & GetLess1000(Less1000Rem)
    End If
    GetLess1000 = ot
End Function

Function GetTens(tens)
    Dim ot
    ot = ""
    Select Case tens
     Case 1

     Case 2
        ot = ot & "Twenty"
     Case 3
        ot = ot & "Thirty"
     Case 4
        ot = ot & "Forty"
     Case 5
        ot = ot & "Fifty"
     Case 6
        ot = ot & "Sixty"
     Case 7
        ot = ot & "Seventy"
     Case 8
        ot = ot & "Eighty"
     Case 9
        ot = ot & "Ninety"
     Case Else
    End Select
    GetTens = ot
End Function

Function GetDigit(digit)
    Dim ot
    ot = ""
    Select Case digit
     Case "0"
        ot = "Zero"
     Case "1"
        ot = "One"
     Case "2"
        ot = "Two"
     Case "3"
        ot = "Three"
     Case "4"
        ot = "Four"
     Case "5"
        ot = "Five"
     Case "6"
        ot = "Six"
     Case "7"
        ot = "Seven"
     Case "8"
        ot = "Eight"
     Case "9"
        ot = "Nine"
     Case "10"
        ot = "Ten"
     Case "11"
        ot = "Eleven"
     Case "12"
        ot = "Twelve"
     Case Else
    End Select
    GetDigit = ot
End Function

Function GetDrugs(rcpID)
    Dim rst, sql, fnlamt, html, testID

    Set rst = CreateObject("ADODB.Recordset")

    sql = " SELECT PatientFlag2.FlagInfo1 FROM Receipt"
    sql = sql & " LEFT JOIN PatientFlag2 ON Receipt.ReceiptInfo1 = PatientFlag2.PatientFlag2ID"
    sql = sql & " WHERE Receipt.ReceiptID = '" & rcpID & "'"

    html = " "
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.movefirst
        ItemList = Split(rst.fields("FlagInfo1"), "**")
        
        For Each item In ItemList
            details = Split(item, "||")
            
            If UBound(details) >= 2 Then
                drugTable = details(0)
                drugid = details(1)
                drugQty = details(2)
                
                html = html & GetDrugItems(drugTable, drugid, drugQty)
            End If
        Next
    End If

    rst.Close
    Set rst = Nothing
    GetDrugs = html
End Function

Sub GetDrugItems()
    Dim rst, sql
    
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "SELECT DrugSaleItems.Drugid, Drug.DrugName, unitcost, Qty, unitcost*Qty TotalPrice "
    sql = sql & "From DrugSaleItems "
    sql = sql & "JOIN Drug ON DrugSaleItems.DrugID = Drug.DrugID "
    sql = sql & "WHERE DrugSaleid = 'D122070018'"
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    If rst.recordCount > 0 Then
    rst.movefirst
        Do While Not rst.EOF
             response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & rst.fields("DrugName") & "</td>"
                'response.write "<td class='mytd'>" & FormatNumber(rst.fields("unitCost"), 2, , , -1) & "</td>"
                response.write "<td class='mytd'>" & rst.fields("Qty") & "</td>"
                response.write "<td class='mytd'>" & rst.fields("TotalPrice") & "</td>"
                response.write "</tr>"

                rst.MoveNext
                
        Loop
    End If
    rst.Close
    Set rst = Nothing
End Sub

Function GetTotalAmount()
    Dim rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "SELECT SUM(unitcost*Qty) TotalAmount FROM DrugSaleItems WHERE DrugSaleid = 'D122070018' "
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
        If .recordCount > 0 Then
            GetTotalAmount = .fields("TotalAmount")
        End If
        .Close
    End With
    Set rst = Nothing
End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
'None
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>

