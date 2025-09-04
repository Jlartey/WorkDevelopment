'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim drugsaleid
drugsaleid = Trim(Request.QueryString("drugsaleid"))
    StylesAdded
    Printout

Sub Printout()
    'Dim bal
    response.write " <div style='width: 74mm; margin: auto' >"
    response.write " <main class='main' style='width: 70mm; margin: auto'>"
    response.write " <div class='header'>"
    response.write "     <div style='padding: 0.5rem 0.5rem;' class='first-head'>"
    response.write "         <img src='images/banner1.bmp'/>"
    response.write "     </div>"
    response.write "     <div class='first-head'>"
    'response.write "         <img src='images/banner1.bmp'/>"
    response.write "         <div style='flex-direction: column; padding: 0.5rem 0.5rem;'>"
    response.write "             <h2>International Maritime Hospital(IMaH)</h2>"
    response.write "             <p style='text-align: center;'>Tel: +233(0)303-220030</p>"
    response.write "             <p style='text-align: center;'>Location: Community One, Tema, Accra</p>"
    
    'response.write "             <p style='text-align: left;'>Customer Name: <b>" & GetComboName("DrugSale", "& drugSaleID &") & "</b></p>"
    'response.write "<p style='text-align: left;'>Customer Name: <b>" & GetComboName("DrugSale", "'" & drugSaleID & "'") & "</b></p>"

    'End If
    response.write "         </div>"
    response.write "     </div>"
    response.write " </div>"
   response.write "     <div style='padding: 0.5rem 0.5rem;'>"

    response.write "         <tr>"
    response.write "             <td colspan='2'>"
    response.write "                 <table style='width: 60%;margin: 0 auto'>"
    response.write "                     <tr>"
    response.write "                         <td class='myfont'><b>ITEM NAME</b></td>"
    response.write "                         <td><b>PRICE</b></td>"
    response.write "                         <td class='myfont'><b>QTY</b></td>"
    response.write "                         <td class='myfont'><b>TOTAL</b></td>"
    response.write "                     </tr>"
    GetDrugItems drugsaleid
    response.write "         <tr>"
    response.write "             <td></td>"
    response.write "             <td></td>"
    response.write "             <td></td>"
    response.write "         </tr>"
    response.write "         <tr>"
    response.write "             <td class='myfont' colspan='2'><b>AMOUNT: </b></td>"
    response.write "             <td class='myfont'><span><b>GH&#8373: " & GetTotalAmount & "</b></span></td>"
    response.write "         </tr>"
    response.write "                 </table>"
    response.write "             </td>"
    response.write "         </tr>"

    response.write "     </table>"
    response.write " </main>"

    response.write " </div>"
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
    response.write "     .myfont{"
    response.write "         font-size: 12px;"
    response.write "     }"
    response.write " </style>"

End Sub

Sub GetDrugItems(drugsaleid)
    
    Dim rst, sql
    
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "SELECT DrugSaleItems.Drugid, Drug.DrugName, unitcost, Qty, unitcost*Qty AS TotalPrice "
    sql = sql & "From DrugSaleItems "
    sql = sql & "JOIN Drug ON DrugSaleItems.DrugID = Drug.DrugID "
    sql = sql & "WHERE DrugSaleid = '" & drugsaleid & "'"

    
'    response.write sql
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    If rst.recordCount > 0 Then
    rst.MoveFirst
        Do While Not rst.EOF
             response.write "<tr class='myfont'>"
                response.write "<td class='mytd'>" & rst.fields("DrugName") & "</td>"
                'response.write "<td class='mytd'>" & (FormatNumber(CStr(rst.fields("unitcost")), 2, , , -1)) & "</td>"
                response.write "<td class='mytd'>" & rst.fields("Qty") & "</td>"
                response.write "<td class='mytd'>" & (FormatNumber(CStr(rst.fields("TotalPrice")), 2, , , -1)) & "</td>"
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
    
    sql = "SELECT SUM(unitcost*Qty) As TotalAmount FROM DrugSaleItems WHERE DrugSaleid = '" & drugsaleid & "' "
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
        If .recordCount > 0 Then
            GetTotalAmount = (FormatNumber(CStr(.fields("TotalAmount")), 2, , , -1))
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
