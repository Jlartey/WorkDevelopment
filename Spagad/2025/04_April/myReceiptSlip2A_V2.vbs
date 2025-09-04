'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim drugsaleid
drugsaleid = Trim(Request.QueryString("drugsaleid"))

Styles
PrintOut

Sub Styles()
    response.Write "    <style>"
        response.Write "      .header {"
        response.Write "        display: flex;"
        response.Write "        flex-direction: column;"
        response.Write "        justify-content: center;"
        response.Write "        align-items: center;"
        response.Write "      }"
        response.Write "      .header h2 {"
        response.Write "        margin: 2px 0px;"
        response.Write "        font-size: larger;"
        response.Write "        color: black;"
        response.Write "      }"
        response.Write "      .first-head {"
        response.Write "        display: flex;"
        response.Write "        justify-content: center;"
        response.Write "        align-items: center;"
        response.Write "        font-size: small;"
        response.Write "        padding-bottom: 10px;"
        response.Write "      }"
        response.Write "      .first-head p {"
        response.Write "        margin: 2px 0px;"
        response.Write "      }"
        response.Write "      .first-head img {"
        response.Write "        height: 8vh;"
        response.Write "        width: auto;"
        response.Write "      }"
        response.Write "      .main {"
        response.Write "        position: relative;"
        response.Write "        font-family: sans-serif;"
        response.Write "        padding: 0.5rem;"
        response.Write "      }"
        response.Write "      .main::after {"
        response.Write "        content: '';"
        response.Write "        display: block;"
        response.Write "        position: absolute;"
        response.Write "        top: 0;"
        response.Write "        right: 0;"
        response.Write "        bottom: 0;"
        response.Write "        left: 0;"
        response.Write "        background-image: url(images/banner1.bmp);"
        response.Write "        opacity: 0.06;"
        response.Write "        z-index: -1;"
        response.Write "        background-position: center;"
        response.Write "        background-size: 80%;"
        response.Write "        background-repeat: no-repeat;"
        response.Write "      }"
        response.Write "      .header-style {"
        response.Write "        width: 100%;"
        response.Write "        display: flex;"
        response.Write "        flex-wrap: wrap;"
        response.Write "        justify-content: space-between;"
        response.Write "        font-size: 12px;"
        response.Write "        padding: 8px 0;"
        response.Write "        border-bottom: 1px solid #ccc;"
        response.Write "        margin-bottom: 10px;"
        response.Write "      }"
        response.Write "      .header-style label {"
        response.Write "        margin: 2px 0;"
        response.Write "      }"
        response.Write "      .table {"
        response.Write "        width: 100%;"
        response.Write "        font-family: monospace;"
        response.Write "        border-collapse: collapse;"
        response.Write "      }"
        response.Write "      .table td {"
        response.Write "        padding: 4px;"
        response.Write "      }"
        response.Write "      .myfont {"
        response.Write "        font-size: 12px;"
        response.Write "      }"
        response.Write "      .amount-row td {"
        response.Write "        padding-top: 10px;"
        response.Write "      }"
    response.Write "    </style>"
End Sub

Sub PrintOut()
    response.Write "    <div style=""width: 80mm; margin: auto"">"
    response.Write "      <main class=""main"" style=""width: 75mm; margin: 0 auto"">"
    response.Write "        <div class=""header"">"
    response.Write "          <div class=""first-head"">"
    response.Write "            <img src=""images/banner1.bmp"" />"
    response.Write "          </div>"
    response.Write "          <div class=""first-head"">"
    response.Write "            <div style=""flex-direction: column; padding: 8px 8px"">"
    response.Write "              <h2 style=""text-align: center"">"
    response.Write "                International Maritime Hospital(IMaH)"
    response.Write "              </h2>"
    response.Write "              <p style=""text-align: center"">Tel: +233(0)303-220030</p>"
    response.Write "              <p style=""text-align: center"">"
    response.Write "                Location: Community One, Tema, Accra"
    response.Write "              </p>"
    response.Write "            </div>"
    response.Write "          </div>"
    response.Write "        </div>"
    response.Write ""
    response.Write "        <div class=""header-style"">"
    response.Write "          <label>Item Sale ID: "& drugsaleid &"</label>"
    response.Write "          <label>Cashier: Lardy Tatania Palazar</label>"
    response.Write "          <label>Time: 8th April, 2024</label>"
    response.Write "        </div>"
    response.Write ""
    response.Write "        <table class=""table"">"
    response.Write "          <tr>"
    response.Write "            <td class=""myfont""><b>ITEM NAME</b></td>"
    response.Write "            <td class=""myfont""><b>QTY</b></td>"
    response.Write "            <td class=""myfont""><b>PRICE</b></td>"
    response.Write "            <td class=""myfont""><b>TOTAL</b></td>"
    response.Write "          </tr>"
    GetDrugItems drugsaleid
    response.Write "          <tr class=""amount-row"">"
    response.Write "            <td class=""myfont"" colspan=""2""><b>AMOUNT:</b></td>"
    response.Write "            <td class=""myfont"" colspan=""2""><b>GH&#8373: " & GetTotalAmount & "</b></td>"
    response.Write "          </tr>"
    response.Write "        </table>"
    response.Write "      </main>"
    response.Write "    </div>"
    response.Write "  "

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
                response.write "<td class='mytd'>" & (FormatNumber(CStr(rst.fields("unitcost")), 2, , , -1)) & "</td>"
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

