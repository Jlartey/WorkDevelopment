'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim consultReviewID
consultReviewID = Trim(Request.QueryString("consultReviewID"))

Styles
PrintOut

Sub Styles()
    response.write "    <style>"
        response.write "      .header {"
        response.write "        display: flex;"
        response.write "        flex-direction: column;"
        response.write "        justify-content: center;"
        response.write "        align-items: center;"
        response.write "      }"
        response.write "      .header h2 {"
        response.write "        margin: 2px 0px;"
        response.write "        font-size: larger;"
        response.write "        color: black;"
        response.write "      }"
        response.write "      .first-head {"
        response.write "        display: flex;"
        response.write "        justify-content: center;"
        response.write "        align-items: center;"
        response.write "        font-size: small;"
        response.write "        padding-bottom: 10px;"
        response.write "      }"
        response.write "      .first-head p {"
        response.write "        margin: 2px 0px;"
        response.write "      }"
        response.write "      .first-head img {"
        response.write "        height: 8vh;"
        response.write "        width: auto;"
        response.write "      }"
        response.write "      .main {"
        response.write "        position: relative;"
        response.write "        font-family: sans-serif;"
        response.write "        padding: 0.5rem;"
        response.write "      }"
        response.write "      .main::after {"
        response.write "        content: '';"
        response.write "        display: block;"
        response.write "        position: absolute;"
        response.write "        top: 0;"
        response.write "        right: 0;"
        response.write "        bottom: 0;"
        response.write "        left: 0;"
        response.write "        background-image: url(images/banner1.bmp);"
        response.write "        opacity: 0.06;"
        response.write "        z-index: -1;"
        response.write "        background-position: center;"
        response.write "        background-size: 80%;"
        response.write "        background-repeat: no-repeat;"
        response.write "      }"
        response.write "      .header-style {"
        response.write "        width: 100%;"
        response.write "        display: flex;"
        response.write "        flex-wrap: wrap;"
        response.write "        justify-content: space-between;"
        response.write "        font-size: 12px;"
        response.write "        padding: 8px 0;"
        response.write "        border-bottom: 1px solid #ccc;"
        response.write "        margin-bottom: 10px;"
        response.write "      }"
        response.write "      .header-style label {"
        response.write "        margin: 2px 0;"
        response.write "      }"
        response.write "      .table {"
        response.write "        width: 100%;"
        response.write "        font-family: monospace;"
        response.write "        border-collapse: collapse;"
        response.write "      }"
        response.write "      .table td {"
        response.write "        padding: 4px;"
        response.write "      }"
        response.write "      .myfont {"
        response.write "        font-size: 12px;"
        response.write "      }"
        response.write "      .amount-row td {"
        response.write "        padding-top: 10px;"
        response.write "      }"
    response.write "    </style>"
End Sub

Sub PrintOut()
    response.write "    <div style=""width: 80mm; margin: auto"">"
    response.write "      <main class=""main"" style=""width: 75mm; margin: 0 auto"">"
    response.write "        <div class=""header"">"
    response.write "          <div class=""first-head"">"
    response.write "            <img src=""images/banner1.bmp"" />"
    response.write "          </div>"
    response.write "          <div class=""first-head"">"
    response.write "            <div style=""flex-direction: column; padding: 8px 8px"">"
    response.write "              <h2 style=""text-align: center"">"
    response.write "                International Maritime Hospital(IMaH)"
    response.write "              </h2>"
    response.write "              <p style=""text-align: center"">Tel: +233(0)303-220030</p>"
    response.write "              <p style=""text-align: center"">"
    response.write "                Location: Community One, Tema, Accra"
    response.write "              </p>"
    response.write "            </div>"
    response.write "          </div>"
    response.write "        </div>"
    response.write ""
    response.write "        <div class=""header-style"">"
    response.write "          <label>Item Sale ID: " & consultReviewID & "</label>"
    response.write "          <label>Cashier: " & GetCashierName(consultReviewID) & "</label>"
    response.write "          <label>Date: " & GetReceiptDate(consultReviewID) & "</label>"
    response.write "        </div>"
    response.write ""
    response.write "        <table class=""table"">"
    response.write "          <tr>"
    response.write "            <td class=""myfont""><b>ITEM NAME</b></td>"
    response.write "            <td class=""myfont""><b>QTY</b></td>"
    response.write "            <td class=""myfont""><b>PRICE</b></td>"
    response.write "            <td class=""myfont""><b>TOTAL</b></td>"
    response.write "          </tr>"
    GetReceiptDetails consultReviewID
    response.write "          <tr class=""amount-row"">"
    response.write "            <td class=""myfont"" colspan=""2""><b>AMOUNT:</b></td>"
    response.write "            <td class=""myfont"" colspan=""2""><b>GH&#8373: " & GetTotalAmount(consultReviewID) & "</b></td>"
    response.write "          </tr>"
    response.write "        </table>"
    response.write "      </main>"
    response.write "    </div>"
    response.write "  "

End Sub


Sub GetReceiptDetails(consultReviewID)

    Dim rst, sql

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT ConsultReview.ConsultReviewID,ConsultReview.ConsultReviewName "
    sql = sql & ",ConsultReview.ConsultReviewDate, Treatment.TreatmentID "
    sql = sql & ", TreatCharges.Qty AS qty, TreatCharges.UnitCost, TreatCharges.FinalAmt "
    sql = sql & "FROM ConsultReview "
    sql = sql & "INNER JOIN TreatCharges ON TreatCharges.ConsultReviewID = ConsultReview.ConsultReviewID "
    sql = sql & "INNER JOIN Treatment ON Treatment.TreatmentID = TreatCharges.TreatmentID "
    sql = sql & "WHERE Treatment.BillGroupID = 'LSHHI170' AND ConsultReview.consultreviewid = '" & consultReviewID & "' "
    sql = sql & "ORDER BY ConsultReview.ConsultReviewDate desc"

    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.recordCount > 0 Then
    rst.MoveFirst
        Do While Not rst.EOF
             response.write "<tr class='myfont'>"
                response.write "<td class='mytd'>" & GetComboName("Treatment", rst.fields("TreatmentID")) & "</td>"
                response.write "<td class='mytd'>" & rst.fields("qty") & "</td>"
                response.write "<td class='mytd'>" & (FormatNumber(CStr(rst.fields("UnitCost")), 2, , , -1)) & "</td>"
                response.write "<td class='mytd'>" & (FormatNumber(CStr(rst.fields("FinalAmt")), 2, , , -1)) & "</td>"
                response.write "</tr>"

                rst.MoveNext

        Loop
    End If
    rst.Close
    Set rst = Nothing
End Sub

Function GetTotalAmount(consultReviewID)
    Dim rst, sql
    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT SUM(TreatCharges.FinalAmt) FinalAmt "
    sql = sql & "From ConsultReview "
    sql = sql & "INNER JOIN TreatCharges ON TreatCharges.ConsultReviewID = ConsultReview.ConsultReviewID "
    sql = sql & "INNER JOIN Treatment ON Treatment.TreatmentID = TreatCharges.TreatmentID "
    sql = sql & "WHERE Treatment.BillGroupID = 'LSHHI170' AND ConsultReview.ConsultReviewID = '" & consultReviewID & "'"
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4

        If .recordCount > 0 Then
            GetTotalAmount = (FormatNumber(CStr(.fields("FinalAmt")), 2, , , -1))
        End If
        .Close
    End With
    Set rst = Nothing
End Function


Function GetCashierName(consultReviewID)
    Dim rst, sql
    
    Set rst = Server.CreateObject("ADODB.Recordset")
    
    sql = "SELECT DISTINCT Staff.StaffName "
    sql = sql & "FROM ConsultReview "
    sql = sql & "INNER JOIN TreatCharges ON TreatCharges.ConsultReviewID = ConsultReview.ConsultReviewID "
    sql = sql & "INNER JOIN Treatment ON Treatment.TreatmentID = TreatCharges.TreatmentID "
    sql = sql & "INNER JOIN SystemUser ON SystemUser.SystemUserID = TreatCharges.SystemUserID "
    sql = sql & "INNER JOIN Staff ON Staff.StaffID = SystemUser.StaffID "
    sql = sql & "WHERE Treatment.BillGroupID = 'LSHHI170' AND ConsultReview.ConsultReviewID = '" & consultReviewID & "'"
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    If Not rst.EOF Then
        GetCashierName = rst.fields("StaffName")
    Else
        GetCashierName = " "
    End If
    
    rst.Close
   
    Set rst = Nothing
End Function

Function GetReceiptDate(consultReviewID)
    Dim rst, sql

    Set rst = Server.CreateObject("ADODB.Recordset")
    
    sql = "SELECT TOP  1 CONVERT(VARCHAR(20), ConsultReview.ConsultReviewDate, 106) + ' ' "
    sql = sql & "+ CONVERT(VARCHAR(8), ConsultReview.ConsultReviewDate, 108) AS ReceiptDate "
    sql = sql & "FROM ConsultReview "
    sql = sql & "INNER JOIN TreatCharges ON TreatCharges.ConsultReviewID = ConsultReview.ConsultReviewID "
    sql = sql & "INNER JOIN Treatment ON Treatment.TreatmentID = TreatCharges.TreatmentID "
    sql = sql & "WHERE Treatment.BillGroupID = 'LSHHI170' AND ConsultReview.consultreviewid = '" & consultReviewID & "'"

    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    If Not rst.EOF Then
        GetReceiptDate = rst.fields("ReceiptDate")
    Else
        GetReceiptDate = " "
    End If
    
    rst.Close

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






