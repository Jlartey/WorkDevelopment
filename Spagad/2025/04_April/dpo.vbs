'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim drugPurOrderID
drugPurOrderID = Trim(Request.Querystring("DrugPurOrderID"))
Styles
Printout

Sub Styles()
    response.write "    <style>"
    response.write "      .main {"
    response.write "        margin-top: 5px;"
    response.write "        font-size: 12pt;"
    response.write "        font-family: Arial, Helvetica, sans-serif;"
    response.write "        width: 210mm;"
    response.write "        height: 297mm;"
    response.write "        margin-left: auto;"
    response.write "        margin-right: auto;"
    response.write "        text-align: left; "
    response.write "      }"
    response.write ""
    response.write "      .header {"
    response.write "        width: 100%;"
    response.write "        margin: 1.6px 0;"
    response.write "        display: flex;"
    response.write "        align-items: center;"
    response.write "        justify-content: center;"
    response.write "        gap: 10px;"
    response.write "        position: relative;"
    response.write "      }"
    response.write ""
    response.write "      .image {"
    response.write "        position: absolute;"
    response.write "        left: 60px;"
    response.write "        margin-top: 20px;"
    response.write "      }"
    response.write ""
    response.write "      .header-text {"
    response.write "        flex-grow: 1;"
    response.write "        text-align: center;"
    response.write "      }"
    response.write ""
    response.write "      .header-text h1 {"
    response.write "        font-size: 1.5rem;"
    response.write "        margin: 0;"
    response.write "      }"
    response.write ""
    response.write "      .header-text h3 {"
    response.write "        font-size: 1rem;"
    response.write "        margin: 0.25rem 0;"
    response.write "      }"
    response.write ""
    response.write "      .contact-details {"
    response.write "        font-size: 10pt;"
    response.write "        line-height: 5px;"
    response.write "        width: auto;"
    response.write "        max-width: 25rem;"
    response.write "        margin: 0 0 0 auto;"
    response.write "        text-align: left; "
    response.write "      }"
    response.write ""
    response.write "      .contact-details p {"
    response.write "        display: flex;"
    response.write "        align-items: center;"
    response.write "      }"
    response.write ""
    response.write "      .contact-details p strong {"
    response.write "        width: 12.5rem;"
    response.write "        display: inline-block;"
    response.write "      }"
    response.write ""
    response.write "      .contact-details p span {"
    response.write "        margin-left: 0.3125rem;"
    response.write "        padding-right: 0.5rem;"
    response.write "      }"
    response.write ""
    response.write "      .address-and-po-details {"
    response.write "        display: flex;"
    response.write "        flex-direction: row;"
    response.write "        gap: 93px;"
    response.write "        margin-top: 5px;"
    response.write "        width: 100%;"
    response.write "      }"
    response.write ""
    response.write "      .address {"
    response.write "        flex: 1;"
    response.write "        flex-basis: 100px;"
    response.write "        flex-wrap: wrap;"
    response.write "        font-weight: bold;"
    response.write "        line-height: 10px;"
    response.write "        text-align: left; "
    response.write "      }"
    response.write ""
    response.write "      .po-details {"
    response.write "        flex: 2;"
    response.write "        line-height: 0.3125rem;"
    response.write "        font-size: 0.875rem;"
    response.write "        text-align: left; "
    response.write "      }"
    response.write ""
    response.write "      .po-details p {"
    response.write "        display: flex;"
    response.write "        align-items: center;"
    response.write "      }"
    response.write ""
    response.write "      .po-details p label {"
    response.write "        width: 12.5rem;"
    response.write "        display: inline-block;"
    response.write "        font-weight: bold;"
    response.write "      }"
    response.write ""
    response.write "      .po-details p span {"
    response.write "        margin-left: 0.3125rem;"
    response.write "        padding-right: 0.5rem;"
    response.write "      }"
    response.write ""
    response.write "      .table-div {"
    response.write "        width: 100%;"
    response.write "        margin-top: 1rem;"
    response.write "      }"
    response.write ""
    response.write "      .mytable {"
    response.write "        width: 100%;"
    response.write "        border-collapse: collapse;"
    response.write "      }"
    response.write ""
    response.write "      .myth,"
    response.write "      .mytd {"
    response.write "        border: 1px solid black;"
    response.write "        padding: 0.3rem;"
    response.write "        text-align: left;"
    response.write "      }"
    response.write ""
    response.write "      .signatories {"
    response.write "        margin-top: 3rem;"
    response.write "        line-height: 1rem;"
    response.write "        text-align: left; /* Ensure left alignment */"
    response.write "      }"
    response.write ""
    response.write "      .remarks {"
    response.write "        margin-top: 1rem;"
    response.write "        line-height: 1.2rem;"
    response.write "        font-size: 10pt;"
    response.write "      }"
    response.write ""
    response.write "      .remarks p {"
    response.write "        display: flex;"
    response.write "        align-items: flex-start;"
    response.write "        gap: 1rem;"
    response.write "      }"
    response.write ""
    response.write "      .remarks p label {"
    response.write "        flex: 0 0 100px;"
    response.write "        text-align: left;"
    response.write "        font-weight: bold;"
    response.write "      }"
    response.write ""
    response.write "      .remarks p span {"
    response.write "        flex: 1;"
    response.write "        text-align: left;"
    response.write "      }"
    response.write ""
    response.write "      .tdr {"
    response.write "        text-align: right;"
    response.write "      }"
    response.write ""
    response.write "      * {"
    response.write "        overflow: visible !important;"
    response.write "      }"
    response.write ""
    response.write "      @media screen and (max-width: 768px) {"
    response.write "        .main {"
    response.write "          margin: 0.5rem;"
    response.write "          width: 100%;"
    response.write "          height: auto;"
    response.write "        }"
    response.write ""
    response.write "        .header {"
    response.write "          flex-direction: column;"
    response.write "          align-items: center;"
    response.write "          gap: 0.5rem;"
    response.write "          position: static;"
    response.write "        }"
    response.write ""
    response.write "        .image {"
    response.write "          position: static;"
    response.write "          left: 0;"
    response.write "        }"
    response.write ""
    response.write "        .header-text {"
    response.write "          margin-top: 0.5rem;"
    response.write "        }"
    response.write ""
    response.write "        .contact-details {"
    response.write "          margin: 0.5rem 0;"
    response.write "          max-width: 100%;"
    response.write "          text-align: left; "
    response.write "        }"
    response.write ""
    response.write "        .address-and-po-details {"
    response.write "          flex-direction: column;"
    response.write "          gap: 1rem;"
    response.write "        }"
    response.write ""
    response.write "        .address,"
    response.write "        .po-details {"
    response.write "          flex: none;"
    response.write "          width: 100%;"
    response.write "          text-align: left; /* Ensure left alignment on smaller screens */"
    response.write "        }"
    response.write ""
    response.write "        .signatories {"
    response.write "          text-align: left; /* Ensure left alignment on smaller screens */"
    response.write "        }"
    response.write ""
    response.write "        .remarks p {"
    response.write "          flex-direction: column;"
    response.write "          gap: 0.5rem;"
    response.write "        }"
    response.write ""
    response.write "        .remarks p label,"
    response.write "        .remarks p span {"
    response.write "          flex: none;"
    response.write "          width: 100%;"
    response.write "          text-align: left;"
    response.write "        }"
    response.write "      }"
    response.write "    </style>"
End Sub

Sub Printout()

End Sub
    ' Declare variables for SQL and recordset
    Dim rst, sql, tot, drg, qty, ucst, amt, uom, terms

    ' Initialize the recordset
    Set rst = CreateObject("ADODB.Recordset")

    ' Initialize the total
    tot = 0

    response.write "  <div class=""main"">"
    response.write "    <div class=""header"">"
    response.write "      <div class=""image"">"
    response.write "        <img src=""images/letterhead5.jpg"" width=""90"" height=""90"" alt=""Imah Logo"" />"
    response.write "      </div>"
    response.write "      <div class=""header-text"">"
    response.write "        <h1>INTERNATIONAL MARITIME HOSPITAL</h1>"
    response.write "        <h3>Purchase Order No: " & GetRecordField("DrugPurOrderID") & "</h3>"
    response.write "      </div>"
    response.write "    </div>"
    response.write ""
    response.write "    <div class=""contact-details"">"
    response.write "      <p><strong>Telephone</strong> <span>:</span></p>"
    response.write "      <p><strong>Fax</strong> <span>:</span></p>"
    response.write "      <p><strong>Tax registration number</strong> <span>:</span></p>"
    response.write "    </div>"
    response.write ""
    response.write "    <div class=""address-and-po-details"">"
    response.write "      <div class=""address"">"
    response.write "        <p>To: " & GetRecordField("SupplierName") & "</p>"
    response.write "        <p style=""padding-left: 28px"">" & GetComboNameFld("Supplier", GetRecordField("SupplierID"), "Address") & "</p>"
    response.write "        <p style=""padding-left: 28px"">" & GetComboNameFld("Supplier", GetRecordField("SupplierID"), "City") & "</p>"
    'Response.write "        <p style=""padding-left: 28px"">GHA</p>"
    response.write "      </div>"
    response.write "      <div class=""po-details"">"
    
    response.write "        <p><label>Creation date</label> <span>: " & GetRecordField("WorkingDayName") & "</span></p>"
    response.write "        <p>"
    response.write "          <label>Print date/time</label>"
    response.write "          <span>: " & GetRecordField("PurchaseOrderDate") & "</span>"
    response.write "        </p>"
    response.write "        <p>"
    response.write "          <label>Prepayment obligation</label>"
    response.write "          <span>: No</span>"
    response.write "        </p>"
    response.write "        <p><label>Currency</label> <span>: GHS</span></p>"
    response.write "        <p><label>Delivery date</label> <span>: </span></p>"
    response.write "        <p><label>PR Number</label> <span>: </span></p>"
    response.write "        <p><label>RFQ Number</label> <span>: </span></p>"
    response.write "        <p><label>Requisitioner</label> <span>: </span></p>"
    response.write "        <p><label>Created by</label> <span>: " & GetComboName("Staff", GetComboNameFld("SystemUser", GetRecordField("SystemUserID"), "StaffID")) & "</span></p>"
    response.write "      </div>"
    response.write "    </div>"
    response.write ""
    response.write "    <div class=""table-div"">"
    response.write "      <table class=""mytable"">"
    response.write "        <thead>"
    response.write "          <tr>"
    response.write "            <th class=""myth"">Item</th>"
    response.write "            <th class=""myth"">Description</th>"
    response.write "            <th class=""myth"">Qty</th>"
    response.write "            <th class=""myth"">Unit</th>"
    response.write "            <th class=""myth"">Unit Price</th>"
    response.write "            <th class=""myth tdr"">Amount</th>"
    response.write "          </tr>"
    response.write "        </thead>"
    response.write "        <tbody>"

    ' Query 1: Fetch items from drugPurOrderItems
    sql = "select *, DrugPurOrder.KeyPrefix from drugPurOrderItems "
    sql = sql & "Join DrugPurOrder ON DrugPurOrder.DrugPurOrderID = DrugPurOrderItems.DrugPurOrderID "
    sql = sql & "where DrugPurOrderItems.drugpurorderid='" & drugPurOrderID & "' order by drugid"
    
    response.write sql
    With rst
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                drg = .fields("drugid")
                qty = .fields("orderquantity")
                ucst = .fields("orderamount1")
                amt = .fields("orderamount2")
                uom = .fields("unitofmeasureid")
                terms = .fields("KeyPrefix")
                
                tot = tot + amt
                response.write "          <tr>"
                response.write "            <td class=""mytd"">" & UCase(drg) & "</td>"
                response.write "            <td class=""mytd"">" & GetComboName("Drug", drg) & "</td>"
                response.write "            <td class=""mytd tdr"">" & CStr(qty) & "</td>"
                response.write "            <td class=""mytd"">" & GetComboName("UnitOfMeasure", uom) & "</td>"
                response.write "            <td class=""mytd tdr"">" & FormatNumber(ucst, 4, , , -1) & "</td>"
                response.write "            <td class=""mytd tdr"">" & FormatNumber(amt, 2, , , -1) & "</td>"
                response.write "          </tr>"
                .MoveNext
            Loop
        End If
        .Close
    End With

    ' Query 2: Fetch items from drugPurOrderItems2 (By Tender)
    sql = "select *, DrugPurOrder.KeyPrefix from drugPurOrderItems2 "
    sql = sql & "Join DrugPurOrder ON DrugPurOrder.DrugPurOrderID = DrugPurOrderItems2.DrugPurOrderID "
    sql = sql & "where DrugPurOrderItems2.drugpurorderid='" & drugPurOrderID & "' order by drugid"
    
    response.write sql
    With rst
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                drg = .fields("drugid")
                qty = .fields("orderquantity")
                ucst = .fields("orderamount1")
                amt = .fields("orderamount2")
                uom = .fields("unitofmeasureid")
                erms = .fields("KeyPrefix")
                
                tot = tot + amt
                response.write "          <tr>"
                response.write "            <td class=""mytd"">" & UCase(drg) & "</td>"
                response.write "            <td class=""mytd"">" & GetComboName("Drug", drg) & "</td>"
                response.write "            <td class=""mytd tdr"">" & CStr(qty) & "</td>"
                response.write "            <td class=""mytd"">" & GetComboName("UnitOfMeasure", uom) & "</td>"
                response.write "            <td class=""mytd tdr"">" & FormatNumber(ucst, 4, , , -1) & "</td>"
                response.write "            <td class=""mytd tdr"">" & FormatNumber(amt, 2, , , -1) & "</td>"
                response.write "          </tr>"
                .MoveNext
            Loop
        End If
        .Close
    End With

    ' Write the totals
    ' Subtotal (dynamically calculated)
    response.write "          <tr>"
    response.write "            <td class=""mytd""></td>"
    response.write "            <td class=""mytd""></td>"
    response.write "            <td colspan=""3"" class=""mytd tdr"">"
    response.write "              <strong>SubTotal</strong>"
    response.write "            </td>"
    response.write "            <td class=""mytd tdr""><strong> " & FormatNumber(tot, 2, , , -1) & "</strong></td>"
    response.write "          </tr>"
    ' Discount (unchanged)
    response.write "          <tr>"
    response.write "            <td class=""mytd""></td>"
    response.write "            <td class=""mytd""></td>"
    response.write "            <td colspan=""3"" class=""mytd tdr"">"
    response.write "              <strong>Discount</strong>"
    response.write "            </td>"
    response.write "            <td class=""mytd tdr""><strong>0.00</strong></td>"
    response.write "          </tr>"
    ' Msc. Charges (unchanged)
    response.write "          <tr>"
    response.write "            <td class=""mytd""></td>"
    response.write "            <td class=""mytd""></td>"
    response.write "            <td colspan=""3"" class=""mytd tdr"">"
    response.write "              <strong>Msc. Charges</strong>"
    response.write "            </td>"
    response.write "            <td class=""mytd tdr""><strong> 0.00</strong></td>"
    response.write "          </tr>"
    ' Total VAT (unchanged)
    response.write "          <tr>"
    response.write "            <td class=""mytd""></td>"
    response.write "            <td class=""mytd""></td>"
    response.write "            <td colspan=""3"" class=""mytd tdr"">"
    response.write "              <strong>Total VAT</strong>"
    response.write "            </td>"
    response.write "            <td class=""mytd tdr""><strong>0.00</strong></td>"
    response.write "          </tr>"
    ' Total Amount (same as Subtotal since no additional charges)
    response.write "          <tr>"
    response.write "            <td class=""mytd""></td>"
    response.write "            <td class=""mytd""></td>"
    response.write "            <td colspan=""3"" class=""mytd tdr"">"
    response.write "              <strong>Total Amount</strong>"
    response.write "            </td>"
    response.write "            <td class=""mytd tdr""><strong>" & FormatNumber(tot, 2, , , -1) & "</strong></td>"
    response.write "          </tr>"
    response.write "        </tbody>"
    response.write "      </table>"
    response.write "    </div>"
    response.write ""
    response.write "    <div class=""signatories"">"
    response.write "      <p>"
    response.write "        ............................................................................."
    response.write "      </p>"
    response.write "      <p>CEO (International Maritime Hospital)</p>"
    response.write "      <p><strong>Delivery Terms:</strong> <span></span></p>"
    response.write "      <p><strong>Delivery Mode: </strong><span></span></p>"
    response.write "      <p><strong>Payment Terms: </strong><span>Net of 60 days</span></p>"
    response.write "    </div>"
    response.write ""
    
    
    GetTC terms
    
    Sub GetTC(str)
    Dim arr
    If str <> "" Then
        response.write "<table><tbody>"
        For Each Line In Split(str, "|||")
            arr = Split(Line, "::")
            response.write "<tr>"
            response.write "    <td><b><i>" & UCase(Trim(arr(0))) & "</i></b></td>"
            response.write "    <td>" & Trim(arr(1)) & "</td>"
            response.write "</tr>"
        Next
        response.write "</tbody></table>"
    Else
        response.write "    <div class=""remarks"">"
        response.write "      <p><u>Remarks</u></p>"
        response.write "      <p><strong>The Terms and Conditions from IMaH are as follows:</strong></p>"
        response.write "      <p>"
        response.write "        <label><u>WARRANTY:</u></label>"
        response.write "        <span>Goods supplied under this Purchase Order must be covered by warranty and clearly stated on the delivery Documents</span>"
        response.write "      </p>"
        response.write "      <p>"
        response.write "        <label><u>ACCEPTANCE:</u></label>"
        response.write "        <span>Goods delivered are deemed to be accepted only after the IMaH has gone through its internal Goods acceptance processes.</span>"
        response.write "      </p>"
        response.write "      <p>"
        response.write "        <label><u>CANCELLATION:</u></label>"
        response.write "        <span>IMaH reserves the right to cancel the Purchase Order at any time prior to delivery and shall not be subject to any charges or fees as a result of the cancellation</span>"
        response.write "      </p>"
        response.write "      <p>"
        response.write "        <label><u>DELIVERY:</u></label>"
        response.write "        <span>The specific quantity ordered must be delivered in full to the buyers address as stated</span>"
        response.write "      </p>"
        response.write "      <p>"
        response.write "        <label><u>TERMINATION:</u></label>"
        response.write "        <span>The seller must comply with stated delivery date. IMaH reserves the right to cancel the order without notice when the date elapses</span>"
        response.write "      </p>"
        response.write "      <p>"
        response.write "        <label><u>TAXES:</u></label>"
        response.write "        <span>Invoices will be subjected to existing Government of Ghana withholding taxes unless seller provides proof of exemption</span>"
        response.write "      </p>"
        response.write "    </div>"
    End If
End Sub
    
'    Sub GetTC(str)
'    Dim arr
'    If str <> "" Then
'        response.write "<table><tbody>"
'        For Each Line In Split(str, "|||")
'            arr = Split(Line, "::")
'            response.write "<tr>"
'            response.write "    <td><b><i>" & UCase(Trim(arr(0))) & "</i></b></td>"
'            response.write "    <td>" & Trim(arr(1)) & "</td>"
'            response.write "</tr>"
'        Next
'        response.write "</tbody></table>"
'    Else
'        response.write "    <div class=""remarks"">"
'        response.write "      <p><u>Remarks</u></p>"
'        response.write "      <p><strong>The Terms and Conditions from IMaH are as follows:</strong></p>"
'        response.write "      <p>"
'        response.write "        <label><u>WARRANTY:</u></label>"
'        response.write "        <span"
'        response.write "          >Goods supplied under this Purchase Order must be covered by warranty"
'        response.write "          and clearly stated on the delivery Documents</span"
'        response.write "        >"
'        response.write "      </p>"
'        response.write "      <p>"
'        response.write "        <label><u>ACCEPTANCE:</u></label>"
'        response.write "        <span"
'        response.write "          >Goods delivered are deemed to be accepted only after the IMaH has"
'        response.write "          gone through its internal Goods acceptance processes.</span"
'        response.write "        >"
'        response.write "      </p>"
'        response.write "      <p>"
'        response.write "        <label><u>CANCELLATION:</u></label>"
'        response.write "        <span"
'        response.write "          >IMAH reserves the right to cancel the Purchase Order at any time"
'        response.write "          prior to delivery and shall not be subject to any charges or fees as a"
'        response.write "          result of the cancellation</span>"
'        response.write "      </p>"
'        response.write "      <p>"
'        response.write "        <label><u>DELIVERY:</u></label>"
'        response.write "        <span"
'        response.write "          >The specific quantity ordered must be delivered in full to the buyers"
'        response.write "          address as stated</span>"
'        response.write "      </p>"
'        response.write "      <p>"
'        response.write "        <label><u>TERMINATION:</u></label>"
'        response.write "        <span"
'        response.write "          >The seller must comply with stated delivery date. IMaH reserves the"
'        response.write "          right to cancel the order without notice when the date elapse</span>"
'        response.write "      </p>"
'        response.write "      <p>"
'        response.write "        <label><u>TAXES:</u></label>"
'        response.write "        <span"
'        response.write "          >Invoices will be subjected to existing Government of Ghana"
'        response.write "          withholding taxes unless seller provides proof of exemption</span>"
'        response.write "      </p>"
'        response.write "    </div>"
'
'    End If
'    End Sub
    
    
    
  response.write "  </div>"

    ' Clean up the recordset
    Set rst = Nothing
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
