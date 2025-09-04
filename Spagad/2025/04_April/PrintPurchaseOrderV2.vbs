Sub Printout()
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
    sql = "SELECT *, DrugPurOrder.KeyPrefix FROM drugPurOrderItems "
    sql = sql & "INNER JOIN DrugPurOrder ON DrugPurOrder.DrugPurOrderID = DrugPurOrderItems.DrugPurOrderID "
    sql = sql & "WHERE DrugPurOrderItems.drugpurorderid='" & Trim(GetRecordField("DrugPurOrderID")) & "' ORDER BY drugid"
    
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
    sql = "SELECT *, DrugPurOrder.KeyPrefix FROM drugPurOrderItems2 "
    sql = sql & "INNER JOIN DrugPurOrder ON DrugPurOrder.DrugPurOrderID = DrugPurOrderItems2.DrugPurOrderID "
    sql = sql & "WHERE DrugPurOrderItems2.drugpurorderid='" & Trim(GetRecordField("DrugPurOrderID")) & "' ORDER BY drugid"
    
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
    
    response.write "  </div>"

    ' Clean up the recordset
    Set rst = Nothing
End Sub