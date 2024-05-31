
Dim vld, txpID, txpBranchID, pgMode

vld = True
txpID = Trim(Request("inpTaxPayerID"))
txpBranchID = Trim(Request("inpTaxPayerBranchID"))

pgMode = UCase(objPage.pagemode)
If pgMode = "EDITMODE" Then
    If HasTaxPayerBranchChanged(txpID, txpBranchID) Then
       vld = False
       SetPageMessages "You cannot edit a Tax Payer's Branch."
    End If
End If

' Will prevent a save 
If Not vld Then
    If objPage.rtnHdlProcessPoint Then
        objPage.hdlProcessPoint = False
    End If
End If

Function HasTaxPayerBranchChanged(txpID, inTxpBranchID)
    Dim txp, ot
    txp = GetComboNameFld("TaxPayer", txpID, "TaxPayerBranchID")
    If txp <> "" Then
        ot = (UCase(inTxpBranchID) <> UCase(txp))
    Else
        ot = True
    End If
    HasTaxPayerBranchChanged = ot
End Function
