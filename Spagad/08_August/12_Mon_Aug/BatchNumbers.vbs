
Dim sqlDropdown
sqsqlDropdownl = "SELECT DISTINCT InventoryRefNo FROM IncomingDrugItems"

Set rstDropdown = CreateObject("ADODB.Recordset")
rstDropdown.Open sqlDropdown, conn, 3, 4

dropdownOptions = ""

With rstDropdown
  If .RecordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
      optionHTML = "<option value='" & .Fields("InventoryRefNo").Value & "'>" & .Fields("InventoryRefNo").Value & "</option>"
      dropdownOptions = dropdownOptions & optionHTML
      .MoveNext
    Loop
  End If
End With

rstDropdown.Close
Set rstDropdown = Nothing



Dim sqlDropdown
sqlDropdown = "SELECT DISTINCT InventoryRefNo FROM IncomingDrugItems"

Set rstDropdown = CreateObject("ADODB.Recordset")
rstDropdown.Open sqlDropdown, conn, 3, 4

dropdownOptions = ""

With rstDropdown
  If Not .EOF Then
    Do While Not .EOF
      dropdownOptions = dropdownOptions & "<option value='" & .Fields("InventoryRefNo").Value & "'>" & .Fields("InventoryRefNo").Value & "</option>"
      .MoveNext
    Loop
  End If
End With

rstDropdown.Close
Set rstDropdown = Nothing

-- Text
http://192.168.5.11/rmchms01/wpgPrtPrintLayoutAll.asp?PrintLayoutName=DrugStockTakeSheet0&PositionForTableName=DrugStore&DrugStoreID=

