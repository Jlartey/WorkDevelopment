'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim pettyCashId, entryNumber, employeeName, department, amount, paymentType, purpose, paymentDate

pettyCashId = Trim(Request.querystring("PettyCashID"))
getPettyCashDetails pettyCashId
MainPage

Sub getPettyCashDetails(pettyCashId)
    Dim sql, sql2, rst
    
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "SELECT PettyCashID, SystemUserID, JobScheduleID, AuthorizedByTypeID, PaymentDetails, "
    sql = sql & "convert(varchar(20), PaymentDate, 106) PaymentDate FROM PettyCash "
    sql = sql & "WHERE PettyCashID = '" & pettyCashId & "'"

    'response.write sql
    
    With rst
        .open sql, conn, 3, 4
        If .RecordCount > 0 Then
            entryNumber = .fields("PettyCashID")
            employeeName = .fields("SystemUserID")
            department = .fields("JobScheduleID")
            paymentType = .fields("AuthorizedByTypeID")
            purpose = .fields("PaymentDetails")
            paymentDate = .fields("PaymentDate")
        End If
    End With
    
    rst.Close
    Set rst = Nothing
    
    Set rst = CreateObject("ADODB.Recordset")
     
    sql2 = "SELECT SUM(amount) amount FROM PettyCashPayment WHERE PettyCashId = '" & pettyCashId & "'"
    
    With rst
        .open sql2, conn, 3, 4
        If Not .EOF Then
            amount = .fields("amount")
            
        End If
    End With
    
    rst.Close
    Set rst = Nothing

End Sub

Sub MainPage()
   Response.Write "<!DOCTYPE html>"
    Response.Write "<html lang=""en"">"
    Response.Write "  <head>"
    Response.Write "    <meta charset=""UTF-8"" />"
    Response.Write "    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />"
    Response.Write "    <title>Request for Payment - Petty Cash</title>"
    Response.Write "    <style>"
    
    Response.Write "      .heading {"
    Response.Write "        text-align: center;"
    Response.Write "        text-transform: uppercase;"
    Response.Write "        margin-bottom: 20px;"
    Response.Write "        font-family: Arial, sans-serif;"
    Response.Write "        display: flex; "
    Response.Write "        justify-content: center; "
    Response.Write "        align-items: center; "
    Response.Write "      }"
    
    Response.Write "      .container {"
    Response.Write "        border: 3px solid black;"
    Response.Write "        padding: 20px 20px;"
    Response.Write "        margin-top: 20px;"
    Response.Write "        width: 90vw;"
    Response.Write "        margin: 0 auto;"
    Response.Write "        font-family: Arial, sans-serif;"
    Response.Write "        line-height: 1.6;"
    Response.Write "      }"
    
    Response.Write "      .row {"
    Response.Write "        display: flex;"
    Response.Write "        margin-bottom: 10px;"
    Response.Write "      }"
    
    Response.Write "      .label {"
    Response.Write "        width: 200px;"
    Response.Write "        margin-right: 20px;"
    Response.Write "        text-align: right;"
    Response.Write "      }"
    
    Response.Write "      .value {"
    Response.Write "        flex: 1;"
    Response.Write "        text-align: left;"
    Response.Write "      }"
    
    Response.Write "      .signature-section {"
    Response.Write "        margin-top: 30px;"
    Response.Write "      }"
    
    Response.Write "      .signature-section .row {"
    Response.Write "        justify-content: space-around;"
    Response.Write "      }"
    
    Response.Write "      .signature {"
    Response.Write "        text-align: center;"
    Response.Write "      }"
    
    Response.Write "      .title {"
    Response.Write "        margin-top: -20px;"
    Response.Write "        margin-bottom: 40px;"
    Response.Write "      }"
    
    
    Response.Write "      .underline {"
    Response.Write "        border-bottom: 1px solid black;"
    Response.Write "        display: inline-block;"
    Response.Write "        width: 12.5rem; "
    Response.Write "      }"

    Response.Write "    </style>"
    Response.Write "  </head>"
    
    Response.Write "  <body>"
    'images/banner1.bmp
    
    Response.Write " <div class=""heading"">"
        Response.Write "<img src=""images/banner1.bmp"" alt=""logo"" />"
        Response.Write "<div style=""margin-left: 1.25rem"">"
            Response.Write " <h3>Foundation of Orthopaedic and Complex Spine</h3> "
            Response.Write " <h3 style=""margin-top: -5px;"">Request for Payment - Petty Cash </h3>"
        Response.Write "</div>"
    Response.Write " </div>"
    
    Response.Write "    <div class=""container"">"
    Response.Write "      <!-- <div style=""display: flex; justify-content: space-between""> -->"
    Response.Write "      <div class=""row"">"
    
    Response.Write "        <span class=""label"">Entry No:</span>"
    Response.Write "        <span class=""value""><strong>" & entryNumber & "</strong></span>"
    
    Response.Write "        <span class=""label"">Date :</span>"
    Response.Write "        <span class=""value"">" & paymentDate & "</span>"
    Response.Write "      </div>"
    Response.Write ""
    Response.Write "      <div class=""row"">"
    Response.Write "        <span class=""label"">EmployeeName:</span>"
    Response.Write "        <span class=""value""><strong>" & employeeName & "</strong></span>"
    Response.Write "      </div>"
    
    Response.Write "      <div class=""row"">"
    Response.Write "        <span class=""label"">Department:</span>"
    Response.Write "        <span class=""value"">" & GetComboName("JobSchedule", department) & "</span>"
    Response.Write "      </div>"
    
    Response.Write "      <div class=""row"">"
    Response.Write "        <span class=""label"">Amount:</span>"
    Response.Write "        <span class=""value""><strong>" & FormatNumber(amount, 2, , , -1) & "</strong></span>"
    Response.Write "      </div>"
    
    Response.Write "      <div class=""row"">"
    Response.Write "        <span class=""label"">Amount (In words):</span>"
    Response.Write "        <span class=""value"">Two Hundred Cedis</span>"
    Response.Write "      </div>"
    Response.Write "      <div class=""row"">"
    Response.Write "        <span class=""label"">Payment Type:</span>"
    Response.Write "        <span class=""value"">" & GetComboName("AuthorizedByType", paymentType) & "</span>"
    Response.Write "      </div>"
    Response.Write "      <div class=""row"">"
    Response.Write "        <span class=""label"">Purpose:</span>"
    Response.Write "        <span class=""value"">"
    Response.Write "         " & purpose & ""
    Response.Write "        </span>"
    Response.Write "      </div>"
'    response.write "      <div class=""row"">"
'    response.write "        <span class=""label"">Receipt Status:</span>"
'    response.write "        <span class=""value"">I would produce receipt/evidence of payment</span>"
'    response.write "      </div>"
    Response.Write ""
    Response.Write "      <div class=""signature-section"">"
    Response.Write "        <div class=""row"">"
    Response.Write "          <div class=""signature"">"
    Response.Write "            <p>Authority : <span class=""underline""></span></p>"
    '<img src=""https://www.badensports.com/cdn/shop/products/SMA-01_2000x.jpg?v=1629409085"" alt=""logo"" width=""50"" height=""50""/>
    Response.Write "            <div class=""title"">(Department Head)</div>"
    Response.Write "          </div>"
    Response.Write "          <div class=""signature"">"
    Response.Write "            <p>Approved : <span class=""underline""></span></p>"
    Response.Write "            <div class=""title"">(Chief Administrative Officer)</div>"
    Response.Write "          </div>"
    Response.Write "        </div>"
    Response.Write "        <div class=""row"">"
    Response.Write "          <div class=""signature"">"
    Response.Write "            <p>Check By : <span class=""underline""></span></p>"
    Response.Write "            <div class=""title"">(Finance Officer)</div>"
    Response.Write "          </div>"
    Response.Write "          <div class=""signature"" style=""margin-top: 20px"">"
    Response.Write "            Claimant :<span class=""underline""></span> "
    Response.Write "          </div>"
    Response.Write "        </div>"
    Response.Write "      </div>"
    Response.Write "    </div>"
    Response.Write "  </body>"
    Response.Write "</html>"
 
End Sub

Function GetSignaturePath(UserName)
    Dim ot, staffid

    staffid = GetComboNameFld("SystemUser", UserName, "StaffID")

    ot = "images/signatures/" & staffid & ".png"
    GetSignaturePath = ot
End Function
Function GetStaffName(UserName)
    GetStaffName = GetComboName("Staff", GetComboNameFld("SystemUser", UserName, "StaffID"))
End Function
Function GetDrugPurOrderProUser(drgPur, stg)
    Dim ot, sql, rst
    sql = "select top 1 * from DrugPurOrderPro where DrugPurOrderID='" & drgPur & "' and TransProcessVal2ID='" & stg & "' order by TransProcessDate1 desc "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        ot = Array(rst.fields("SystemUserID").value, rst.fields("TransProcessDate1").value)
    Else
        ot = Array()
    End If

    GetDrugPurOrderProUser = ot
End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
