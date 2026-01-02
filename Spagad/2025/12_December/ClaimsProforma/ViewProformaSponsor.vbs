'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim vst, invoiceNumber, invoiceDate, purpose, company
vst = Request.QueryString("PrintFilter1")

' Invoice specific variables
invoiceNumber = "RMC/INV/MS/" & Right(FormatDate(Now()), 2) & "/" & Right("0" & Month(Now()), 2) & "/" & Right("000" & Day(Now()), 3)
invoiceDate = GetComboNameFld("PerformVar22", vst, "KeyPrefix")
purpose = "MEDICAL SCREENING"
company = GetComboName("PerformVar22", vst)
'response.write startDate & " " & endDate


addCSS
header
preview

Sub header()
response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""width:100%; height:200px;"">"
response.write "<tr>"
response.write "<td><img src=""images/letterhead5.jpg"" style=""width:100%; display:none;""></td>"
'AddReportHeader

response.write "</tr>"

response.write "</table>"
End Sub

Sub addCSS()
  With response
    .write vbCrLf & "<style>"
    .write vbCrLf & "  #container{ width: 720px; margin: 0 auto; text-align:left; padding:0;margin:0;font-size:16px; }"
    .write vbCrLf & "  #rptDescription{ display: flex; justify-content: space-between; font-size:14px; margin-bottom: 20px; padding: 10px 0; }"
    .write vbCrLf & "  #rptDescription table { border-collapse: collapse; }"
    .write vbCrLf & "  #rptDescription table td { padding: 3px 8px; vertical-align: top; }"
    .write vbCrLf & "  #description table tr td { padding: 5px; cellpadding:0;cellspacing:0;font-size:16px;}" 'display: flex; justify-content: normal; gap: 20px; line-height: 0.8; }"
    .write vbCrLf & "  #itemTable{ width: 100%; cellpadding:0;font-size:16px;}"
    .write vbCrLf & "  #itemTable, #itemTable th, #itemTable td{ border: 1px solid #000; border-collapse: collapse; text-transform: uppercase; padding: 3px; font-size:14px; }"
    .write vbCrLf & "  #signatories{ display: flex; justify-content: space-between; padding:0; margin:0; }"
    .write vbCrLf & "  .signature-container { width: 170px; height: 40px; border-top: 2px dotted #444; padding:0; margin:0; }"
    .write vbCrLf & "  .table-details { width: 140px; } .signature-container p {margin:4px;}"
    .write vbCrLf & "  p { padding:0; }"
    .write vbCrLf & "</style>"
  End With
End Sub


Sub preview()
'  patDiag = ExtractDiagnosis
'   refNumber = "-" 'genRefNumber()
   refNumber = genRefNumber()
  With response
  .write vbCrLf & "<div id='container'>"
  .write vbCrLf & "  <h3 style='text-decoration: underline; text-align:left'><strong>INVOICE</strong></h3>"
  .write vbCrLf & "  <header id='rptDescription'>"
  .write vbCrLf & "    <div style='display: flex; justify-content: space-between;'>"
  .write vbCrLf & "      <table>"
  .write vbCrLf & "        <tr>"
  .write vbCrLf & "          <td><strong>INVOICE NO.:</strong></td>"
  .write vbCrLf & "          <td>" & invoiceNumber & "</td>"
  .write vbCrLf & "        </tr>"
  .write vbCrLf & "        <tr>"
  .write vbCrLf & "          <td><strong>INVOICE DATE:</strong></td>"
  .write vbCrLf & "          <td>" & FormatDate(invoiceDate) & "</td>"
  .write vbCrLf & "        </tr>"
  .write vbCrLf & "        <tr>"
  .write vbCrLf & "          <td><strong>PURPOSE:</strong></td>"
  .write vbCrLf & "          <td>" & purpose & "</td>"
  .write vbCrLf & "        </tr>"
  .write vbCrLf & "        <tr>"
  .write vbCrLf & "          <td><strong>COMPANY:</strong></td>"
  .write vbCrLf & "          <td>" & UCase(company) & "</td>"
  .write vbCrLf & "        </tr>"
  .write vbCrLf & "      </table>"
  .write vbCrLf & "    </div>"
  .write vbCrLf & "  </header>"
  .write vbCrLf & "  <article>"
  .write vbCrLf & "    <table id='itemTable'>"
  .write vbCrLf & "      <tr>"
  .write vbCrLf & "        <th>NO.</th>"
  .write vbCrLf & "        <th>DESCRIPTION</th>"
  .write vbCrLf & "        <th style='text-align: right'>UNIT PRICE (GHS)</th>"
  .write vbCrLf & "        <th style='text-align: right'>QUANTITY</th>"
  .write vbCrLf & "        <th style='text-align: right'>TOTAL AMOUNT (GHS)</th>"
  .write vbCrLf & "      </tr>"
  generateMedCost
  .write vbCrLf & "    </table>"
  .write vbCrLf & "  </article>"
  .write vbCrLf & "  <aside style='margin-bottom: 50px;padding:0;'>"
  ' .write vbCrLf & "    <p><strong>PLEASE NOTE:</strong></p>"
  ' .write vbCrLf & "    <ul>"
  ' .write vbCrLf & "      <li>This is not an invoice.</li>"
  ' .write vbCrLf & "      <li>Cost of Treatment is subject to change post operatively</li>"
  ' .write vbCrLf & "    </ul>"
  ' .write vbCrLf & "    <p><strong>Thank you.</strong></p>"
  .write vbCrLf & "  </aside>"
  .write vbCrLf & "  <footer id='signatories'>"
  .write vbCrLf & "    <div class='signature-container'>"
  .write vbCrLf & "      <p>TITUS TWUMASI SEDDOH</p>"
  .write vbCrLf & "      <p>FINANCE MANAGER</p>"
  .write vbCrLf & "    </div>"
  .write vbCrLf & "    <div class='signature-container'>"
  .write vbCrLf & "      <p>BELINDA ASANTE AMANKWA ADDO</p>"
  .write vbCrLf & "      <p>ADMINISTRATOR</p>"
  .write vbCrLf & "    </div>"
  .write vbCrLf & "  </footer>"
  .write vbCrLf & "</div>"
  End With
End Sub

Sub generateMedCost()
  Dim sql, rst, cnt, description, descriptionID, amt, amt1, AmtTot, arr, KeyPrefix, fee, itemType, itemID, itemPrice, itemName
  Set rst = CreateObject("ADODB.RecordSet")
  sql = "SELECT * FROM PerformVar58"
  sql = sql & " WHERE PerformVar58Name = '" & vst & "'"
  cnt = 0
  AmtTot = 0

  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          cnt = cnt + 1
          description = .fields("Description")
          KeyPrefix = .fields("KeyPrefix")
          
          ' Parse the description to get item details (e.g., "drg-CVD-6300-0100-00||196.71")
          arr = Split(description, "||")
          If UBound(arr) >= 1 Then
            descriptionID = arr(0)  ' e.g., "drg-CVD-6300-0100-00"
            itemPrice = CDbl(arr(1))  ' e.g., 196.71
          End If
          
          ' Parse KeyPrefix to get quantity (e.g., "1||196.71")
          If InStr(KeyPrefix, "||") > 0 Then
            KeyPrefixArr = Split(KeyPrefix, "||")
            If UBound(KeyPrefixArr) >= 0 Then
              amt = CDbl(KeyPrefixArr(0))  ' quantity
            End If
          End If
          
          ' Extract item type and ID from descriptionID
          Dim dashPos
          dashPos = InStr(descriptionID, "-")
          If dashPos > 0 Then
            itemType = Left(descriptionID, dashPos - 1)  ' drg, itm, lab, trt
            itemID = Mid(descriptionID, dashPos + 1)     ' the rest as ID
          Else
            itemType = "unknown"
            itemID = descriptionID
          End If
          
          ' Get the item name based on type
          Select Case LCase(itemType)
            Case "drg"
              itemName = GetComboName("Drug", itemID)
            Case "itm"
              itemName = GetComboName("Items", itemID)
            Case "lab"
              itemName = GetComboName("LabTest", itemID)
            Case "trt"
              itemName = GetComboName("Treatment", itemID)
            Case Else
              itemName = itemID
          End Select
          
          ' Calculate total fee
          fee = amt * itemPrice
          AmtTot = AmtTot + fee
          
          response.write vbCrLf & "      <tr>"
          response.write vbCrLf & "        <td>" & cnt & "</td>"
          response.write vbCrLf & "        <td>" & itemName & "</td>"
          response.write vbCrLf & "        <td style='text-align: right'>" & FormatNumber(CStr(itemPrice), 2, , , -1) & "</td>"
          response.write vbCrLf & "        <td style='text-align: right'>" & amt & "</td>"
          response.write vbCrLf & "        <td style='text-align: right'>" & FormatNumber(CStr(fee), 2, , , -1) & "</td>"
          response.write vbCrLf & "      </tr>"
          .MoveNext
        Loop
        
        ' Calculate discount (15% as shown in the image)
        Dim discountPercent, discountAmount, netAmount
        discountPercent = GetComboNameFld("Performvar22", vst, "Description")
'        discountPercent = 15
        If Len(discountPercent) > 0 Then
            discountAmount = AmtTot * (CStr(discountPercent) / 100)
        Else
            discountAmount = 0
        End If
        netAmount = AmtTot - discountAmount
        
        response.write vbCrLf & "      <tr>"
        response.write vbCrLf & "        <td colspan='4' style='text-align: right;'><strong>SUB-TOTAL</strong></td>"
        response.write vbCrLf & "        <td style='text-align: right'>" & FormatNumber(CStr(AmtTot), 2, , , -1) & "</td>"
        response.write vbCrLf & "      </tr>"
        response.write vbCrLf & "      <tr>"
        response.write vbCrLf & "        <td colspan='4' style='text-align: right;'><strong>DISCOUNT (" & discountPercent & "%)</strong></td>"
        response.write vbCrLf & "        <td style='text-align: right'>" & FormatNumber(CStr(discountAmount), 2, , , -1) & "</td>"
        response.write vbCrLf & "      </tr>"
        response.write vbCrLf & "      <tr>"
        response.write vbCrLf & "        <td colspan='4' style='text-align: right;'><strong>NET AMOUNT PAYABLE</strong></td>"
        response.write vbCrLf & "        <td style='text-align: right'>" & FormatNumber(CStr(netAmount), 2, , , -1) & "</td>"
        response.write vbCrLf & "      </tr>"
      End If
    .Close
  End With
End Sub

Function ExtractDiagnosis()
  Dim rst, rst2, sql, emrCmp, emrReq, ky, emrCol, emrDate, sUsr, mSf, cnt
  Dim arr, ul, num, arr2, ul2, num2, otVal, itm, dsc1, dsc2, searchCmp, searchCl, extrDat
  Dim dTyp, emrDat, diagArr, diagnosis

  Set rst = CreateObject("ADODB.Recordset")
  cnt = 0
  totCnt = 0
  gCnt = 0
  
  With rst
      
      emrCmp = " 'TH06008', 'TH06008.1' " 'Diagnosis Comp
      emrCol = "Column2" 'Diagnosis Field

      sql = "select er." & emrCol & " from emrresults as er "
      sql = sql & " left join emrrequest as eq on eq.emrrequestid = er.emrrequestid"
      sql = sql & " where eq.visitationid = '" & vst & "' and er.emrcomponentid IN (" & emrCmp & ")"

      .open qryPro.FltQry(sql), conn, 3, 4
      
      If .RecordCount > 0 Then
      .MoveFirst
      diagnosis = ""
      Do While Not .EOF
        If Not IsNull(.fields(emrCol)) Then
            otVal = Trim(.fields(emrCol))
            If Len(otVal) > 0 Then
                            
                'Final Diag
                searchCmp = "EMRVar2B-TH06500301"
                searchCl = "Column1"
                dTyp = "Final.Diag"
                extrDat = ExtractCheckDetCmp(otVal, searchCmp, searchCl, 0)
                diagArr = Split(extrDat, "||")
                ul = UBound(diagArr)
                For num = 0 To ul
                  If Len(diagnosis) > 0 Then
                    diagnosis = diagnosis & ", "
                  End If
                  diagnosis = diagnosis & GetComboName("disease", diagArr(num))
                Next
            End If
        End If
        .MoveNext
      Loop
    End If
      .Close
  End With
  Set rst = Nothing
  ExtractDiagnosis = diagnosis
End Function

Function GetEMRfld(vst, CompID, column)

    Dim sql, rst
    Set rst = server.CreateObject("ADODB.Recordset")
    GetEMRfld = ""
    sql = "SELECT EMRRequestID FROM EMRRequestItems WHERE EMRDataID = 'E000038' AND VisitationID = '" & vst & "'"
    sql = sql & " ORDER BY EMRDate DESC"
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
          .MoveFirst
          GetEMRfld = getEMRResult(.fields("EMRRequestID"), "E000038", CompID, column)
        End If
    End With
    Set rst = Nothing
    
End Function

Function getEMRResult(emrrequestid, emrDataID, CompID, column)

    Dim sql, rst
    Set rst = server.CreateObject("ADODB.Recordset")
    getEMRResult = ""
    
    sql = "SELECT * FROM emrresults WHERE emrrequestid ='" & emrrequestid & "'"
    sql = sql & " AND emrdataid ='" & emrDataID & "' AND emrcomponentid='" & CompID & "'"
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            If Not IsNull(.fields(column)) Then
                getEMRResult = Trim(.fields(column))
            End If
            If IsNull(.fields(column)) Then
                getEMRResult = "Null"
            End If
        End If
    End With
    Set rst = Nothing
    
End Function

Function genRefNumber()

    Dim sql, rst, yr
    Set rst = server.CreateObject("ADODB.Recordset")
    yr = Right(FormatDate(Now()), 2)
    genRefNumber = "AD/" & yr & "-PA-RMC/00"
    cnt = 0
    
    sql = "SELECT distinct patientid, visitationid From EMRRequestItems"
    sql = sql & " WHERE emrdataid='E000038'"
    sql = sql & " AND EMRDate BETWEEN '" & startDate & "' AND '" & endDate & "'"
    sql = sql & " ORDER BY patientid"
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
            If .RecordCount > 0 Then
                .MoveFirst
                    Do While Not .EOF
                        cnt = cnt + 1
                        If UCase(pat) = UCase(.fields("PatientID")) Then
                            genRefNumber = genRefNumber & cnt
                            Exit Do
                        End If
                        .MoveNext
                    Loop
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

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
