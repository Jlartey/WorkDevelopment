'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

addCSS
processCode
Sub processCode()
  Dim sql, rst, cnt, sponsorID
  Set rst = CreateObject("ADODB.Recordset")
  sponsorID = request.querystring("PrintFilter0")
'  response.write sponsorID

  sql = " SELECT top 10 patientID, initialdependantID FROM InsuredPatient WHERE initialdependantID = 'NONE' and sponsorID = '" & sponsorID & "'"
    
 response.write sql
  cnt = 0

  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
      If .RecordCount > 0 Then
      .MoveFirst

        'response.write "<h3> ANAESTHESIA REPORT </h3>"
        response.write "<table class = 'anaesthesia' > "
        response.write "    <thead> "
        response.write "    <tr class = 'anaesthesia'>"
        response.write "        <th colspan = '13'>List of principals and Dependants for " & GetComboName("Sponsor", sponsorID) & " </th>"
        response.write "    </tr>"
        response.write "    </thead><tbody> "

        
        Do While Not .EOF
          cnt = cnt + 1
          initialdependantID = .fields("initialdependantID")

          response.write "  <tr class = 'queryData'> "
          response.write "      <td>" & (cnt) & "</td> "
          response.write "      <td>" & GetComboName("Patient", .fields("PatientID")) & "</td> "
          response.write "  </tr> "
          getDependants (initialdependantID)
         ' response.flush
          .MoveNext

        Loop
      End If
      response.write "</tbody></table>"
    .Close
    Set rst = Nothing
  End With

End Sub

Sub getDependants(initialdependantID)
    Dim sql, rst
    Set rst = CreateObject("ADODB.RecordSet")
    sql = ""
    sql = "SELECT * FROM insuredPatient WHERE initialdependantID like '%" & initialdependantID & "%'"
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    response.write "<tr class = 'queryData'>"
                    response.write "    <td>&nbsp;</td>"
                    response.write "    <td>" & GetComboName("Patient", .fields("PatientID")) & "</td>"
                    response.write "</tr>"
                    .MoveNext
                Loop
            End If
        .Close
    End With
End Sub

Sub addCSS()
  With response
    .write " <style> "
    .write "    .anaesthesia, .anaesthesia th, .anaesthesia td{ "
    .write "        border: 1px solid silver; "
    .write "        border-collapse: collapse; "
    .write "        padding: 5px; "
    .write "    } "
    .write "    .anaesthesia{ "
    .write "        width: 80vw; "
    .write "        margin: 0 auto; "
    .write "        font-family: sans-serif; "
    .write "        font-size: 13px; "
    .write "        box-sizing: border-box; "
    .write "    }"
    .write "    .anaesthesia tr{page-break-inside:avoid; "
    .write "        page-break-after:auto "
    .write "    } "
    .write "    .anaesthesia th, .anaesthesia td { "
    .write "        border: 1px solid silver; "
    .write "        text-align: center; "
    .write "        padding: 5px; "
    .write "        font-size:13px; "
    .write "        margin: 0 auto; "
    .write "    } "
    .write "    .tHead{ "
    .write "        position: sticky; top: 0; "
    .write "    }  "
    .write "    .queryData td{ "
    .write "        font-size: 12; "
    .write "    }  "
    .write "    .anaesthesia th{ "
    .write "        background-color: blanchedalmond; "
    .write "        text-align: center; "
    .write "        font-weight: bold;"
    .write "        font-size: 14px;color:#000;"
    .write "   } "
    .write " </style> "
  End With
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
