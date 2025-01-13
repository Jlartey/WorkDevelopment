'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

response.write "<style> table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 80vw;; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{ background-color: #FC046A; color:#fff } table#myTable thead .h_title{ background-color: blanchedalmond; } table#myTable thead .h_names{ font-size: 14px; background-color: blanchedalmond; position: sticky; top: 0;} table#myTable tbody td{text-align:center;} table#myTable .last{background-color: #3C8F6D;color:#fff;font-weight:700;text-align:center;}</style>"
Sub DisplayCorpAttendanceReport(yrs)
  Dim rst, sql, sp, rst1, rst2, tot, cnt
  Dim mth, typ, pos, mthPos, num, gTot
  Dim arrSp(1000)
  Dim arrMth(1000)
  
  Set rst = CreateObject("ADODB.Recordset")
  Set rst1 = CreateObject("ADODB.Recordset")
  Set rst2 = CreateObject("ADODB.Recordset")
  
  'SpecialistType
  pos = 0
  mthPos = 0
  gTot = 0
  sql = "SELECT DISTINCT v.sponsorid as sid, s.sponsorname as sname, v.insurancegroupid as igp FROM visitation AS v join Sponsor AS s ON v.sponsorid=s.sponsorid WHERE v.insurancegroupid='CORP' and WorkingYearID='" & yrs & "' ORDER BY sname"
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      response.write "<table id='myTable'> <thead><tr class='h_res'><th colspan='15'>Found " & rst.RecordCount & " Results...</th></tr> "
      response.write "<tr class='h_title'><th colspan='15'>Generated Attendance Report " & GetComboName("WorkingYear", yrs) & " for CORPORATE COMPANIES</th></tr>"
      'WorkingMonth
      sql = "SELECT DISTINCT workingmonthid FROM Visitation WHERE WorkingYearID='" & yrs & "' ORDER BY workingmonthid"
      rst1.open qryPro.FltQry(sql), conn, 3, 4
      If rst1.RecordCount > 0 Then
        'Header
        response.write "<tr class='h_names'><th>No.</th><th>Corporate Sponsor Names</th>"
          
        rst1.MoveFirst
        Do While Not rst1.EOF
          mth = rst1.fields("WorkingMonthid")
          response.write "<th>" & GetComboName("WorkingMonth", mth) & "</th>"
          rst1.MoveNext
        Loop
        response.write "<th class='last'>Total</th></tr></thead><tbody>"
        
        Do While Not rst.EOF
          typ = rst.fields("sid")
          pos = pos + 1
          response.write "<tr><td>" & CStr(pos) & "</td><td style='text-align:left; align-items:center;'>" & GetComboName("Sponsor", typ) & "</td>"
          mthPos = 0
          tot = 0
          rst1.MoveFirst
          Do While Not rst1.EOF
            mth = rst1.fields("WorkingMonthid")
            cnt = 0
            mthPos = mthPos + 1
            
            ' sql = "select count(Patientid) as cnt from visitation"
            ' sql = sql & " where WorkingYearID='" & yrs & "' and workingMonthid='" & mth & "' and Sponsorid='" & typ & "'"

            sql = "SELECT SUM(dyi.ppl) AS cnt, dyi.workingmonthid FROM ( "
            sql = sql & " SELECT workingdayid, workingmonthid, COUNT(distinct(patientid)) AS ppl FROM Visitation "
            sql = sql & " WHERE insurancegroupid='CORP' AND Sponsorid='" & typ & "' AND visitationid NOT LIKE '%-C' AND workingMonthid='" & mth & "' AND WorkingYearID='" & yrs & "' "
            sql = sql & " GROUP BY workingdayid,workingmonthid "
            sql = sql & " ) AS dyi GROUP BY dyi.workingmonthid ORDER BY dyi.workingmonthid"
            
            rst2.open qryPro.FltQry(sql), conn, 3, 4
            If rst2.RecordCount > 0 Then
              rst2.MoveFirst
              If Not IsNull(rst2.fields("cnt")) Then
                If IsNumeric(rst2.fields("cnt")) Then
                  cnt = rst2.fields("cnt")
                End If
              End If
            End If
            rst2.Close
            
            If cnt > 0 Then
              response.write "<td>" & FormatNumber(CStr(cnt), 0, , , -1) & "</td>"
            Else
              response.write "<td>0</td>"
            End If
            
            tot = tot + cnt
            gTot = gTot + cnt

            'Month Totals
            If IsNumeric(arrMth(mthPos)) Then
              arrMth(mthPos) = arrMth(mthPos) + cnt
            Else
              arrMth(mthPos) = cnt
            End If
            rst1.MoveNext
          Loop
          'Total
          If tot > 0 Then
            response.write "<td>" & FormatNumber(CStr(tot), 0, , , -1) & "</td></tr>"
          Else
            response.write "<td>0</td></tr>"
          End If
          rst.MoveNext
        Loop
      End If
      rst1.Close
      'Final Totals
      response.write "<tr class='last'><td colspan='2'>TOTALS</td>"
      For num = 1 To mthPos
        If IsNumeric(arrMth(num)) Then
          response.write "<td>" & FormatNumber(CStr(arrMth(num)), 0, , , -1) & "</td>"
        Else
          response.write "<td>0</td>"
        End If
        response.write "</td>"
      Next
      response.write "<td>" & FormatNumber(CStr(gTot), 0, , , -1) & "</td></tr>"
      response.write "</tbody></table><br><br><br><br><br>"
    End If
    .Close
  End With
  Set rst = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
End Sub

Sub DisplayINSAttendanceReport(yrs)
  Dim rst, sql, sp, rst1, rst2, tot, cnt
  Dim mth, typ, pos, mthPos, num, gTot
  Dim arrSp(1000)
  Dim arrMth(1000)
  
  Set rst = CreateObject("ADODB.Recordset")
  Set rst1 = CreateObject("ADODB.Recordset")
  Set rst2 = CreateObject("ADODB.Recordset")
  
  'SpecialistType
  pos = 0
  mthPos = 0
  gTot = 0
  sql = "SELECT DISTINCT v.sponsorid as sid, s.sponsorname as sname, v.insurancegroupid as igp FROM visitation AS v join Sponsor AS s ON v.sponsorid=s.sponsorid WHERE v.insurancegroupid='INS' and WorkingYearID='" & yrs & "' ORDER BY sname"
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      response.write "<table id='myTable'> <thead><tr class='h_res'><th colspan='15'>Found " & rst.RecordCount & " Results...</th></tr> "
      response.write "<tr class='h_title'><th colspan='15'>Generated Attendance Report " & GetComboName("WorkingYear", yrs) & " for INSURANCE COMPANIES</th></tr>"
      'WorkingMonth
      sql = "SELECT DISTINCT workingmonthid FROM Visitation WHERE WorkingYearID='" & yrs & "' ORDER BY workingmonthid"
      rst1.open qryPro.FltQry(sql), conn, 3, 4
      If rst1.RecordCount > 0 Then
        'Header
        response.write "<tr class='h_names'><th>No.</th><th>Insurance Sponsor Names</th>"
          
        rst1.MoveFirst
        Do While Not rst1.EOF
          mth = rst1.fields("WorkingMonthid")
          response.write "<th>" & GetComboName("WorkingMonth", mth) & "</th>"
          rst1.MoveNext
        Loop
        response.write "<th class='last'>Total</th></tr></thead><tbody>"
        
        Do While Not rst.EOF
          typ = rst.fields("sid")
          pos = pos + 1
          response.write "<tr><td>" & CStr(pos) & "</td><td style='text-align:left; align-items:center;'>" & GetComboName("Sponsor", typ) & "</td>"
          mthPos = 0
          tot = 0
          rst1.MoveFirst
          Do While Not rst1.EOF
            mth = rst1.fields("WorkingMonthid")
            cnt = 0
            mthPos = mthPos + 1
            
            ' sql = "select count(Patientid) as cnt from visitation"
            ' sql = sql & " where WorkingYearID='" & yrs & "' and workingMonthid='" & mth & "' and Sponsorid='" & typ & "'"

            sql = "SELECT SUM(dyi.ppl) AS cnt, dyi.workingmonthid FROM ( "
            sql = sql & " SELECT workingdayid, workingmonthid, COUNT(distinct(patientid)) AS ppl FROM Visitation "
            sql = sql & " WHERE insurancegroupid='INS' AND Sponsorid='" & typ & "' AND visitationid NOT LIKE '%-C' AND workingMonthid='" & mth & "' AND WorkingYearID='" & yrs & "' "
            sql = sql & " GROUP BY workingdayid,workingmonthid "
            sql = sql & " ) AS dyi GROUP BY dyi.workingmonthid ORDER BY dyi.workingmonthid"
            
            rst2.open qryPro.FltQry(sql), conn, 3, 4
            If rst2.RecordCount > 0 Then
              rst2.MoveFirst
              If Not IsNull(rst2.fields("cnt")) Then
                If IsNumeric(rst2.fields("cnt")) Then
                  cnt = rst2.fields("cnt")
                End If
              End If
            End If
            rst2.Close
            
            If cnt > 0 Then
              response.write "<td>" & FormatNumber(CStr(cnt), 0, , , -1) & "</td>"
            Else
              response.write "<td>0</td>"
            End If
            
            tot = tot + cnt
            gTot = gTot + cnt

            'Month Totals
            If IsNumeric(arrMth(mthPos)) Then
              arrMth(mthPos) = arrMth(mthPos) + cnt
            Else
              arrMth(mthPos) = cnt
            End If
            rst1.MoveNext
          Loop
          'Total
          If tot > 0 Then
            response.write "<td>" & FormatNumber(CStr(tot), 0, , , -1) & "</td></tr>"
          Else
            response.write "<td>0</td></tr>"
          End If
          rst.MoveNext
        Loop
      End If
      rst1.Close
      'Final Totals
      response.write "<tr class='last'><td colspan='2'>TOTALS</td>"
      For num = 1 To mthPos
        If IsNumeric(arrMth(num)) Then
          response.write "<td>" & FormatNumber(CStr(arrMth(num)), 0, , , -1) & "</td>"
        Else
          response.write "<td>0</td>"
        End If
        response.write "</td>"
      Next
      response.write "<td>" & FormatNumber(CStr(gTot), 0, , , -1) & "</td></tr>"
      response.write "</tbody></table><br><br><br><br><br>"
    End If
    .Close
  End With
  Set rst = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
End Sub


Sub DisplayCASHAttendanceReport(yrs)
  Dim rst, sql, sp, rst1, rst2, tot, cnt
  Dim mth, typ, pos, mthPos, num, gTot
  Dim arrSp(1000)
  Dim arrMth(1000)
  
  Set rst = CreateObject("ADODB.Recordset")
  Set rst1 = CreateObject("ADODB.Recordset")
  Set rst2 = CreateObject("ADODB.Recordset")
  
  'SpecialistType
  pos = 0
  mthPos = 0
  gTot = 0
  sql = "SELECT DISTINCT v.sponsorid as sid, s.sponsorname as sname, v.insurancegroupid as igp FROM visitation AS v join Sponsor AS s ON v.sponsorid=s.sponsorid WHERE v.insurancegroupid='CASH' and WorkingYearID='" & yrs & "' ORDER BY sname"
  With rst
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .MoveFirst
      response.write "<table id='myTable'> <thead><tr class='h_res'><th colspan='15'>Found " & rst.RecordCount & " Results...</th></tr> "
      response.write "<tr class='h_title'><th colspan='15'>Generated Attendance Report " & GetComboName("WorkingYear", yrs) & " for PRIVATE CLIENTS</th></tr>"
      'WorkingMonth
      sql = "SELECT DISTINCT workingmonthid FROM Visitation WHERE WorkingYearID='" & yrs & "' ORDER BY workingmonthid"
      rst1.open qryPro.FltQry(sql), conn, 3, 4
      If rst1.RecordCount > 0 Then
        'Header
        response.write "<tr class='h_names'><th>No.</th><th>SPONSOR</th>"
          
        rst1.MoveFirst
        Do While Not rst1.EOF
          mth = rst1.fields("WorkingMonthid")
          response.write "<th>" & GetComboName("WorkingMonth", mth) & "</th>"
          rst1.MoveNext
        Loop
        response.write "<th class='last'>Total</th></tr></thead><tbody>"
        
        Do While Not rst.EOF
          typ = rst.fields("sid")
          pos = pos + 1
          response.write "<tr><td>" & CStr(pos) & "</td><td style='text-align:left; align-items:center;'>" & GetComboName("Sponsor", typ) & "</td>"
          mthPos = 0
          tot = 0
          rst1.MoveFirst
          Do While Not rst1.EOF
            mth = rst1.fields("WorkingMonthid")
            cnt = 0
            mthPos = mthPos + 1
            
            ' sql = "select count(Patientid) as cnt from visitation"
            ' sql = sql & " where WorkingYearID='" & yrs & "' and workingMonthid='" & mth & "' and Sponsorid='" & typ & "' and VisitationID NOT LIKE '%-C%'"

            sql = "SELECT SUM(dyi.cash) AS cnt, dyi.workingmonthid FROM ( "
            sql = sql & " SELECT workingdayid, workingmonthid, COUNT(distinct(patientid)) AS cash FROM Visitation "
            sql = sql & " WHERE insurancegroupid='CASH' AND Sponsorid='" & typ & "' AND visitationid NOT LIKE '%-C' AND workingMonthid='" & mth & "' AND WorkingYearID='" & yrs & "' "
            sql = sql & " GROUP BY workingdayid,workingmonthid "
            sql = sql & " ) AS dyi GROUP BY dyi.workingmonthid ORDER BY dyi.workingmonthid"
            
            rst2.open qryPro.FltQry(sql), conn, 3, 4
            If rst2.RecordCount > 0 Then
              rst2.MoveFirst
              If Not IsNull(rst2.fields("cnt")) Then
                If IsNumeric(rst2.fields("cnt")) Then
                  cnt = rst2.fields("cnt")
                End If
              End If
            End If
            rst2.Close
            
            If cnt > 0 Then
              response.write "<td>" & FormatNumber(CStr(cnt), 0, , , -1) & "</td>"
            Else
              response.write "<td>0</td>"
            End If
            
            tot = tot + cnt
            gTot = gTot + cnt

            'Month Totals
            If IsNumeric(arrMth(mthPos)) Then
              arrMth(mthPos) = arrMth(mthPos) + cnt
            Else
              arrMth(mthPos) = cnt
            End If
            rst1.MoveNext
          Loop
          'Total
          If tot > 0 Then
            response.write "<td>" & FormatNumber(CStr(tot), 0, , , -1) & "</td></tr>"
          Else
            response.write "<td>0</td></tr>"
          End If
          rst.MoveNext
        Loop
      End If
      rst1.Close
      'Final Totals
      response.write "<tr class='last'><td colspan='2'>TOTALS</td>"
      For num = 1 To mthPos
        If IsNumeric(arrMth(num)) Then
          response.write "<td>" & FormatNumber(CStr(arrMth(num)), 0, , , -1) & "</td>"
        Else
          response.write "<td>0</td>"
        End If
        response.write "</td>"
      Next
      response.write "<td>" & FormatNumber(CStr(gTot), 0, , , -1) & "</td></tr>"
      response.write "</tbody></table>"
    End If
    .Close
  End With
  Set rst = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
End Sub

Dim yrs

yrs = Trim(Request.QueryString("printfilter0"))
DisplayCorpAttendanceReport yrs
DisplayINSAttendanceReport yrs
DisplayCASHAttendanceReport yrs

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
'None
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
