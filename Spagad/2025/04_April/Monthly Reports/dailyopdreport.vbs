'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim arPeriod, periodStart, periodEnd, mth
server.scripttimeout = 1800
mth = Request.QueryString("PrintFilter")

periodStart = mth
periodEnd = mth
GenerateCorpAttendanceReport periodStart, periodEnd


response.write "<style> table#myTable, table#myTable th, table#myTable td{border: 1px solid silver; border-collapse: collapse; padding: 5px;} table#myTable{width: 80vw;; margin: 0 auto; font-size: 13px; font-family: sans-serif; box-sizing: border-box; } table#myTable thead{ text-align: center; } table#myTable thead th{padding: 4px;} table#myTable thead .h_res{ background-color: #FC046A; color:#fff } table#myTable thead .h_title{ background-color: blanchedalmond; } table#myTable thead .h_names{ font-size: 14px;} table#myTable tbody td{text-align:center;} table#myTable .last{background-color: #3C8F6D;color:#fff;font-weight:700;text-align:center;}  </style>"

Sub GenerateCorpAttendanceReport(periodStart, periodEnd)
  Dim sql, rst, cnt, rst0, sto, fems, mens, mensTot, femsTot, firstVisit, firstVisitTot
  Dim cash, ins, corp, cashTot, insTot, corpTot, olds, news, susequentVisit, susequentVisitTot
  Set rst0 = CreateObject("ADODB.RecordSet")
  Set rst = CreateObject("ADODB.RecordSet")
  
  sql = "SELECT DISTINCT workingdayid FROM Visitation"
  sql = sql & " WHERE workingmonthid between '" & periodStart & "' and '" & periodEnd & "'"
  sql = sql & " ORDER BY workingdayid"
  rst0.open qryPro.FltQry(sql), conn, 3, 4
  If rst0.RecordCount > 0 Then
    rst0.MoveFirst
    cnt = 0

    response.write "<table id='myTable'> <thead><tr class='h_res'><th colspan='10'>Found " & rst0.RecordCount & " Results...</th></tr> "
    response.write "<tr class='h_title'><th colspan='10'>Generated Attendance Report</th></tr>"
    response.write "<tr class='h_names'><th>No.</th><th>Working Day</th><th>Males</th><th>Females</th><th>Private</th><th>Company</th><th>Insurance</th><th>New</th><th>Old</th><th class='last'>Total</th></tr></thead><tbody>"
        
    Do While Not rst0.EOF
      sto = rst0.fields("workingdayid")
      
      sql = "select count(distinct(patientid)) as fems from visitation"
      sql = sql & " where patientID <> 'P1' AND visitationid NOT LIKE '%-C%' AND workingdayid='" & sto & "' and genderid='GEN02'"
        
      rst.open qryPro.FltQry(sql), conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          If Not IsNull(.fields("fems")) Then
            If IsNumeric(.fields("fems")) Then
              fems = CDbl(.fields("fems"))
              femsTot = femsTot + fems
            End If
          End If
        End If
        rst.Close
      End With
      
      sql = "select count(distinct(patientid)) as gents from visitation"
      sql = sql & " where patientID <> 'P1' AND visitationid NOT LIKE '%-C%' AND workingdayid='" & sto & "' and genderid='GEN01'"
      rst.open qryPro.FltQry(sql), conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          If Not IsNull(.fields("gents")) Then
            If IsNumeric(.fields("gents")) Then
              mens = CDbl(.fields("gents"))
              mensTot = mensTot + mens
            End If
          End If
        End If
        rst.Close
      End With

      sql = "SELECT COUNT(distinct(patientid)) AS cash FROM Visitation"
      sql = sql & " WHERE patientID <> 'P1' AND visitationid NOT LIKE '%-C%' AND workingdayid='" & sto & "' AND insurancegroupid='CASH' AND visitationid NOT LIKE '%-C'"
      rst.open qryPro.FltQry(sql), conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          If Not IsNull(.fields("cash")) Then
            If IsNumeric(.fields("cash")) Then
              cash = CDbl(.fields("cash"))
              cashTot = cashTot + cash
            End If
          End If
        End If
        rst.Close
      End With

      sql = "SELECT COUNT(distinct(patientid)) AS corp FROM Visitation"
      sql = sql & " WHERE patientID <> 'P1' AND visitationid NOT LIKE '%-C%' AND workingdayid='" & sto & "' AND insurancegroupid='CORP'"
      rst.open qryPro.FltQry(sql), conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          If Not IsNull(.fields("corp")) Then
            If IsNumeric(.fields("corp")) Then
              corp = CDbl(.fields("corp"))
              corpTot = corpTot + corp
            End If
          End If
        End If
        rst.Close
      End With

      sql = "SELECT COUNT(distinct(patientid)) AS ins FROM Visitation"
      sql = sql & " WHERE patientID <> 'P1' AND visitationid NOT LIKE '%-C%' AND workingdayid='" & sto & "' AND insurancegroupid='INS'"
      rst.open qryPro.FltQry(sql), conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          If Not IsNull(.fields("ins")) Then
            If IsNumeric(.fields("ins")) Then
              ins = CDbl(.fields("ins"))
              insTot = insTot + ins
            End If
          End If
        End If
        rst.Close
      End With

      sql = "SELECT COUNT(distinct(patientid)) AS firstVisit FROM Visitation"
      sql = sql & " WHERE workingdayid='" & sto & "' AND visittypeID='v001' AND patientID <> 'P1' AND visitationid NOT LIKE '%-C'"
      rst.open qryPro.FltQry(sql), conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          If Not IsNull(.fields("firstVisit")) Then
            If IsNumeric(.fields("firstVisit")) Then
              firstVisit = CDbl(.fields("firstVisit"))
              firstVisitTot = firstVisitTot + firstVisit
            End If
          End If
        End If
        rst.Close
      End With

      sql = "SELECT COUNT(distinct(patientid)) AS susequentVisit FROM Visitation"
      sql = sql & " WHERE workingdayid='" & sto & "' AND visittypeID<>'v001' AND patientID <> 'P1' AND visitationid NOT LIKE '%-C'"
      rst.open qryPro.FltQry(sql), conn, 3, 4
      With rst
        If .RecordCount > 0 Then
          .MoveFirst
          If Not IsNull(.fields("susequentVisit")) Then
            If IsNumeric(.fields("susequentVisit")) Then
              susequentVisit = CDbl(.fields("susequentVisit"))
              susequentVisitTot = susequentVisitTot + susequentVisit
            End If
          End If
        End If
        rst.Close
      End With

      'adding both genders to get total
      ot = mens + fems
      sTot = mensTot + femsTot
      If ot > 0 Then
        cnt = cnt + 1
        response.write "<tr><td>" & CStr(cnt) & "</td> <td>" & GetComboName("WorkingDay", sto) & "</td> <td>" & mens & "</td><td>" & fems & "</td><td>" & cash & "</td><td>" & corp & "</td><td>" & ins & "</td><td>" & firstVisit & "</td><td>" & susequentVisit & "</td>  <td class='last'>" & ot & "</td></tr>"
      End If
      rst0.MoveNext
    Loop
    
    response.write "<tr class='last'><td></td><td>Total</td><td>" & mensTot & "</td><td>" & femsTot & "</td><td>" & cashTot & "</td><td>" & corpTot & "</td><td>" & insTot & "</td><td>" & firstVisitTot & "</td><td>" & susequentVisitTot & "</td><td>" & sTot & "</td></tr></tbody></table>"
  End If
  rst0.Close
  Set rst = Nothing
  Set rst0 = Nothing
End Sub
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>
'None
'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
