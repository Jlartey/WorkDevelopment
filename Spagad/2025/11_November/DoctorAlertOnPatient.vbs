Dim topLeftHtmlMsg, topRightHtmlMsg, bottomLeftHtmlMsg, bottomRightHtmlMsg
Dim topLeftTimeOut, topRightTimeOut, bottomLeftTimeOut, bottomRightTimeOut 'Alert Timeout in seconds
Dim lnkCnt
lnkCnt = 0

InitAlert 'Initialize Alert
SetupAlert 'Setup Alert
FinalizeAlert 'Finalize Alert

Sub SetupAlert()
  Dim sql, currDt, hrsDur, sDt, rst, noPat, sql2, ky, cnt, ot, ot2
  Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, urHtml
  Set rst = CreateObject("ADODB.Recordset")
  
  hrsDur = 24
  currDt = Now()
  sDt = DateAdd("h", (-1) * hrsDur, currDt)
  noPat = 0
  
  sql = "select count(VisitationID) as cnt from Visitation "
  sql = sql & " where SpecialistID='" & uname & "' and VisitDate between '" & FormatDateDetail(sDt) & "' and '" & FormatDateDetail(currDt) & "'"
  With rst
    .open sql, conn, 3, 4
    If .recordCount > 0 Then
      .MoveFirst
      If Not IsNull(.fields("cnt")) Then
        If IsNumeric(.fields("cnt")) Then
          noPat = .fields("cnt")
        End If
      End If
    End If
    .Close
  End With
  'If noPat > 0 Then
    ot = "<b><u><font color=""red"">My OPD Patients [Past 24 hours]</font></u></b><br>"
    
    'Clickable Url Link
    lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=MonitorVisitationDoc&PositionForTableName=WorkingDay&WorkingDayID=DAY20160401"
    lnkText = "Click for Details"
    urHtml = "<a href='javascript:window.open(""" & lnkUrl & """, ""_blank"", ""scrollbars=yes"")'>" & lnkText & "</a>"
    
    
    
    ot = ot & "<b>No. of Patient :</b>&nbsp;&nbsp;" & CStr(noPat) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & urHtml
    
    sql = "select MedicalOutComeID,count(VisitationID) as cnt from Visitation"
    sql = sql & " where SpecialistID='" & uname & "' and VisitDate between '" & FormatDateDetail(sDt) & "' and '" & FormatDateDetail(currDt) & "'"
    sql = sql & " group by MedicalOutComeID order by MedicalOutComeID"
    With rst
      .open sql, conn, 3, 4
      If .recordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          If Not IsNull(.fields("cnt")) Then
            If IsNumeric(.fields("cnt")) Then
              ky = .fields("MedicalOutComeID")
              cnt = .fields("cnt")
              ot = ot & "<br><b>" & GetComboName("MedicalOutCome", ky) & " :</b>&nbsp;&nbsp;" & CStr(cnt)
            End If
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
    bottomRightHtmlMsg = "<table border=""0"" cellspacing=""0"" cellpadding=""3"" style=""border-collapse:collapse; border-color:#eeeeee""><tr>"
    bottomRightHtmlMsg = bottomRightHtmlMsg & "<td valign=""top""><div style=""font-size:10pt;background-color:#ffffff"">" & ot & "</div></td>" 'OPD
  'End If
  
    'WARD
    sql = "select count(VisitationID) as cnt from Admission "
    sql = sql & " where MedicalStaff2ID='" & uname & "' and AdmissionStatusID='A001'"
    With rst
      .open sql, conn, 3, 4
      If .recordCount > 0 Then
        .MoveFirst
        If Not IsNull(.fields("cnt")) Then
          If IsNumeric(.fields("cnt")) Then
            noPat = .fields("cnt")
          End If
        End If
      End If
      .Close
    End With
  'If noPat > 0 Then
    ot = "<b><u><font color=""red"">My Current Ward Patients</font></u></b><br>"
    
    'Clickable Url Link
    lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=MonitorVisitationDoc&PositionForTableName=WorkingDay&WorkingDayID=DAY20160401"
    lnkText = "Click for Details"
    urHtml = "" '"<a href='javascript:window.open(""" & lnkUrl & """, ""_blank"", ""scrollbars=yes"")'>" & lnkText & "</a>"
    
    
    
    ot = ot & "<b>No. of Patient :</b>&nbsp;&nbsp;" & CStr(noPat) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & urHtml
    bottomRightHtmlMsg = bottomRightHtmlMsg & "<td valign=""top""><div style=""font-size:10pt;background-color:#ffffff"">" & ot & "</div></td>" 'WARD
    bottomRightHtmlMsg = bottomRightHtmlMsg & "</tr></table>"
  'End If
  Set rst = Nothing
End Sub
Sub InitAlert()
  'Alert Message Html
  topLeftHtmlMsg = ""
  topRightHtmlMsg = ""
  bottomLeftHtmlMsg = ""
  bottomRightHtmlMsg = ""
  
  'Alert Timeout in seconds
  topLeftTimeOut = 60
  topRightTimeOut = 60
  bottomLeftTimeOut = 60
  bottomRightTimeOut = 60
  
  SetPageVariable "UserClientAlertOutput", ""
End Sub

Sub FinalizeAlert()
  Dim ot
  ot = ""
  ot = topLeftHtmlMsg & "3*%+?" & CStr(topLeftTimeOut)
  ot = ot & "2*%+?" & topRightHtmlMsg & "3*%+?" & CStr(topRightTimeOut)
  ot = ot & "2*%+?" & bottomLeftHtmlMsg & "3*%+?" & CStr(bottomLeftTimeOut)
  ot = ot & "2*%+?" & bottomRightHtmlMsg & "3*%+?" & CStr(bottomRightTimeOut)
  
  SetPageVariable "UserClientAlertOutput", ot
End Sub
Function GetUrlLink2(lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth)
  Dim plusMinus, imgName, lnkOpClNavPop, align
   plusMinus = ""
   imgName = ""
   align = ""
   lnkOpClNavPop = inout & "||" & navPop & "||800||600||CLOSE"
  GetUrlLink2 = GetPrtNavLink(lnkID, plusMinus, imgName, lnkText, lnkUrl, lnkOpClNavPop, fntSize, fntColor, bgColor, align, wdth)
End Function


