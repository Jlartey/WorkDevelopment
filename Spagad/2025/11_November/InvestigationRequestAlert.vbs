'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

SetupAlert

Sub SetupAlert()
    Dim sql, currDt, hrsDur, sDt, rst, noPat, sql2, ky, cnt, ot
    Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, urHtml
    Dim prtLyoName, testGroupID, fnd
    
    fnd = False
    Set rst = CreateObject("ADODB.Recordset")
    testGroupID = ""
    If UCase(jSchd) = "S13" Or UCase(jSchd) = "DPT005" Then
        prtLyoName = "MonitorVisitationLab"
        testGroupID = "B13"
    ElseIf UCase(jSchd) = "S19" Or UCase(jSchd) = "DPT011" Then
        prtLyoName = "MonitorVisitationRad"
        testGroupID = "B19"
    End If

    If testGroupID <> "" Then
        AddRequests testGroupID, prtLyoName
        response.write PageScripts(testGroupdID)
        response.write PageStyle
    End If
End Sub
Sub AddRequests(testGroupID, prtLyoName)
    Dim sql, currDt, hrsDur, sDt, rst, noPat, sql2, ky, cnt, ot
    Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, urHtml
    
    Set rst = CreateObject("ADODB.Recordset")
    hrsDur = 24
    currDt = Now()
    sDt = DateAdd("h", (-1) * hrsDur, currDt)
    noPat = 0
    
    sql = "select LabByDoctorStatusID, count(LabTestID) as cnt from ("
    sql = sql & " select LabByDoctorStatusID, LabTestID from LabByDoctor"
    sql = sql & " where BranchID='" & brnch & "' and TestGroupID='" & testGroupID & "' and PrescriptionDate > CAST(GETDATE() AS DATE) " ' between '" & FormatDateDetail(sDt) & "' and '" & FormatDateDetail(currDt) & "'"
    sql = sql & " and LabByDoctorStatusID in ('L001', 'L003')" 'pending requests + results entered
    sql = sql & " union all "
    sql = sql & " select 'L003' as LabByDoctorStatusID, LabTestID from investigation "
    sql = sql & " where BranchID='" & brnch & "' and TestGroupID='" & testGroupID & "' and RequestDate > CAST(GETDATE() AS DATE) " ' between '" & FormatDateDetail(sDt) & "' and '" & FormatDateDetail(currDt) & "'"
    sql = sql & " ) as [report]"
    sql = sql & " group by LabByDoctorStatusID order by LabByDoctorStatusID"
    
    With rst
        .open sql, conn, 3, 4
        If .recordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If Not IsNull(.fields("cnt")) Then
                    If IsNumeric(.fields("cnt")) Then
                        ky = .fields("LabByDoctorStatusID")
                        cnt = .fields("cnt")
                        ot = ot & "<div "
                        If cnt > 0 Then
                            If UCase(ky) = "L001" Then
                                ot = ot & " class='blink-red' "
'                            ElseIf UCase(ky) = "L002" Then
'                                ot = ot & " class='blink' "
                            ElseIf UCase(ky) = "L003" Then
                                ot = ot & " class='blink-yellow' "
                            End If
                        End If
                        ot = ot & " style='padding: px;'><label class='label-style'>" & GetComboName("LabByDoctorStatus", ky) & " :</label>&nbsp;&nbsp;<label class='label-style' id='" & ky & "'>" & CStr(cnt) & "</label></div>"
                    End If
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    
    'Clickable Url Link
    
    lnkUrl = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=" & prtLyoName & "&PositionForTableName=WorkingDay&WorkingDayID=DAY20160401"
    lnkText = "Click for Details"
    'urHtml = "<br><a class='label-style' id='anchorElement' href='javascript:window.open(""" & lnkUrl & """, ""_blank"", ""scrollbars=yes"")'>" & lnkText & "</a>"
    urHtml = "<a class='label-style' id='anchorElement' target='_blank' href='" & lnkUrl & "'>" & lnkText & "</a>"
    ot = "<div>" & ot & "</div>" & urHtml
    
    response.write ot
End Sub
Sub AddPatientCount()
    Dim sql, currDt, hrsDur, sDt, rst, noPat, sql2, ky, cnt, ot
    Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, urHtml
    
    Set rst = CreateObject("ADODB.Recordset")
    hrsDur = 24
    currDt = Now()
    sDt = DateAdd("h", (-1) * hrsDur, currDt)
    noPat = 0
    
    sql = "select 'Number Of Patients' AS [ ], count(distinct VisitationID) as cnt from LabByDoctor "
    sql = sql & " where BranchID='" & brnch & "' and TestGroupID='B13' and PrescriptionDate between '" & FormatDateDetail(sDt) & "' and '" & FormatDateDetail(currDt) & "'"

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
    ot = "<b><u><font color=""red"">Lab Request [Past 24 hours]</font></u></b><br>"

    ot = ot & "<b>No. of Patient :</b>&nbsp;&nbsp;" & CStr(noPat) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    
    response.write ot
End Sub
Function PageStyle()
    Dim str
    str = ""
    str = str & vbCrLf & "<style>"
    str = str & vbCrLf & " .label-style{ font-size: 9pt; font-weight: bold;}"
    str = str & vbCrLf & " .blink-green { animation-name: blink-green; animation-duration: 2s; animation-iteration-count: infinite; }"
    str = str & vbCrLf & " .blink-yellow { animation-name: blink-yellow; animation-duration: 2s; animation-iteration-count: infinite; }"
    str = str & vbCrLf & " .blink-red { animation-name: blink-red; animation-duration: 2s; animation-iteration-count: infinite; }"
    str = str & vbCrLf & " @keyframes blink-yellow{ to {color: yellow;}}"
    str = str & vbCrLf & " @keyframes blink-red{ to {color: red;}}"
    str = str & vbCrLf & " @keyframes blink-green { to {color: green;}}"
    str = str & vbCrLf & "</style>"
    
    PageStyle = str
End Function
Function PageScripts(testGroupID)
    Dim str
    str = str & vbCrLf & "<script>"
    str = str & vbCrLf & "  var ele = document.getElementById('trPrintControl'); "
    str = str & vbCrLf & "  if(ele){ ele.style.display='none'; } "
    'str = str & vbCrLf & "  setTimeout(function(){window.location.reload();}, 23500);"
    str = str & vbCrLf & "  ele = document.getElementById('anchorElement');"
    str = str & vbCrLf & "  if(ele) { var x = ele.parentElement; if(x){ x.style.textAlign='left';}} "
    str = str & vbCrLf & "</script>"
    
    str = str & vbCrLf & "<script type='text/javascript'>"
'    str = str & vbCrLf & "  console.log('hello, world'); "
    str = str & vbCrLf & "  var isRunning=false; "
    str = str & vbCrLf & "  function updateRequestValues(){ "
    str = str & vbCrLf & "      if(isRunning) { return; } else{ isRunning=true; } "
    str = str & vbCrLf & "      var url = 'wpgXmlHttp.asp?ProcedureName=GetIncomingInvestigation&TestGroupID=" & testGroupID & "' "
    str = str & vbCrLf & "      Helpers.xmlhttprequest(Helpers.ucase('GET'), url, function(readyState, responseText){ "
    str = str & vbCrLf & "          str = ReplaceXmlHttpComment(responseText); "
'    str = str & vbCrLf & "          console.log(str); "
    str = str & vbCrLf & "          for(elx of str.split('|*|')){ "
'    str = str & vbCrLf & "              console.log(elx); "
    str = str & vbCrLf & "              var x = elx.split('||');"
    str = str & vbCrLf & "              if(x){ "
    str = str & vbCrLf & "                 if(x.length > 1 ){ "
'    str = str & vbCrLf & "                  console.log(x); "
    str = str & vbCrLf & "                  document.getElementById(x[0]).innerText= x[1]; "
    str = str & vbCrLf & "                 } "
    str = str & vbCrLf & "              } "
    str = str & vbCrLf & "              "
    str = str & vbCrLf & "          }"
    str = str & vbCrLf & "          isRunning = false; "
    str = str & vbCrLf & "          "
    str = str & vbCrLf & "      }); "
    str = str & vbCrLf & " } "
    str = str & vbCrLf & " setInterval(updateRequestValues, 10000); "
    str = str & vbCrLf & "</script>"
    
    PageScripts = str
End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
