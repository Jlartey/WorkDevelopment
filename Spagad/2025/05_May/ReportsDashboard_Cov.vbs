'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim href, rptName

response.write "<style>" & _
    "a{text-decoration:none;font-family:arial;font-size:14px;color:#4169E1 }" & _
    ".content:hover a{color:#696969;text-decoration:none;}" & _
    ".content:hover {background-color:#ffffff;}" & _
    ".heading{font-family:arial;font-size:18px;background-color:#A9A9A9;border-radius:5px;color:#303030}" & _
    ".reportName{font-family:arial;font-size:20px;}" & _
    ".content{background-color:#DCDCDC;border-radius:5px;text-align:center;} " & _
    ".report{border-collapse: separate;}" & _
    ".admission{padding-top:40px;}" & _
    ".newCard{padding:2rem;}" & _
    ".card-deck{display:flex; align-items: center; justify-content: center;}" & _
    ".card{flex:1; margin: 0.5rem;}" & _
    "</style>"
     InitPageScript

  Sub Card(cardColor, cardHeader, cardTitle)
 response.write "<div class='card " & cardColor & " mb-3' style='max-width: 18rem;'>"
  response.write "<div class='card-header'>" & cardHeader & "</div>"
  response.write "<div class='card-body'>"
    response.write "<h5 class='card-title'>" & cardTitle & "</h5>"
  response.write "</div>"
response.write "</div>"
 End Sub
 
 Function GetStat(req)
    Dim ot, rst, sql, adm, wrd, md, sDt, eDt
    Set rst = CreateObject("ADODB.Recordset")
    sDt = FormatDate(Now()) & " 00:00:00"
    eDt = FormatDate(Now()) & " 23:59:59"
    ot = 0
        With rst
            sql = "select distinct "
            If req = "visit" Then
            sql = sql & " * from visitation where visitdate "
            ElseIf req = "admit" Then
            sql = sql & " * from admission where admissiondate "
            ElseIf req = "death" Then
            sql = sql & " visitationid from emrrequest emrreq, emrresults emrres where "
            sql = sql & " emrreq.emrrequestid = emrres.emrrequestid and emrres.emrdataid = 'emr141'"
            sql = sql & " and emrreq.emrdate "
            End If
            sql = sql & " between '" & sDt & "' and '" & eDt & "' "
            .open qryPro.FltQry(sql), conn, 3, 4
            If .RecordCount > 0 Then
                GetStat = .RecordCount
            Else
                GetSat = 0
            End If
            .Close
        End With
    Set rst = Nothing
End Function

response.write "<br><br>"
response.write "<div class='card-deck'>"

Card "bg-warning", "Today's Admission", GetStat("admit")

Card "bg-info", "Today's Attendance", GetStat("visit")

Card "bg-danger", "Today's Death", GetStat("death")
response.write "</div>"


response.write "<table class='report' cellpadding='8' cellspacing='2'>"
response.write "<tr><td class='reportName' colspan='100'><center>REPORTS</center></td></tr>"
response.write "<br/>"

    DisplayReports "Registration"
    DisplayReports "Attendance"
    DisplayReports "Accounts"
    DisplayBrowseView "Visitation"
response.write "</table>"

 

Sub DisplayReports(reportClass)
Dim sql, rst, availableFilter
Server.scripttimeout = 3600

Set rst = Server.CreateObject("ADODB.RecordSet")
    
 sql = "SELECT * from printlayout where description like '%" & reportClass & "%' "
      
With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
        response.write "<tr>"
        response.write "<td class='heading' colspan='100'><center>" & reportClass & "</center></td><td> </td>"
        response.write "</tr>"
                .MoveFirst
                Do While Not .EOF
                availableFilter = .fields("PrintInputFilter")
                If Len(availableFilter) > 1 Then
                    href = "wpgPrtPrintInputFilter.asp?PrintLayoutName=" & .fields("printlayoutid") & "&PositionForTableName=WorkingDay&WorkingDayID="
                Else
                    href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=" & .fields("printlayoutid") & "&PositionForTableName=WorkingDay&WorkingDayID="
                End If
                rptName = .fields("printlayoutname")
                response.write "<td class='content'><a href=""javascript:window.open('" & href & "', 'newwin', 'resizable'); void(0);"">" & rptName & "</a></td>"
                                cnt = cnt + 1
                                If cnt Mod 3 = 0 Then
                                    response.write "</tr>"
                                End If
                        .MoveNext
                Loop
        End If
        .Close
End With

End Sub
Sub DisplayBrowseView(browseViewClass)
Dim sql, rst, availableFilter
Server.scripttimeout = 3600

Set rst = Server.CreateObject("ADODB.RecordSet")
    
 sql = "SELECT * from browseview where reportinfo4 like '%" & browseViewClass & "%' "
      
With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
        response.write "<tr>"
        'response.write "<td class='heading' colspan='100'><center>" & reportClass & "</center></td><td> </td>"
        response.write "<td class='heading' colspan='100'><center>" & browseViewClass & "</center></td><td> </td>"  '@DanielTetteh
        response.write "</tr>"
                .MoveFirst
                Do While Not .EOF
                'href = "wpgbrowseViewLayout.asp?BrowseviewName=" & .fields("printlayoutid") & "&PositionForTableName=WorkingDay&WorkingDayID="
                href = "wpgbrowseViewLayout.asp?BrowseviewName=" & .fields("browseviewid") & "&PositionForTableName=WorkingDay&WorkingDayID=" '@DanielTetteh
                rptName = .fields("browseviewname")
                response.write "<td class='content'><a href=""javascript:window.open('" & href & "', 'newwin', 'resizable'); void(0);"">" & rptName & "</a></td>"
                                cnt = cnt + 1
                                If cnt Mod 3 = 0 Then
                                    response.write "</tr>"
                                End If
                        .MoveNext
                Loop
        End If
        .Close
End With

End Sub
Sub DisplayReports2()
    Dim sql, rst, availableFilter, rptName
    Dim cnt, reportCount

    Server.scripttimeout = 3600

    Set rst = Server.CreateObject("ADODB.RecordSet")

    sql = "SELECT printlayoutid, printlayoutname, description FROM printlayout WHERE LEN(Description) > 1 "
    sql = sql & " and PrintLayoutID in (select PrintLayoutID from PrintOutAlloc poa where poa.JobScheduleID = 'systemadmin') "
    sql = sql & " GROUP BY Description, printlayoutname, PrintLayoutID"
    
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If .fields("description") <> rptName Then
                    If rptName <> "" Then
                        response.write "</tr></table>"
                    End If
                    rptName = .fields("description")
                    response.write "<table>"
                    response.write "<tr><td class='heading' colspan='3'><center>" & rptName & "</center></td></tr>"
                    cnt = 0
                End If
                If cnt Mod 3 = 0 Then
                    response.write "<tr>"
                End If

                If Len(availableFilter) > 1 Then
                    href = "wpgPrtPrintInputFilter.asp?PrintLayoutName=" & .fields("printlayoutid") & "&PositionForTableName=WorkingDay&WorkingDayID="
                Else
                    href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=" & .fields("printlayoutid") & "&PositionForTableName=WorkingDay&WorkingDayID="
                End If

                response.write "<td class='content'><a href=""javascript:window.open('" & href & "', 'newwin', 'resizable'); void(0);"">" & .fields("printlayoutname") & "</a></td>"
                cnt = cnt + 1

                If cnt Mod 3 = 0 Then
                    response.write "</tr>"
                End If

                .MoveNext
            Loop
            response.write "</tr></table>"
        End If
        .Close
    End With
End Sub
Sub InitPageScript()
  htStr = ""
  htStr = htStr & "<meta name=""viewport"" content=""width=device-width, initial-scale=1"">"
    htStr = htStr & "<script src=""Scripts/jquery-3.3.1.js"" LANGUAGE=""javascript"">"
      htStr = htStr & "</script>"
  htStr = htStr & "<link rel='stylesheet' type='text/css' href='CSS/jquery.dataTables.min.css'>"
  htStr = htStr & "<link rel='stylesheet' type='text/css' href='CSS/bootstrap.minV4.css'>"
  response.write htStr
End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
