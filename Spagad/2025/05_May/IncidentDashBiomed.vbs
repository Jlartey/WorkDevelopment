'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

tableStyles
IncidentDashBiomed

Sub IncidentDashBiomed()
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT "
    sql = sql & "AssetIncidentID, AssetIncidentName, convert(varchar(20), IncidentDate1, 106) IncidentDate1, "
    sql = sql & "IncidentPriorityID, AssignedToID, ReportedBy, DirectoryID, IncidentStatusID, IncidentDetail, "
    sql = sql & "convert(VARCHAR(20), EndDate) EndDate, DirectoryTypeID "
    sql = sql & "FROM AssetIncident"


    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>IncidentID</th>"
            response.write "<th class='myth'>Incident Name</th>"
            response.write "<th class='myth'>Date</th>"
            response.write "<th class='myth'>Priority</th>"
            response.write "<th class='myth'>Assigned To</th>"
            response.write "<th class='myth'>Reported By</th>"
            response.write "<th class='myth'>On Behalf Of</th>"
            response.write "<th class='myth'>Status</th>"
            response.write "<th class='myth'>Incident Detail</th>"
            response.write "<th class='myth'>Completion Date</th>"
            response.write "<th class='myth'>Directory Type</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("AssetIncidentID") & "</td>"
                response.write "<td class='mytd'>" & .fields("AssetIncidentName") & "</td>"
                response.write "<td class='mytd'>" & .fields("IncidentDate1") & "</td>"
                response.write "<td class='mytd'>" & .fields("IncidentPriorityID") & "</td>"
                response.write "<td class='mytd'>" & .fields("AssignedToID") & "</td>"
                response.write "<td class='mytd'>" & .fields("ReportedBy") & "</td>"
                response.write "<td class='mytd'>" & .fields("DirectoryID") & "</td>"
                response.write "<td class='mytd'>" & .fields("IncidentStatusID") & "</td>"
                response.write "<td class='mytd'>" & .fields("IncidentDetail") & "</td>"
                response.write "<td class='mytd'>" & .fields("EndDate") & "</td>"
                response.write "<td class='mytd'>" & .fields("EndDDirectoryTypeIDate") & "</td>"
                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop

            response.write "</table>"
        Else
            response.write "<h1>No records found</h1>"
        End If
        
        .Close
    End With
    
    Set rst = Nothing
End Sub


Sub tableStyles()
    response.write "<style>"
        response.write ".mytable {"
        response.write "    width: 120vw;"
        response.write "    border-collapse: collapse;"
        response.write "    margin: 20px 0;"
        response.write "    font-size: 16px;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "}"
        response.write ".mytable, .myth, .mytd {"
        response.write "    border: 1px solid #dddddd;"
        response.write "}"
        response.write ".myth, .mytd {"
        response.write "    padding: 12px;"
        response.write "    text-align: left;"
        response.write "}"
        response.write ".myth {"
        response.write "    background-color: #f2f2f2;"
        response.write "    color: #333;"
        response.write "    font-weight: bold;"
        response.write "}"
        response.write ".mytr:nth-child(even) {"
        response.write "    background-color: #f9f9f9;"
        response.write "}"
        response.write ".mytr:hover {"
        response.write "    background-color: #f1f1f1;"
        response.write "}"
        response.write ".myth {"
        response.write "    text-transform: uppercase;"
        response.write "}"
        response.write "h1 {"
        response.write "    font-size: 18px;"
        response.write "    color: #555;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "    margin: 20px 0;"
        response.write "}"
response.write "</style>"

End Sub



'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
