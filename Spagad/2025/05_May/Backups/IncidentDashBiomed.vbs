'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
response.write "    <link rel='preconnect' href='https://fonts.googleapis.com'>"
response.write "    <link rel='preconnect' href='https://fonts.gstatic.com' crossorigin>"
response.write "    <link href='https://fonts.googleapis.com/css2?family=Ubuntu:ital,wght@0,300;0,400;0,500;0,700;1,300;1,400;1,500;1,700&display=swap' rel='stylesheet'>"
response.write "        <style>"
response.write "        .incident{"
response.write "            border-collapse: collapse;"
response.write "            background-color: #F1F4FD;"
response.write "            font-family: 'Ubuntu', sans-serif;"
response.write "            font-weight: 300;"
response.write "        }"
response.write "        .main{"
response.write "            font-family: 'Ubuntu', sans-serif;"
response.write "            font-weight: 300;"
response.write "        }"
response.write "        .incident thead{"
response.write "            background-color: #1E1F22;"
response.write "            color: white;"
response.write "        }"
response.write "        .incident th, .incident td{"
response.write "            border: 1px solid gray;"
response.write "            padding: 5px 20px;"
response.write "            "
response.write "        }"
response.write "        .incident select,.main select{"
response.write "            background-color: #3B3B3B;"
response.write "            color: white;"
response.write "            padding: 5px 15px;"
response.write "            border-radius: 4px;"
response.write "            font-family: inherit;"
response.write "        }"
response.write "    </style>"

Dim filterValue
filterValue = Trim(Request.QueryString("filter-value"))

response.write "<div class='main' style='margin:20px 0;display:flex;place-items:center;place-content:left;gap:10px;'><label>filter</label><select onchange='filter(this)'><option value='' " & compare("", filterValue) & ">NONE</option><option value='C002' " & compare("C002", filterValue) & ">Solved</option>"
response.write "<option value='C001'' " & compare("C001", filterValue) & ">Pending</option>"
response.write "<option value='C003'' " & compare("C003", filterValue) & ">Unsolved</option></select></div>"
response.write "<script>"
response.write "function filter(el){"
response.write "window.location.href = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=incidentDash&PositionForTableName=WorkingDay&WorkingDayID=&filter-value=' + el.value;"
response.write "};"
response.write "    function onSelect(complid,type,el){"
response.write "        url = 'wpgXMLHTTP.asp?ProcedureName=assignIncident&type=' + type + '&complaintid=';"
response.write "        url += complid + '&value=' + el.value;"
response.write "        fetch(url).then(response=>response.json()).then(data=>console.log(data));"
response.write "    };"
response.write "</script>"
response.write "<div class='main'>"
response.write "<table class='incident'>"
response.write "        <thead>"
response.write "            <tr>"
response.write "                <th>DATE</th>"
response.write "                <th>DEPARTMENT</th>"
response.write "                <th>REQUESTED BY</th>"
response.write "                <th>INCIDENT TYPE</th>"
response.write "                <th>ASSIGN TO</th>"
response.write "                <th>ON BEHALF OF</th>"
response.write "                <th>STATUS</th>"
response.write "                <th>INCIDENT DETAIL</th>"
response.write "                <th>COMPLETION DATE</th>"
response.write "            </tr>"
response.write "        </thead>"
response.write "        <tbody>"


Dim rst
Set rst = CreateObject("ADODB.RecordSet")
sql = "SELECT systemcomplaintID,complainttypeid,complaintDetail,complaintDate,departmentid,systemuserid,ActionTaken,ComplaintDate "
sql = sql & " FROM SystemComplaint "
If Len(filterValue) > 0 Then
    If filterValue = "C003" Then
        sql = sql & " WHERE CAST(actiontaken AS NVARCHAR) <> 'C001' AND CAST(actiontaken AS NVARCHAR) <> 'C002' "
    Else
        sql = sql & " WHERE CAST(actiontaken AS NVARCHAR) = '" & filterValue & "' "
    End If
End If
sql = sql & " ORDER BY SystemComplaint.ComplaintDate DESC "

rst.open qryPro.FltQry(sql), conn, 3, 4
If rst.RecordCount > 0 Then
    rst.movefirst
    Do While Not rst.EOF
        colorgrid = ""
        actionTaken = rst.fields("ActionTaken")
        If actionTaken = "C002" Then
            colorgrid = "background-color:#ccffcc;"
        ElseIf actionTaken = "C001" Then
            colorgrid = "background-color:#ffffaa;"
        End If
        response.write "    <tr style='" & colorgrid & "'>"
        response.write "        <td>" & rst.fields("ComplaintDate") & "</td>"
        response.write "        <td>" & GetComboName("Department", rst.fields("departmentid")) & "</td>"
        response.write "        <td>" & GetComboName("Staff", GetComboNameFld("SystemUser", rst.fields("systemuserid"), "StaffID")) & "</td>"
        response.write "        <td>" & GetComboName("ComplaintType", rst.fields("complainttypeid")) & "</td>"
        response.write "        <td>"
        getITStaff rst.fields("systemcomplaintID")
        response.write "        </td>"
        response.write "        <td>"
        getBehlfStaff rst.fields("systemcomplaintID")
        response.write "        </td>"
        response.write "        <td>"
        getStatus rst.fields("systemcomplaintID")
        response.write "        </td>"
        response.write "        <td style='min-width:400px;'>" & rst.fields("complaintDetail") & "</td>"
        response.write "        <td>" & complaintDate(rst.fields("systemcomplaintID")) & "</td>"
        response.write "    </tr>"
        rst.MoveNext
        response.Flush
    Loop
End If
rst.Close
Set rst = Nothing
response.write "        </tbody></table>"
response.write "</div>"

Sub getITStaff(complaintid)
    Dim rst, rst2
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    sql = "SELECT staffid,staffname FROM Staff WHERE stafftypeid = 'SFT001' and staffid <> 'STF001' and StaffStatusID<>'S002'"
    rst.open qryPro.FltQry(sql), conn, 3, 4
    sql = "SELECT Performvar13name FROM PerformVar13 WHERE PerformVar13id = '" & complaintid & "'"
    rst2.open qryPro.FltQry(sql), conn, 3, 4
    response.write "<select onchange=""onSelect('" & complaintid & "','person',this)""><option selected hidden disabled>select user</option><option value=''>NONE</option>"
    If rst.RecordCount > 0 Then
        rst.movefirst
        If rst2.RecordCount > 0 Then
            If Not IsNull(rst2.fields("performvar13name")) Then
                arr = Split(rst2.fields("Performvar13name"), "||")
                If UBound(arr) >= 0 Then
                    selectedStaff = arr(0)
                End If
            End If
        Else
            selectedStaff = ""
        End If
        Do While Not rst.EOF
            response.write "<option value=""" & rst.fields("staffid") & """ " & compare(rst.fields("staffid"), selectedStaff) & ">" & rst.fields("staffname") & "</option>"
            rst.MoveNext
        Loop
    End If
    response.write "</select>"
    rst.Close
    rst2.Close
    Set rst2 = Nothing
    Set rst = Nothing
End Sub
Sub getBehlfStaff(complaintid)
    Dim rst, rst2
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    sql = "SELECT staffid,staffname FROM Staff WHERE staffid <> 'STF001' and StaffStatusID<>'S002'"
    rst.open qryPro.FltQry(sql), conn, 3, 4
    sql = "SELECT Performvar13name FROM PerformVar13 WHERE PerformVar13id = '" & complaintid & "'"
    rst2.open qryPro.FltQry(sql), conn, 3, 4
    response.write "<select onchange=""onSelect('" & complaintid & "','behalf',this)""><option selected hidden disabled>select user</option><option value=''>NONE</option>"
    If rst.RecordCount > 0 Then
        rst.movefirst
        If rst2.RecordCount > 0 Then
            If Not IsNull(rst2.fields("Performvar13name")) Then
                arr = Split(rst2.fields("Performvar13name"), "||")
                If UBound(arr) >= 0 Then
                    selectedStaff = arr(1)
                End If
            End If
        Else
            selectedStaff = ""
        End If
        Do While Not rst.EOF
            response.write "<option value=""" & rst.fields("staffid") & """ " & compare(rst.fields("staffid"), selectedStaff) & ">" & rst.fields("staffname") & "</option>"
            rst.MoveNext
        Loop
    End If
    response.write "</select>"
    rst.Close
    rst2.Close
    Set rst2 = Nothing
    Set rst = Nothing
End Sub

Sub getStatus(complaintid)
    Dim rst, rst2
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    sql = "SELECT complaintstatusid,complaintstatusname FROM ComplaintStatus "
    rst.open qryPro.FltQry(sql), conn, 3, 4
    sql = "SELECT Description FROM PerformVar13 WHERE PerformVar13id = '" & complaintid & "'"
    rst2.open qryPro.FltQry(sql), conn, 3, 4
    response.write "<select onchange=""onSelect('" & complaintid & "','status',this)""><option selected hidden disabled>select status</option><option value=''>NONE</option>"
    If rst.RecordCount > 0 Then
        rst.movefirst
        If rst2.RecordCount > 0 Then
            selectedStatus = rst2.fields("Description")
        Else
            selectedStatus = ""
        End If
        Do While Not rst.EOF
            response.write "<option value=""" & rst.fields("complaintstatusid") & """  " & compare(rst.fields("complaintstatusid"), selectedStatus) & ">" & rst.fields("complaintstatusname") & "</option>"
            rst.MoveNext
        Loop
    End If
    response.write "</select>"
    rst2.Close
    rst.Close
    Set rst = Nothing
    Set rst2 = Nothing
End Sub

Function compare(v1, v2)
    If v1 = v2 Then
        compare = "selected"
    Else
        compare = ""
    End If
End Function

Function complaintDate(complaintid)
   Dim rst, date1
   date1 = ""
   Set rst = CreateObject("ADODB.RecordSet")
   sql = "SELECT KeyPrefix FROM PerformVar13 WHERE performvar13id = '" & complaintid & "'"
   rst.open qryPro.FltQry(sql), conn, 3, 4
   If rst.RecordCount > 0 Then
    date1 = rst.fields("KeyPrefix")
   End If
   rst.Close
   Set rst = Nothing
   complaintDate = date1
End Function
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
