
Dim selectedDay

selectedDay = Trim(Request.QueryString("WorkingDayID"))
'ShowControlPanel (BuildDict(Request.QueryString))
If UCase(Request.QueryString("appointOnly")) = UCase("true") Then
    GetAppointmentListing (BuildDict(Request.QueryString))
Else
    GetAppointmentVisitListing (BuildDict(Request.QueryString))
End If

Function getcontrolpanel(filterDict)
    Dim str, html
    html = html & "<div class='main-page' style='/*margin-top:20px;*//*width:98vw*/'>"
        html = html & " <div class='top-bar'>"
        html = html & GetAppointmentCreatePrintLink(filterDict)
        html = html & GetAppointmentDateSelector(filterDict)
        html = html & GetAppointmentDoctorSelector(filterDict)
        html = html & GetViewTypeSelector(filterDict)
        html = html & GetVisitCreatorLink(filterDict)
        html = html & GetAppointmentCreatorLink(filterDict)
        html = html & " </div>"
        html = html & " <div class='main-content'>"
        'html = html & GetAppointmentVisitListing(filterDict)
        html = html & " </div>"
    html = html & "</div>"

    getcontrolpanel = html
End Function
Function GetAppointmentCreatePrintLink(filterDict)
    Dim ot
    ot = ot & "<div class='no-print' style='display:inline-block;margin-left:20px;'><span onclick='openPrintPage()' style='cursor:pointer;color:#2196F3;'>Print</span></div> "
    GetAppointmentCreatePrintLink = ot
End Function
Function GetVisitCreatorLink(filterDict)
    Dim html, href
    
    href = "wpgPrtPrintlayoutAll.asp?PositionForTableName=WorkingDay&PrintLayoutName=SelectPatientQuick"
    GetVisitCreatorLink = "<div class='no-print' style='float:right;margin-right:20px;'>" & GetLink(href, "Search/Create Patient/Visit (Quick!)", "#1E5B5B") & "</div>"
End Function
Function GetViewTypeSelector(filterDict)
    Dim ot
    ot = ot & "<div class='no-print' style='display:inline-block;margin-left:20px;'><input type='checkbox' id='appointOnly' onchange='toggleAppointOnly(this)' " & IIF(filterDict("appointOnly") = "true", "checked", "") & "/>Appointment Only</div> "
    GetViewTypeSelector = ot
End Function
Function BuildDict(qryString)
    Dim ot, Key, cWDay

    cWDay = FormatWorkingDay(Now())
    Set ot = CreateObject("Scripting.Dictionary")
    ot.CompareMode = 1
    For Each Key In qryString
        ot.Add Key, Trim(qryString(Key))
    Next
    ot("AppointStartDayID") = IIF(ot("AppointStartDayID") = "", cWDay, ot("AppointStartDayID"))
    
    If UCase(ot("AppointEndDayID")) < UCase(ot("AppointStartDayID")) Then
        ot("AppointEndDayID") = ot("AppointStartDayID")
    End If
    
    Set BuildDict = ot
End Function
Function GetPageCSS()
    Dim html

    html = html & "     .appoint-listing{width:100%;margin-top:20px;} "
    html = html & "     .force-gray, .force-gray *{color:gray!important;} "
    html = html & "     .appoint-listing th{background-color:#e3e4e5;height:40px;} "
    html = html & "     .appoint-listing, .appoint-listing *{ "
    html = html & "         /*text-transform:uppercase;font-size:13px;*/"
    html = html & "     }"
    html = html & "     .appoint-listing .blink{ "
    html = html & "         animation: blink-animation 1s steps(5, start) infinite;-webkit-animation: blink-animation 1s steps(5, start) infinite;"
    html = html & "     }"
    html = html & "     @keyframes blink-animation {"
    html = html & "       to {"
    html = html & "         visibility: hidden;"
    html = html & "       }"
    html = html & "     }"
    html = html & "     @-webkit-keyframes blink-animation {"
    html = html & "       to {"
    html = html & "         visibility: hidden;"
    html = html & "       }"
    html = html & "     }"
    html = html & "     option{padding:3px;}"

    GetPageCSS = html
End Function
Function GetPageJS(filterDict)
    Dim html, href
    
    href = "wpgPrtPrintlayoutAll.asp?PositionForTableName=WorkingDay&PrintLayoutName=DynamicTableLoader&LoadInterval=60000&ProcedureName=MonitorAppointmentJSON"
    href = href & "&Parameter=Print||Yes"
    
    html = html & " function  processDateChange(select, paramName){ "
    html = html & "     select.disabled = true;"
    html = html & "     let url_search = new URL(window.location.href);"
    html = html & "     " & filterDict("DynamicTableName") & ".setParam(paramName, select.options[select.selectedIndex].value);"
    html = html & " }"
    html = html & " function  processSpecialistChange(select){ "
    html = html & "     select.disabled = true;"
    html = html & "     let url_search = new URL(window.location.href);"
    html = html & "     " & filterDict("DynamicTableName") & ".setParam('SpecialistID', select.options[select.selectedIndex].value);"
    html = html & " }"
    html = html & ""
    html = html & " function toggleAppointOnly(checkbox){"
    html = html & "     checkbox.disabled = true;"
    html = html & "     let url_search = new URL(window.location.href);"
    html = html & "     " & filterDict("DynamicTableName") & ".setParam('appointOnly', checkbox.checked);"
    html = html & " "
    html = html & " }"
    
    html = html & " function openPopup(anc){"
    html = html & "     let win=window.open(anc.dataset.href, '_blank', 'resizeable=yes,scrollbars=yes,width=820,height=560,status=yes');"
    html = html & "     "
    html = html & "     let intvl = setInterval(function(){"
    html = html & "         if(win.closed !== false){"
    html = html & "             clearInterval(intvl);"
    'html = html & "             window.location.reload();"
    html = html & "          }"
    html = html & "     }, 200);"
    html = html & "}"
    html = html & " function openPrintPage(){ "
''    html = html & "     let url = " & filterDict("DynamicTableName") & ".url;"
''        html = html & "     let url = '" & href & "'; "
''        html = html & "     window.open(url, '_blank');"
    html = html & " window.print();"
    html = html & " } "

    GetPageJS = html
End Function
Function GetAppointmentDateSelector(filterDict)
    Dim html, sql, rst, dt, cWDay, startSelect, endSelect, stDy, dayName, isSelected, startDay, endDay, enDy

    dt = Now()
    stDy = UCase(filterDict("AppointStartDayID"))
    enDy = UCase(filterDict("AppointEndDayID"))
    cWDay = FormatWorkingDay(dt)
    startDay = FormatWorkingDay(DateAdd("d", dt, -60))
    endDay = FormatWorkingDay(DateAdd("d", dt, 30))

    Set rst = CreateObject("ADODB.RecordSet")

    sql = "select distinct WorkingDay.WorkingDayID, WorkingDay.WorkingDayName "
    sql = sql & " from WorkingDay "
    sql = sql & " left join Appointment on Appointment.AppointDayID=WorkingDay.WorkingDayID"
    sql = sql & " left join Visitation on Visitation.VisitationID=WorkingDay.WorkingDayID"
    sql = sql & " where (WorkingDay.WorkingDayID>='" & startDay & "' and WorkingDay.WorkingDayID<='" & endDay & "'"
    sql = sql & " and (Appointment.WorkingDayID is not null or Visitation.WorkingDayID is not null)) or WorkingDay.WorkingDayID='" & cWDay & "' "
    sql = sql & " order by WorkingDay.WorkingDayID asc, WorkingDay.WorkingDayName asc "
    
    startSelect = "<select onchange=""processDateChange(this, 'AppointStartDayID')"" style=""padding:5px;"">"
    startSelect = startSelect & "<option></option>"
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        
        Do While Not rst.EOF
            isSelected = IIF(stDy = rst.fields("WorkingDayID"), "selected", "")
            dayName = IIF(rst.fields("WorkingDayID") = cWDay, "Today -> " & rst.fields("WorkingDayName") & "", rst.fields("WorkingDayName"))
            startSelect = startSelect & "<option value=""" & rst.fields("WorkingDayID") & """ " & isSelected & ">" & dayName & "</option>"
            rst.MoveNext
        Loop
        rst.Close
    End If
    startSelect = startSelect & "</select>"
    Set rst = Nothing
    
    Set rst = CreateObject("ADODB.RecordSet")

    sql = "select distinct WorkingDay.WorkingDayID, WorkingDay.WorkingDayName "
    sql = sql & " from WorkingDay "
    sql = sql & " left join Appointment on Appointment.AppointDayID=WorkingDay.WorkingDayID"
    sql = sql & " left join Visitation on Visitation.VisitationID=WorkingDay.WorkingDayID"
    sql = sql & " where (WorkingDay.WorkingDayID>='" & startDay & "' and WorkingDay.WorkingDayID<='" & endDay & "'"
    sql = sql & "   and (Appointment.WorkingDayID is not null or Visitation.WorkingDayID is not null) "
    If stDy <> "" Then
        sql = sql & " and Appointment.AppointDayID>='" & stDy & "'"
    End If
    sql = sql & ")"
    sql = sql & "  or WorkingDay.WorkingDayID='" & cWDay & "' "
    sql = sql & " order by WorkingDay.WorkingDayID asc, WorkingDay.WorkingDayName asc "
    endSelect = "<select onchange=""processDateChange(this, 'AppointEndDayID')"" style=""padding:5px;"">"
    endSelect = endSelect & "<option></option>"
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            isSelected = IIF(enDy = rst.fields("WorkingDayID"), "selected", "")
            dayName = IIF(rst.fields("WorkingDayID") = cWDay, "Today -> " & rst.fields("WorkingDayName") & "", rst.fields("WorkingDayName"))
            endSelect = endSelect & "<option value=""" & rst.fields("WorkingDayID") & """ " & isSelected & ">" & dayName & "</option>"
            rst.MoveNext
        Loop
        rst.Close
    End If
    endSelect = endSelect & "</select>"
    Set rst = Nothing

    html = "<div style='display:inline-block;margin-left:20px;'>From Day:" & startSelect & " To Day:" & endSelect & "</div>"
    GetAppointmentDateSelector = html
End Function
Function GetAppointmentDoctorSelector(filterDict)
    Dim html, sql, rst, dt, cWDay, htmlSelect, dy, dayName, isSelected

    dt = Now()
    dy = UCase(Trim(filterDict("AppointStartDayID")))
    cWDay = FormatWorkingDay(dt)
    Set rst = CreateObject("ADODB.RecordSet")

    sql = "select case when Staff.StaffID='STF001' then JobSchedule.JobScheduleName else Staff.StaffName end as StaffName"
    sql = sql & " , Appointment.SpecialistID from Appointment "
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=Appointment.SpecialistID "
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID "
    sql = sql & " left join JobSchedule on JobSchedule.JobScheduleID=SystemUser.JobScheduleID"
    sql = sql & " where Appointment.AppointDayID between '" & filterDict("AppointStartDayID") & "'  and  '" & filterDict("AppointEndDayID") & "' "
    sql = sql & " union "
    sql = sql & "select case when Staff.StaffID='STF001' then JobSchedule.JobScheduleName else Staff.StaffName end as StaffName"
    sql = sql & ", Visitation.SpecialistID from Visitation "
    sql = sql & " left join SystemUser on SystemUser.SystemUserID=Visitation.SpecialistID "
    sql = sql & " left join Staff on Staff.StaffID=SystemUser.StaffID "
    sql = sql & " left join JobSchedule on JobSchedule.JobScheduleID=SystemUser.JobScheduleID"
    sql = sql & " where Visitation.WorkingDayID between '" & filterDict("AppointStartDayID") & "'  and  '" & filterDict("AppointEndDayID") & "' "
    

    htmlSelect = "<select style='padding:5px;' onchange='processSpecialistChange(this)'>"
    htmlSelect = htmlSelect & "<option></option>"
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        
        Do While Not rst.EOF
            isSelected = IIF(UCase(filterDict("SpecialistID")) = UCase(rst.fields("SpecialistID")), "selected", "")
            htmlSelect = htmlSelect & "<option value=""" & rst.fields("SpecialistID") & """ " & isSelected & ">" & rst.fields("StaffName") & "</option>"
            rst.MoveNext
        Loop
        rst.Close
    End If
    htmlSelect = htmlSelect & "</select>"
    Set rst = Nothing

    html = "<div style='display:inline-block;margin-left:20px;'>Specialist:" & htmlSelect & "</div>"
    GetAppointmentDoctorSelector = html
End Function
Function GetAppointmentCreatorLink(filterDict)
    Dim html, href
    
    href = "wpgAppointment.asp?PageMode=AddNew&PullUpData=PatientID||P4"
    GetAppointmentCreatorLink = "<div class='no-print' style='float:right;margin-right:20px;'>" & GetLink(href, "Book Appointment", "green") & "</div>"

End Function
Function GetLink(href, linkText, lnkColor)
    Dim html, defColor
    defColor = IIF(Trim(lnkColor) = "", "#2196F3", lnkColor)

    html = "<div>"
    html = html & "<span data-href=""" & href & """ style='display:inline-block;color:" & defColor & ";/*font-weight:bold;*/text-transform:none;cursor:pointer;' "
    html = html & "onclick=""openPopup(this)"""
    html = html & ">" & linkText & "</span>"
    html = html & "</div>"
    GetLink = html
End Function
Function GetAppointmentListing(filterDict)
    Dim html, rst, sql, href, jsonDict, rows, rowDict, header, rowData, patDetail, apntDetail
 
    Set jsonDict = CreateObject("Scripting.Dictionary")
    Set rows = CreateObject("System.Collections.ArrayList")
    Set header = CreateObject("System.Collections.ArrayList")

    sql = " select Patient.PatientID, Gender.GenderName, Patient.BirthDate "
    sql = sql & " , (case when Patient.PatientID='P4' then Appointment.AppointmentName else Patient.PatientName end) as PatientName"
    sql = sql & " , Appointment.AppointmentID, 'APPOINT_ONLY' as Type "
    sql = sql & " , Appointment.AppointStartTimeID, Appointment.AppointEndTimeID"
    sql = sql & " , AppointmentCat.AppointmentCatName + ' -> ' + AppointmentCatType.AppointmentCatTypeName as AppointDetail"
    sql = sql & " , Appointment.SpecialistID as AppointDoctorID"
    sql = sql & " , Appointment.AppointmentStatusID"
    sql = sql & " , (case when Patient.PatientID='P4' then Appointment.KeyPrefix else Patient.ResidencePhone end) as [PhoneNo]"
    sql = sql & " from Appointment "
    sql = sql & " left join Patient on Patient.PatientID=Appointment.PatientID"
    sql = sql & " left join AppointmentCat on AppointmentCat.AppointmentCatID=Appointment.AppointmentCatID"
    sql = sql & " left join AppointmentCatType on AppointmentCatType.AppointmentCatTypeID=Appointment.AppointmentCatTypeID"
    sql = sql & " left join Gender on Gender.GenderID=Patient.GenderID "
    sql = sql & " where 1=1 "
    If filterDict("AppointEndDayID") = "" Then
        sql = sql & "   and Appointment.AppointDayID='" & filterDict("AppointStartDayID") & "' "
    Else
        sql = sql & "   and ( (Appointment.AppointDayID between '" & filterDict("AppointStartDayID") & "' and '" & filterDict("AppointEndDayID") & "') )"
    End If
    If filterDict("SpecialistID") <> "" Then
        sql = sql & " and (Appointment.SpecialistID='" & filterDict("SpecialistID") & "' or )"
    End If
    sql = sql & "  order by Appointment.AppointDayID asc, Appointment.AppointStartTimeID asc" ', Visitation.VisitDate asc"

    'header.Add "No."
    header.Add "Appoint No."
    header.Add "Patient Name"
    header.Add "Details"
    header.Add "Control"

    jsonDict.Add "javascript-text", GetPageJS(filterDict)
    jsonDict.Add "css-text", GetPageCSS()
    jsonDict.Add "class", "appoint-listing"
    jsonDict.Add "style", "width:99.9%;margin:auto;"
    jsonDict.Add "header", header
    jsonDict.Add "title", ("Monitor Appointments -> " & GetComboName("AppointDay", filterDict("AppointStartDayID")))
    jsonDict.Add "show-row-numbers", True
    jsonDict.Add "control-panel", getcontrolpanel(filterDict)

    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            Set rowData = CreateObject("System.Collections.ArrayList")
            Set rowDict = CreateObject("Scripting.Dictionary")

            href = "wpgAppointment.asp?PageMode=ProcessSelect&AppointmentID=" & rst.fields("AppointmentID")
            rowData.Add GetLink(href, rst.fields("AppointmentID"), "") 'appoint id

            patDetail = "<div style='text-transform:uppercase;'>" & GetPatientName(rst.fields("PatientName")) & "</div>"
            If (Len(rst.fields("PatientID")) > 3) Then
                patDetail = patDetail & "<div style='font-size:10pt;padding-top:5px;'>Folder No: " & rst.fields("PatientID")
                patDetail = patDetail & ", " & rst.fields("GenderName")
                If Not IsNull(rst.fields("BirthDate")) Then
                    age = DateDiff("yyyy", rst.fields("BirthDate"), Now())
                    patDetail = patDetail & ", " & age & "Y"
                End If
                patDetail = patDetail & "</div>"
            Else
                patDetail = patDetail & "<div style='font-size:10pt;padding-top:5px;'>"
                pDt = GetComboNameFld("Appointment", rst.fields("AppointmentID"), "AppointDate2")
                If Not IsNull(pDt) And IsDate(pDt) Then
                    age = DateDiff("yyyy", pDt, Now())
                    patDetail = patDetail & " " & age & "Y"
                End If
                patDetail = patDetail & "</div>"
            End If
            patDetail = patDetail & "<div style='font-size:10pt;padding-top:5px;'>Contact: " & rst.fields("PhoneNo") & "</div>"
            rowData.Add patDetail
            
            apntDetail = "<div style='color:#a94442'>" & GetComboName("AppointStartTime", rst.fields("AppointStartTimeID")) & " - " & GetComboName("AppointEndTime", rst.fields("AppointEndTimeID")) & "</div>"
            apntDetail = apntDetail & rst.fields("AppointDetail") & "<br/>" & GetDoctorName(rst.fields("AppointDoctorID"))
            rowData.Add apntDetail
        
            rowData.Add GetControlLinks(rst.fields)
        
            rowDict.Add "row-data", rowData
            If rst.fields("AppointmentStatusID") = "A004" Then
                rowDict.Add "is-canceled", True
            End If
            
            rows.Add rowDict

            rst.MoveNext
        Loop
        jsonDict.Add "rows", rows
        rst.Close
    End If
    If IsObject(response) Then
        response.Clear
        response.ContentType = "application/json"
        response.write JSONStringify(jsonDict)
    End If
End Function
Function GetAppointmentVisitListing(filterDict)
    Dim html, rst, sql, href, jsonDict, rows, rowDict, header, rowData, patDetail, apntDetail
 
    Set jsonDict = CreateObject("Scripting.Dictionary")
    Set rows = CreateObject("System.Collections.ArrayList")
    Set header = CreateObject("System.Collections.ArrayList")

    'has either appointments or visits
    sql = " select top 400 Visitation.VisitationID, Visitation.VisitStatusID, Patient.PatientID, Gender.GenderName, Patient.BirthDate "
    sql = sql & " , (case when Patient.PatientID='P4' then Appointment.AppointmentName else Patient.PatientName end) as PatientName"
    sql = sql & " , Appointment.AppointmentID, 'APPOINT_VISIT' as Type "
    sql = sql & " , Appointment.AppointStartTimeID, Appointment.AppointEndTimeID"
    sql = sql & " , VstSpecGroup.SpecialistGroupName + ' -> ' + VstSpecType.SpecialistTypeName as VisitDetail"
    sql = sql & " , AppointmentCat.AppointmentCatName + ' -> ' + AppointmentCatType.AppointmentCatTypeName as AppointDetail"
    sql = sql & " , Appointment.SpecialistID as AppointDoctorID, Visitation.SpecialistID as VisitDoctorID"
    sql = sql & " , (case when Appointment.AppointmentID is null then 'VISIT_ONLY' when Visitation.VisitationID is null then 'APPOINT_ONLY' else 'APPOINT_VISIT' end) as Type "
    sql = sql & " , Appointment.AppointmentStatusID"
    sql = sql & " , (case when Patient.PatientID='P4' then Appointment.KeyPrefix else Patient.ResidencePhone end) as [PhoneNo]"
    sql = sql & " from Appointment "
    sql = sql & " full outer join Visitation on Visitation.PatientID=Appointment.PatientID and Appointment.AppointmentID=Visitation.VisitInfo4"
    sql = sql & " left join Patient on Patient.PatientID=Appointment.PatientID or Patient.PatientID=Visitation.PatientID "
    sql = sql & " left join SpecialistGroup as VstSpecGroup on VstSpecGroup.SpecialistGroupID=Visitation.SpecialistGroupID"
    sql = sql & " left join SpecialistType as VstSpecType on VstSpecType.SpecialistTypeID=Visitation.SpecialistTypeID "
    sql = sql & " left join AppointmentCat on AppointmentCat.AppointmentCatID=Appointment.AppointmentCatID"
    sql = sql & " left join AppointmentCatType on AppointmentCatType.AppointmentCatTypeID=Appointment.AppointmentCatTypeID"
    sql = sql & " left join Gender on Gender.GenderID=Patient.GenderID "
    sql = sql & " where 1=1 "
    If filterDict("AppointEndDayID") = "" Then
        sql = sql & "   and (Visitation.WorkingDayID='" & filterDict("AppointStartDayID") & "' or Appointment.AppointDayID='" & filterDict("AppointStartDayID") & "')"
    Else
        sql = sql & "   and ( (Visitation.WorkingDayID between '" & filterDict("AppointStartDayID") & "' and '" & filterDict("AppointEndDayID") & "') or (Appointment.AppointDayID between '" & filterDict("AppointStartDayID") & "' and '" & filterDict("AppointEndDayID") & "') )"
    End If
    If filterDict("SpecialistID") <> "" Then
        sql = sql & " and (Visitation.SpecialistID='" & filterDict("SpecialistID") & "' or Appointment.SpecialistID='" & filterDict("SpecialistID") & "')"
    End If
    sql = sql & "  order by Visitation.VisitDate asc, Appointment.AppointDayID asc, Appointment.AppointStartTimeID asc"

    'header.Add "No."
    header.Add "Appoint No."
    header.Add "Patient Name"
    header.Add "Visit no."
    header.Add "Details"
    
'    header.Add "Pharm"
    header.Add "Lab"
'    header.Add "Rad"
'    header.Add "Adm"
    'header.Add "Proc"
    header.Add "Control"

    jsonDict.Add "javascript-text", GetPageJS(filterDict)
    jsonDict.Add "css-text", GetPageCSS()
    jsonDict.Add "class", "appoint-listing"
    jsonDict.Add "style", "width:99.9%;margin:auto;"
    jsonDict.Add "header", header
    jsonDict.Add "title", ("Monitor Appointments -> " & GetComboName("AppointDay", filterDict("AppointStartDayID")))
    jsonDict.Add "show-row-numbers", True
    jsonDict.Add "control-panel", getcontrolpanel(filterDict)

    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            Set rowData = CreateObject("System.Collections.ArrayList")
            Set rowDict = CreateObject("Scripting.Dictionary")

            
            'rowDict.Add "class", IIF(IsNull(rst.fields("VisitationID")), "force-gray", "")
            'rowData.Add rst.AbsolutePosition 'no

            href = "wpgAppointment.asp?PageMode=ProcessSelect&AppointmentID=" & rst.fields("AppointmentID")
            rowData.Add GetLink(href, rst.fields("AppointmentID"), "") 'appoint id

            patDetail = "<div style='text-transform:uppercase;'>" & GetPatientName(rst.fields("PatientName")) & "</div>"
            If (Len(rst.fields("PatientID")) > 3) Then
                patDetail = patDetail & "<div style='font-size:10pt;padding-top:5px;'>Folder No: " & rst.fields("PatientID")
                patDetail = patDetail & ", " & rst.fields("GenderName")
                If Not IsNull(rst.fields("BirthDate")) Then
                    age = DateDiff("yyyy", rst.fields("BirthDate"), Now())
                    patDetail = patDetail & ", " & age & "Y"
                End If
                patDetail = patDetail & "</div>"
            Else
                patDetail = patDetail & "<div style='font-size:10pt;padding-top:5px;'>"
                pDt = GetComboNameFld("Appointment", rst.fields("AppointmentID"), "AppointDate2")
                If Not IsNull(pDt) And IsDate(pDt) Then
                    age = DateDiff("yyyy", pDt, Now())
                    patDetail = patDetail & " " & age & "Y"
                End If
                patDetail = patDetail & "</div>"
            End If
            patDetail = patDetail & "<div style='font-size:10pt;padding-top:5px;'>Contact: " & rst.fields("PhoneNo") & "</div>"
            If (Len(rst.fields("PatientID")) > 3) Then
                If UCase(rst.fields("Type")) = "APPOINT_ONLY" Then
                    patDetail = patDetail & ""
                End If
                patDetail = patDetail & "<div style='font-size:11px;padding-top:5px;'>" & GetComboName("Sponsor", GetComboNameFld("Visitation", rst.fields("VisitationID"), "SponsorID")) & "</div>"
            End If
            rowData.Add patDetail
            
            'href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation&WorkFlowNav=POP&VisitationID=" & rst.fields("VisitationID")
            'href = "wpgVisitation.asp?PageMode=ProcessSelect&VisitationID=" & rst.fields("VisitationID")
      href = "wpgVisitation.asp?PageMode=ProcessSelect&VisitationID=" & rst.fields("VisitationID")
  If HasPrintOutAccess(jSchd, "PatientMedicalRecord") Then
      href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientMedicalRecord&PositionForTableName=WorkingDay&PositionForCtxTableName=Visitation&VisitationID=" & rst.fields("VisitationID")
  ElseIf Len(HasModuleMgrAccess(jSchd, "Visitation")) > 0 Then
      href = "wpgSelectModuleManager.asp?PositionForTableName=Visitation&PositionForCtxTableName=Visitation&VisitationID=" & rst.fields("VisitationID")
  ElseIf HasPrintOutAccess(jSchd, "VisitationRCP") Then
      href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationRCP&PositionForTableName=Visitation&AddOtherSection=All&PositionForCtxTableName=Visitation&VisitationID=" & rst.fields("VisitationID")
  End If
            rowData.Add IIF(rst.fields("VisitationID") = "", "-", GetLink(href, rst.fields("VisitationID"), "")) 'visit no
            
            If UCase(rst.fields("Type")) = "APPOINT_ONLY" Then 'detail
                apntDetail = "<div style='color:#a94442'>" & GetComboName("AppointStartTime", rst.fields("AppointStartTimeID")) & " - " & GetComboName("AppointEndTime", rst.fields("AppointEndTimeID")) & "</div>"
                apntDetail = apntDetail & rst.fields("AppointDetail") & "<br/>" & GetDoctorName(rst.fields("AppointDoctorID"))
                rowData.Add apntDetail
            Else
                apntDetail = "<div style='color:#a94442'>" & GetComboName("AppointStartTime", rst.fields("AppointStartTimeID")) & " - " & GetComboName("AppointEndTime", rst.fields("AppointEndTimeID")) & "</div>"
                apntDetail = apntDetail & rst.fields("VisitDetail") & "<br/>" & GetDoctorName(rst.fields("VisitDoctorID"))
                rowData.Add apntDetail
            End If
            
'            rowData.Add TryGetPharmCtrlLinks(rst.fields("VisitationID")) 'pharm
            rowData.Add TryGetLabCtrlLinks(rst.fields("VisitationID")) 'lab
'            rowData.Add TryGetRadCtrlLinks(rst.fields("VisitationID")) 'rad
'            rowData.Add TryGetAdmCtrlLinks(rst.fields("VisitationID")) 'adm
            'rowData.Add tryGetProcCtrlLinks(rst.fields("VisitationID")) 'proc
            rowData.Add GetControlLinks(rst.fields)
        
            rowDict.Add "row-data", rowData
            If rst.fields("AppointmentStatusID") = "A004" Then
                rowDict.Add "is-canceled", True
            End If
            
            rows.Add rowDict

            rst.MoveNext
            
        Loop
        jsonDict.Add "rows", rows
        rst.Close
    End If
    If IsObject(response) Then
        response.Clear
        response.ContentType = "application/json"
        response.write JSONStringify(jsonDict)
    End If
End Function
Function HasPrintOutAccess(jb, prt)
  Dim rstTblSql, sql, ot
  ot = False
  Set rstTblSql = CreateObject("ADODB.Recordset")
  With rstTblSql
    sql = "select JobScheduleID from printoutalloc "
    sql = sql & " where printlayoutid='" & prt & "' and jobscheduleid='" & jb & "'"
    .open qryPro.FltQry(sql), conn, 3, 4
    If .recordCount > 0 Then
      .MoveFirst
      ot = True
    End If
    .Close
  End With
  HasPrintOutAccess = ot
  Set rstTblSql = Nothing
End Function
Function HasModuleMgrAccess(jb, tb)
  Dim rstTblSql, sql, ot
  ot = ""
  Set rstTblSql = CreateObject("ADODB.Recordset")
  With rstTblSql
    sql = "select ModuleManagerID from ModuleManageralloc "
    sql = sql & " where tableid='" & tb & "' and jobscheduleid='" & jb & "' order by ModuleManagerID"
    .open qryPro.FltQry(sql), conn, 3, 4
    If .recordCount > 0 Then
      .MoveFirst
      ot = .fields("ModuleManagerID")
    End If
    .Close
  End With
  HasModuleMgrAccess = ot
  Set rstTblSql = Nothing
End Function
Function GetPatientName(patName)
    Dim ot, kyVal, arrVal
    
    ot = ""
    For Each kyVal In Split(patName, "||")
        arrVal = Split(kyVal, "=")
        If ot <> "" Then ot = ot & " "
        If UBound(arrVal) > 0 Then
            ot = ot & arrVal(1)
        End If
    Next
    
    If ot = "" Then
        ot = Replace(patName, "||", " ")
    End If
    
    GetPatientName = ot
End Function
Function TryGetPharmCtrlLinks(vst)
    Dim sql, rst, ot, presCnt, dispCnt
    ot = ""
    If vst <> "" Then
        presCnt = 0: dispCnt = 0

        sql = "select count(Prescription.PrescriptionID) as [count] from Prescription where Prescription.VisitationID='" & vst & "' "
        Set rst = CreateObject("ADODB.RecordSet")
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.recordCount > 0 Then
            presCnt = rst.fields("Count")
            rst.Close
        End If
        Set rst = Nothing

        sql = "select count(DrugSale.DrugSaleID) as [count] from DrugSale where DrugSale.VisitationID='" & vst & "' "
        Set rst = CreateObject("ADODB.RecordSet")
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.recordCount > 0 Then
            dispCnt = rst.fields("Count")
            rst.Close
        End If
        Set rst = Nothing
        
        If presCnt > 0 Then
            ot = "<span style='display:block;text-transform:none;'>Presc:<b>" & presCnt & "</b></span>"
        End If
        If dispCnt > 0 Then
            ot = ot & "<span style='display:block;text-transform:none;'>Dispense:<b>" & dispCnt & "</b></span>"
        End If
        
    End If
    'If ot = "" Then
    '   ot = "-"
    'End If
    TryGetPharmCtrlLinks = ot
End Function
Function TryGetLabCtrlLinks(vst)
    Dim sql, rst, ot, labByDoc, href

    If vst <> "" Then
        labByDoc = 0
        sql = "select count(LabByDoctor.LabTestID) as [count] from LabByDoctor where LabByDoctor.VisitationID='" & vst & "' and LabByDoctor.TestCategoryID='B13' and LabByDoctor.LabByDoctorStatusID='L001'"
        Set rst = CreateObject("ADODB.RecordSet")
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.recordCount > 0 Then
            labByDoc = rst.fields("Count")
            If labByDoc > 0 Then
                ot = "<span style='display:block;text-transform:none;color:" & IIF(labByDoc > 0, "#e91e63", "black") & "' class='" & IIF(labByDoc > 0, "blink", "") & "'>Doc Req:<b>" & labByDoc & "</b></span>"
            End If
            rst.Close
        End If
        Set rst = Nothing

        sql = "select RequestStatusName, count(LabTestID) as [count]"
        sql = sql & " from ("
        sql = sql & "       select Investigation.LabTestID, Investigation.RequestStatusID from Investigation where TestCategoryID='B13' and VisitationID='" & vst & "'"
        sql = sql & "       union all select Investigation2.LabTestID, Investigation2.RequestStatusID from Investigation2 where TestCategoryID='B13' and VisitationID='" & vst & "'"
        sql = sql & " ) as [Req]"
        sql = sql & " left join RequestStatus on RequestStatus.RequestStatusID=[Req].RequestStatusID "
        sql = sql & " group by RequestStatusName"

        Set rst = CreateObject("ADODB.RecordSet")
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.recordCount > 0 Then
            rst.MoveFirst
            href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=ViewPatientLabRequest&PositionForTableName=WorkingDay&TestCategoryID=B19&VisitationID=" & vst
            Do While Not rst.EOF
                ot = ot & GetLink(href, rst.fields("RequestStatusName") & ":<b>" & rst.fields("Count"), "") & "</b>"
                'ot = ot & "<span style='display:block;text-transform:none;white-space:nowrap;'>" & rst.fields("RequestStatusName") & ":<b>" & rst.fields("Count") & "</b></span>"
                rst.MoveNext
            Loop
            rst.Close
        End If
        Set rst = Nothing

        href = "wpgLabRequest.asp?PageMode=AddNew&PullupData=VisitationID||" & vst
        ot = ot & GetLink(href, IIF(labByDoc > 0, "Process Request", "Make Request"), "#a94442")
    End If

    TryGetLabCtrlLinks = ot
End Function
Function TryGetRadCtrlLinks(vst)
    Dim sql, rst, ot, labByDoc, href, bgColor
        
    If vst <> "" Then
        labByDoc = 0
        sql = "select count(LabByDoctor.LabTestID) as [count] from LabByDoctor where LabByDoctor.VisitationID='" & vst & "' and LabByDoctor.TestCategoryID='B19' and LabByDoctor.LabByDoctorStatusID='L001'"
        Set rst = CreateObject("ADODB.RecordSet")
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.recordCount > 0 Then
            labByDoc = rst.fields("Count")
            If labByDoc > 0 Then
                ot = "<span style='display:block;text-transform:none;color:" & IIF(labByDoc > 0, "#e91e63", "black") & "' class='" & IIF(labByDoc > 0, "blink", "") & "'>Doc Req:<b>" & labByDoc & "</b></span>"
            End If
            rst.Close
        End If
        Set rst = Nothing

        sql = "select RequestStatusName, req.RequestStatusID, count(LabTestID) as [count]"
        sql = sql & " from ("
        sql = sql & "       select Investigation.LabTestID, Investigation.RequestStatusID from Investigation where TestCategoryID='B19' and VisitationID='" & vst & "'"
        sql = sql & "       union all select Investigation2.LabTestID, Investigation2.RequestStatusID from Investigation2 where TestCategoryID='B19' and VisitationID='" & vst & "'"
        sql = sql & " ) as [Req]"
        sql = sql & " left join RequestStatus on RequestStatus.RequestStatusID=[Req].RequestStatusID "
        sql = sql & " group by RequestStatusName, req.RequestStatusID"

        Set rst = CreateObject("ADODB.RecordSet")
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.recordCount > 0 Then
            rst.MoveFirst
            href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=ViewPatientLabRequest&PositionForTableName=WorkingDay&TestCategoryID=B19&VisitationID=" & vst
            Do While Not rst.EOF
                ot = ot & GetLink(href, rst.fields("RequestStatusName") & ":<b>" & rst.fields("Count"), "") & "</b>"
                'ot = ot & "<span style='display:block;text-transform:none;white-space:nowrap;'>" & rst.fields("RequestStatusName") & ":<b>" & rst.fields("Count") & "</b></span>"
                rst.MoveNext
            Loop
            rst.Close
        End If
        Set rst = Nothing

        href = "wpgLabRequest.asp?PageMode=AddNew&PullupData=VisitationID||" & vst
        ot = ot & GetLink(href, IIF(labByDoc > 0, "Process Request", "Make Request"), "#a94442")


        'ot = "<div style='color:"& bgColor &"'>"& ot &"</div>"
    End If

    TryGetRadCtrlLinks = ot
End Function
Function TryGetAdmCtrlLinks(vst)
    Dim ot, rst, sql, href
    
    sql = "select top 1 AdmissionID, WardName, BedName from Admission "
    sql = sql & " inner join Ward on Ward.WardID=Admission.WardID "
    sql = sql & " inner join Bed on Bed.BedID=Admission.BedID "
    sql = sql & " where VisitationID='" & vst & "' order by AdmissionDate desc;"
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        href = "wpgAdmission.asp?PageMode=ProcessSelect&AdmissionID=" & rst.fields("AdmissionID")
        nm = rst.fields("WardName") & "<br/>&nbsp;&nbsp;&nbsp;&nbsp;" & rst.fields("BedName")
        ot = GetLink(href, nm, "#3c763d")
    ElseIf Len(vst) > 0 Then
        href = "wpgAdmission.asp?PageMode=AddNew&PullUpData=VisitationID||" & vst
        ot = GetLink(href, "Admit", "#a94442")
    End If
    TryGetAdmCtrlLinks = ot
End Function
Function GetDoctorName(specialistID)
    Dim ot, staffID
    
    staffID = GetComboNameFld("SystemUser", specialistID, "StaffID")
    If UCase(staffID) = "STF001" Then
        ot = "Doctor" 'GetComboName("JobSchedule", SpecialistID)
    Else
        ot = GetComboName("Staff", staffID)
    End If
    GetDoctorName = ot
End Function
Function GetControlLinks(fields)
    Dim html, href, linkText

    Select Case UCase(fields("Type"))
        Case "APPOINT_ONLY"
            If UCase(fields("PatientID")) = "P4" Then
                linkText = "Register Patient"
                href = "wpgPatient.asp?PageMode=AddNew&AppointmentID=" & fields("AppointmentID")
                html = html & "<div>" & GetLink(href, linkText, "#a94442 !important") & "</div>"
            Else
                'linkText = "Create Visit"
                'href = "wpgVisitation.asp?PageMode=AddNew&AppointmentID=" & fields("AppointmentID")
                
                linkText = "Create Visit"
                'href = "wpgVisitation.asp?PageMode=AddNew&AppointmentID=" & fields("AppointmentID")
                href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=CreateVisitListInsuredSchemes&PositionForTableName=Visitation&PatientID=" & server.URLEncode(fields("PatientID"))
                href = href & "&AppointmentID=" & fields("AppointmentID")
                html = html & "<div style='padding:3px;'>" & GetLink(href, linkText, "#2196f3!important") & "</div>"

                linkText = "Patient Ctrl Panel"
                href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientRCP&PositionForTableName=Patient&PatientID=" & server.URLEncode(fields("PatientID"))
                href = href & "&AppointmentID=" & fields("AppointmentID")
                html = html & "<div style='padding:3px;'>" & GetLink(href, linkText, "#2196f3!important") & "</div>"
            End If
        Case "VISIT_ONLY"
            linkText = "View Bill"
            href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationBill&PositionForTableName=Visitation&VisitationID=" & fields("VisitationID")
            html = html & "<div>" & GetLink(href, linkText, "#3c763d") & "</div>"
            
            linkText = "Patient Ctrl Panel"
            href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientRCP&PositionForTableName=Patient&PatientID=" & server.URLEncode(fields("PatientID"))
            href = href & "&AppointmentID=" & fields("AppointmentID")
            html = html & "<div style='padding:3px;'>" & GetLink(href, linkText, "#2196f3!important") & "</div>"
            
       Case "APPOINT_VISIT"
            linkText = "View Bill"
            href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=VisitationBill&PositionForTableName=Visitation&VisitationID=" & fields("VisitationID")
            html = html & "<div>" & GetLink(href, linkText, "#3c763d") & "</div>"
            
            linkText = "Patient Ctrl Panel"
            href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientRCP&PositionForTableName=Patient&PatientID=" & server.URLEncode(fields("PatientID"))
            href = href & "&AppointmentID=" & fields("AppointmentID")
            html = html & "<div style='padding:3px;'>" & GetLink(href, linkText, "#2196f3!important") & "</div>"
    End Select
    GetControlLinks = html
End Function
Function IIF(expression, trueVal, falseVal)
    If expression = True Then
        IIF = trueVal
    Else
        IIF = falseVal
    End If
End Function
Function JSONStringify(obj)
    Dim Key, tmpKey, value, tmp, ot

    Dim objType: objType = TypeName(obj)
    tmp = ""
    If objType = "Dictionary" Then
        For Each Key In obj.Keys()
            If tmp <> "" Then tmp = tmp & ", "
            tmpKey = """" & Key & """"
            tmp = tmp & tmpKey & ":" & JSONStringify(obj(Key))
        Next
        ot = "{" & tmp & "}"
    ElseIf IsArray(obj) Or objType = "ArrayList" Then
        For Each value In obj
            If tmp <> "" Then tmp = tmp & ", "
            tmp = tmp & JSONStringify(value)
        Next
        ot = "[" & tmp & "]"
    ElseIf objType = "String" Then
        tmp = Replace(obj, "\", "\\")
        tmp = Replace(tmp, """", "\""")
        tmp = Replace(tmp, vbCrLf, "\n")
        tmp = Replace(tmp, vbTab, "\n")
        ot = """" & tmp & """"
    ElseIf objType = "Boolean" Then
        ot = "" & LCase(obj) & ""
    ElseIf objType = "Byte" Then
        ot = CDbl(obj) 'Compatible with JSON.parse
    ElseIf objType = "Integer" Or objType = "Double" Or objType = "Long" Or objType = "Single" Or objType = "Currency" Then
        ot = obj
    ElseIf objType = "Empty" Or objType = "Null" Then
        ot = "null"
    ElseIf objType = "Date" Then
        ot = """" & obj & """"
    Else
        ot = """[Object " & (TypeName(obj)) & "]"""
    End If

    JSONStringify = ot
End Function



