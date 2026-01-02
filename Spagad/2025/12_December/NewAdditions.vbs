Function GetAppointmentCategorySelector(filterDict)
    Dim html, sql, rst, htmlSelect, isSelected

    Set rst = CreateObject("ADODB.RecordSet")

    sql = "SELECT DISTINCT AppointmentCat.AppointmentCatID, AppointmentCat.AppointmentCatName " & _
          "FROM Appointment " & _
          "INNER JOIN AppointmentCat ON AppointmentCat.AppointmentCatID = Appointment.AppointmentCatID " & _
          "WHERE 1=1 "
    
    If filterDict("AppointEndDayID") = "" Then
        sql = sql & " AND Appointment.AppointDayID = '" & filterDict("AppointStartDayID") & "' "
    Else
        sql = sql & " AND Appointment.AppointDayID BETWEEN '" & filterDict("AppointStartDayID") & "' AND '" & filterDict("AppointEndDayID") & "' "
    End If
    
    If filterDict("SpecialistID") <> "" Then
        sql = sql & " AND Appointment.SpecialistID = '" & filterDict("SpecialistID") & "' "
    End If
    
    sql = sql & " ORDER BY AppointmentCat.AppointmentCatName"

    htmlSelect = "<select style='padding:5px;' onchange='processCategoryChange(this)'>"
    htmlSelect = htmlSelect & "<option value=''>All Categories</option>"

    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.recordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            isSelected = IIF(UCase(filterDict("CategoryID")) = UCase(rst.fields("AppointmentCatID")), "selected", "")
            htmlSelect = htmlSelect & "<option value=""" & rst.fields("AppointmentCatID") & """ " & isSelected & ">" & rst.fields("AppointmentCatName") & "</option>"
            rst.MoveNext
        Loop
        rst.Close
    End If
    htmlSelect = htmlSelect & "</select>"
    Set rst = Nothing

    html = "<div style='display:inline-block;margin-left:20px;'>Category:" & htmlSelect & "</div>"
    GetAppointmentCategorySelector = html
End Function