
Dim rst, sysuserid, jsonArray, resp, values

sysuserid = uName

Set resp = CreateObject("Scripting.Dictionary")

Set rst = CreateObject("ADODB.RecordSet")

Set jsonArray = CreateObject("System.Collections.ArrayList")

Set values = CreateObject("Scripting.Dictionary")

resp.Add "success", False


sql = ""
sql = sql & "SELECT DISTINCT PerformVar12.KeyPrefix as 'JobScheduleID' "
sql = sql & "FROM PerformVar12 "
sql = sql & "JOIN SystemUser ON SystemUser.StaffID = PerformVar12.Description "
sql = sql & "WHERE PerformVar12.Description ='" & GetComboNameFld("SystemUser", sysuserid, "StaffID") & "' AND SystemUser.UserStatusID = 'UST001' "
sql = sql & "and SystemUser.JobScheduleID <> PerformVar12.KeyPrefix"

On Error Resume Next

    rst.open qryPro.FltQry(sql), conn, 3, 4

    values.Add "currentProfile", GetComboName("JobSchedule", jSchd)

    If rst.RecordCount > 0 Then

        rst.MoveFirst

        Do While Not rst.EOF

            Dim obj

            Set obj = CreateObject("Scripting.Dictionary")

            obj.Add "id", rst.fields("JobScheduleID").value

            obj.Add "name", GetComboName("JobSchedule", rst.fields("JobScheduleID").value)

            jsonArray.Add obj

            rst.MoveNext

        Loop

    End If

    values.Add "otherProfiles", jsonArray
    'values.add "query", sql

     resp("success") = True

On Error GoTo 0



rst.Close

Set rst = Nothing



resp.Add "values", values



response.Clear

response.contentType = "application/json"

response.write gUtils.JSONStringify(resp)


