<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim datePeriod, periodStart, periodEnd, dateArr
datePeriod = Trim(Request.queryString("PrintFilter"))
periodStart = ""
periodEnd = ""
If datePeriod <> "" Then
    dateArr = Split(datePeriod, "||")
    If UBound(dateArr) >= 1 Then
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If
End If

' Default to full year if no dates provided
If periodStart = "" Then periodStart = "2025-01-01"
If periodEnd = "" Then periodEnd = "2025-12-31"

tableStyles
ImagingRequests

Sub ImagingRequests()
    Dim count, sql, rst
    count = 1
    Set rst = CreateObject("ADODB.Recordset")

    sql = "WITH CombinedData AS (" & vbCrLf
    sql = sql & "    SELECT DoctorName, LabtestName" & vbCrLf
    sql = sql & "    FROM Investigation2" & vbCrLf
    sql = sql & "    JOIN LabTest ON Investigation2.LabtestID = LabTest.LabTestID" & vbCrLf
    sql = sql & "    WHERE Investigation2.TestCategoryID = 'B19'" & vbCrLf
    sql = sql & "      AND LabTest.TestStatusID = 'TST001'" & vbCrLf
    sql = sql & "      AND RequestDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'" & vbCrLf
    sql = sql & "" & vbCrLf
    sql = sql & "    UNION ALL" & vbCrLf
    sql = sql & "" & vbCrLf
    sql = sql & "    SELECT DoctorName, LabtestName" & vbCrLf
    sql = sql & "    FROM Investigation" & vbCrLf
    sql = sql & "    JOIN LabTest ON Investigation.LabtestID = LabTest.LabTestID" & vbCrLf
    sql = sql & "    WHERE Investigation.TestCategoryID = 'B19'" & vbCrLf
    sql = sql & "      AND LabTest.TestStatusID = 'TST001'" & vbCrLf
    sql = sql & "      AND RequestDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'" & vbCrLf
    sql = sql & ")" & vbCrLf
    sql = sql & "SELECT " & vbCrLf
    sql = sql & "    DoctorName," & vbCrLf
    sql = sql & "    SUM(CASE WHEN LabtestName LIKE '%ray%' THEN 1 ELSE 0 END) AS Xrays," & vbCrLf
    sql = sql & "    SUM(CASE WHEN LabtestName NOT LIKE '%ray%' THEN 1 ELSE 0 END) AS Scans," & vbCrLf
    sql = sql & "    COUNT(*) AS Total" & vbCrLf
    sql = sql & "FROM CombinedData" & vbCrLf
    sql = sql & "GROUP BY DoctorName" & vbCrLf
    sql = sql & "ORDER BY DoctorName;"

    ' Optional: uncomment to debug SQL
    ' response.write "<pre>" & Server.HTMLEncode(sql) & "</pre>"

    Dim totalXrays, totalScans, grandTotal
    totalXrays = 0
    totalScans = 0
    grandTotal = 0

    With rst
        .open sql, conn, 0, 1
       
        If Not (.EOF And .BOF) Then
           
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Doctor</th>"
            response.write "<th class='myth'>X-RAYS</th>"
            response.write "<th class='myth'>SCANS</th>"
            response.write "<th class='myth'>TOTAL</th>"
            response.write "</tr>"

            Do While Not .EOF
                Dim xrays, scans, doctorTotal
                xrays = Nz(.fields("Xrays"), 0)
                scans = Nz(.fields("Scans"), 0)
                doctorTotal = xrays + scans

                totalXrays = totalXrays + xrays
                totalScans = totalScans + scans
                grandTotal = grandTotal + doctorTotal

                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("DoctorName") & "</td>"
                response.write "<td class='mytd' style='text-align:center;'>" & xrays & "</td>"
                response.write "<td class='mytd' style='text-align:center;'>" & scans & "</td>"
                response.write "<td class='mytd' style='text-align:center;font-weight:bold;'>" & doctorTotal & "</td>"
                response.write "</tr>"

                .MoveNext
                count = count + 1
            Loop

            ' Total row
            response.write "<tr class='mytr' style='background-color:#e6e6e6; font-weight:bold;'>"
            response.write "<td class='mytd' colspan='2' style='text-align:right;'>GRAND TOTAL</td>"
            response.write "<td class='mytd' style='text-align:center;'>" & totalXrays & "</td>"
            response.write "<td class='mytd' style='text-align:center;'>" & totalScans & "</td>"
            response.write "<td class='mytd' style='text-align:center;'>" & grandTotal & "</td>"
            response.write "</tr>"

            response.write "</table>"
        Else
            response.write "<h1>No records found</h1>"
        End If
       
        .Close
    End With
   
    Set rst = Nothing
End Sub

' Helper function for Nz (Null to Zero)
Function Nz(val, default)
    If IsNull(val) Then
        Nz = default
    Else
        Nz = val
    End If
End Function

Sub tableStyles()
    response.write "<style>"
        response.write ".mytable {"
        response.write " width: 65vw;"
        response.write " border-collapse: collapse;"
        response.write " margin: 20px 0;"
        response.write " font-size: 16px;"
        response.write " font-family: Arial, sans-serif;"
        response.write "}"
        response.write ".mytable, .myth, .mytd {"
        response.write " border: 1px solid #dddddd;"
        response.write "}"
        response.write ".myth, .mytd {"
        response.write " padding: 12px;"
        response.write " text-align: left;"
        response.write "}"
        response.write ".myth {"
        response.write " background-color: #f2f2f2;"
        response.write " color: #333;"
        response.write " font-weight: bold;"
        response.write " text-transform: uppercase;"
        response.write "}"
        response.write ".mytr:nth-child(even) {"
        response.write " background-color: #f9f9f9;"
        response.write "}"
        response.write ".mytr:hover {"
        response.write " background-color: #f1f1f1;"
        response.write "}"
        response.write "h1 {"
        response.write " font-size: 18px;"
        response.write " color: #555;"
        response.write " font-family: Arial, sans-serif;"
        response.write " margin: 20px 0;"
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
'>
