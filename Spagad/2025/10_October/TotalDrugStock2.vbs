'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

tableStyles
dispUserJobSchedules

Sub dispUserJobSchedules()
    Dim count, sql, rst, storeArray, numStores, i
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT DISTINCT DrugStoreID FROM DrugStockLevel"


    With rst
        .open sql, conn, 3, 4
        
        If .recordCount > 0 Then
            storeArray = .GetRows(.recordCount)
            numStores = UBound(storeArray, 2) + 1
            '.MoveFirst
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Drug</th>"
            
            For i = 0 To numStores - 1
                response.write "<th class='myth'>" & storeArray(0, i) & "</th>"
            Next
            
            response.write "</tr class='mytr'>"
        End If
        .Close
    End With

    Set rst2 = CreateObject("ADODB.Recordset")

    sql = "SELECT d.DrugName, "
    For i = 0 To numStores - 1
      sql = sql & " COALESCE(SUM(CASE WHEN ds.DrugStoreID = '" & storeArray(0, i) & "' THEN ds.AvailableQty ELSE 0 END), 0) AS '" & storeArray(0, i) & "', "
    Next 
    sql = sql & " SUM(ds.AvailableQty) AS [Total Drug Quantity]"
    sql = sql & " FROM Drug d"
    sql = sql & " INNER JOIN DrugStockLevel ds ON d.DrugID = ds.DrugID"
    sql = sql & " GROUP BY d.DrugName"
    sql = sql & " ORDER BY [Total Drug Quantity] DESC;"

    response.write sql

    With rst2
      .open sql, conn 3, 4

      If .recordCount > 0
         Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                
                For i = 0 To numStores - 1
                  response.write "<td class='mytd'>" & .fields(" & storeArray(0, i) & ") & "</td>"
                Next

                response.write "</tr class='mytr'>"

                .MoveNext
                count = count + 1
            Loop
      End If
    End rst2
    
    Set rst = Nothing
End Sub


Sub tableStyles()
    response.write "<style>"
        response.write ".mytable {"
        response.write "    width: 65vw;"
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


