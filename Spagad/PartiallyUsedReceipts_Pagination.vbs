Sub dispPartiallyUsedReceipts_Pagination()
    Dim sql, periodStart, periodEnd, datePeriod, count, pageSize, currentPage, totalRecords, totalPages, offset
    datePeriod = Trim(Request.QueryString("Dateperiod"))
    currentPage = CLng(Request.QueryString("page"))
    pageSize = 10 ' Number of records per page

    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If
    
    Set rst = CreateObject("ADODB.RecordSet")

    If (periodStart <> "" And periodEnd <> "") Then
        sql = "SELECT COUNT(*) AS TotalRecords FROM dbo.fn_partiallyUsedReceipts('" & periodStart & "','" & periodEnd & "')"
    Else
        sql = "SELECT COUNT(*) AS TotalRecords FROM dbo.fn_partiallyUsedReceipts('2018-01-01','2018-01-02')"
    End If
    
    With rst
        .Open sql, conn, 3, 4
        totalRecords = .Fields("TotalRecords").Value
    End With
    
    totalPages = Int(totalRecords / pageSize) + IIf(totalRecords Mod pageSize = 0, 0, 1)
    If currentPage <= 0 Then currentPage = 1
    If currentPage > totalPages Then currentPage = totalPages
    
    offset = (currentPage - 1) * pageSize

    If (periodStart <> "" And periodEnd <> "") Then
        sql = "SELECT * FROM dbo.fn_partiallyUsedReceipts('" & periodStart & "','" & periodEnd & "') ORDER BY ReceiptID OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"
    Else
        sql = "SELECT * FROM dbo.fn_partiallyUsedReceipts('2018-01-01','2018-01-02') ORDER BY ReceiptID OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"
    End If
    
    response.write "    <div class='container' style='display: flex; align-items: center; justify-content: center'> "
    response.write "        <div> "
    response.write "            <label for='from'>From</label> "
    response.write "            <input type='date' name='from' id='from'> "
    response.write "        </div> "
    response.write "        <div> "
    response.write "            <label for='to' style='margin-left: 10px'>To</label> "
    response.write "            <input type='date' name='to' id='to'> "
    response.write "        </div> "
    response.write "        <div> "
    response.write "            <button type='button' onclick='updateUrl()'>Show Data</button> <br />"
    response.write "        </div>    "
    response.write "    </div> "
   
    response.write " <br />"
    response.write "<script> "
    response.write "    function updateUrl() { "
    response.write "        const fromDate = document.getElementById('from').value; "
    response.write "        const toDate = document.getElementById('to').value; "
    response.write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp'; "
    response.write "        const params = new URLSearchParams({ "
    response.write "            PrintLayoutName: 'PartiallyUsedReceiptsJoe', "
    response.write "            PositionForTableName: 'WorkingDay', "
    response.write "            WorkingDayID: '' ,"
    response.write "            Dateperiod: fromDate + '||' + toDate"
    response.write "        }); "
    response.write "        const newUrl = baseUrl + '?' + params.toString(); "
    response.write "        console.log(newUrl); "
    response.write "        window.location.href = newUrl; "
    response.write "    } "
    response.write "</script> "
    
    response.write "<h2>Partially Used Receipts From " & periodStart & " To " & periodEnd & " </h2>"
     
    With rst
        .open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            count = offset
            .MoveFirst
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr>"
                response.write "<th class='myth'> No. </th>"
                response.write "<th class='myth'> Receipt Name </th>"
                response.write "<th class='myth'> Receipt ID </th>"
                response.write "<th class='myth'> Remarks </th>"
                response.write "<th class='myth'> Receipt Date</th>"
            response.write "</tr>"
            Do While Not .EOF
                count = count + 1
                response.write "<tr>"
                    response.write "<td class='mytd' align='center'>" & count & "</td>"
                    response.write "<td class='mytd'>" & .Fields("ReceiptName") & "</td>"
                    response.write "<td class='mytd'>" & .Fields("ReceiptID") & "</td>"
                    response.write "<td class='mytd'>" & .Fields("Remarks") & "</td>"
                    response.write "<td class='mytd' align='center'>" & .Fields("ReceiptDate") & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
            response.write "</table>"
        End If
    End With
    
    response.write "<div class='pagination'>"
    If currentPage > 1 Then
        response.write "<a href='?Dateperiod=" & datePeriod & "&page=" & (currentPage - 1) & "'>Previous</a>"
    End If
    If currentPage < totalPages Then
        response.write "<a href='?Dateperiod=" & datePeriod & "&page=" & (currentPage + 1) & "'>Next</a>"
    End If
    response.write "</div>"
End Sub

Sub Styling()
    response.write "<style>"
    response.write " table {"
    response.write "     width: 65vw;"
    response.write "     font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
    response.write "     border-collapse: collapse;"
    response.write " }"
    
    response.write " .container {"
    response.write "    display: flex"
    response.write " } "
    
    response.write " .myth, .mytd {"
    response.write "     border: 1px solid #ddd;"
    response.write "     padding: 10px;"
    response.write " }"
    
    response.write " .mytd {"
    response.write "     text-alig: 1px solid #ddd;"
    response.write "     padding: 8px;"
    response.write " }"
    
    response.write "  tr:nth-child(even) {"
    response.write "    background-color: #f9f9f9;"
    response.write " } "
    
    response.write " .myth {"
    response.write "     background-color: #c2c2c2;"
    response.write "     color: black;"
    response.write "     text-align: center; "
    response.write "     text-transform: uppercase; "
    response.write "     font-size: 18px;"
    response.write " }"
    
    response.write "  button {"
    response.write "     background-color: #0236c4;"
    response.write "     border-radius: 5px;"
    response.write "     border: none;"
    response.write "     margin-left: 50px;"
    response.write "     padding: 5px 20px;"
    response.write "     color: white;"
    response.write "     cursor: pointer;"
    response.write "  }"
    
    response.write "  #to, #from {"
    response.write "    padding: 5px;"
    response.write "    border-radius: 5px;"
    response.write "    cursor: pointer;"
    response.write "  }"
    
    response.write " .pagination {"
    response.write "    text-align: center;"
    response.write "    margin: 20px 0;"
    response.write " }"
    
    response.write " .pagination a {"
    response.write "    margin: 0 5px;"
    response.write "    padding: 10px 15px;"
    response.write "    background-color: #f1f1f1;"
    response.write "    border: 1px solid #ccc;"
    response.write "    text-decoration: none;"
    response.write "    color: #333;"
    response.write " }"
    
    response.write " .pagination a:hover {"
    response.write "    background-color: #ddd;"
    response.write " }"
    
    response.write "</style>"
End Sub




Sub dispPartiallyUsedReceiptsPagination()
    Dim sql, periodStart, periodEnd, datePeriod, count, pageSize, currentPage, offset
    datePeriod = Trim(Request.QueryString("Dateperiod"))
    currentPage = CLng(Request.QueryString("page"))
    pageSize = 10 ' Number of records per page

    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If

    If currentPage <= 0 Then currentPage = 1
    offset = (currentPage - 1) * pageSize

    If (periodStart <> "" And periodEnd <> "") Then
        sql = "SELECT * FROM dbo.fn_partiallyUsedReceipts('" & periodStart & "','" & periodEnd & "') ORDER BY ReceiptID OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"
    Else
        sql = "SELECT * FROM dbo.fn_partiallyUsedReceipts('2018-01-01','2018-01-02') ORDER BY ReceiptID OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"
    End If

    Set rst = CreateObject("ADODB.RecordSet")

    response.write "    <div class='container' style='display: flex; align-items: center; justify-content: center'> "
    response.write "        <div> "
    response.write "            <label for='from'>From</label> "
    response.write "            <input type='date' name='from' id='from'> "
    response.write "        </div> "
    response.write "        <div> "
    response.write "            <label for='to' style='margin-left: 10px'>To</label> "
    response.write "            <input type='date' name='to' id='to'> "
    response.write "        </div> "
    response.write "        <div> "
    response.write "            <button type='button' onclick='updateUrl()'>Show Data</button> <br />"
    response.write "        </div>    "
    response.write "    </div> "
   
    response.write " <br />"
    response.write "<script> "
    response.write "    function updateUrl() { "
    response.write "        const fromDate = document.getElementById('from').value; "
    response.write "        const toDate = document.getElementById('to').value; "
    response.write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp'; "
    response.write "        const params = new URLSearchParams({ "
    response.write "            PrintLayoutName: 'PartiallyUsedReceiptsJoe', "
    response.write "            PositionForTableName: 'WorkingDay', "
    response.write "            WorkingDayID: '' ,"
    response.write "            Dateperiod: fromDate + '||' + toDate"
    response.write "        }); "
    response.write "        const newUrl = baseUrl + '?' + params.toString(); "
    response.write "        console.log(newUrl); "
    response.write "        window.location.href = newUrl; "
    response.write "    } "
    response.write "</script> "
    
    response.write "<h2>Partially Used Receipts From " & periodStart & " To " & periodEnd & " </h2>"

    With rst
        .Open sql, conn, 3, 4
        
        If .RecordCount > 0 Then
            count = offset
            .MoveFirst
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1'>"
            response.write "<tr>"
                response.write "<th class='myth'> No. </th>"
                response.write "<th class='myth'> Receipt Name </th>"
                response.write "<th class='myth'> Receipt ID </th>"
                response.write "<th class='myth'> Remarks </th>"
                response.write "<th class='myth'> Receipt Date</th>"
            response.write "</tr>"
            Do While Not .EOF
                count = count + 1
                response.write "<tr>"
                    response.write "<td class='mytd' align='center'>" & count & "</td>"
                    response.write "<td class='mytd'>" & .Fields("ReceiptName") & "</td>"
                    response.write "<td class='mytd'>" & .Fields("ReceiptID") & "</td>"
                    response.write "<td class='mytd'>" & .Fields("Remarks") & "</td>"
                    response.write "<td class='mytd' align='center'>" & .Fields("ReceiptDate") & "</td>"
                response.write "</tr>"
                .MoveNext
            Loop
            response.write "</table>"

            response.write "<div class='pagination'>"
            If currentPage > 1 Then
                response.write "<a href='?Dateperiod=" & datePeriod & "&page=" & (currentPage - 1) & "'>Previous</a>"
            End If
            response.write "<a href='?Dateperiod=" & datePeriod & "&page=" & (currentPage + 1) & "'>Next</a>"
            response.write "</div>"
        Else
            response.write "<p>No records found.</p>"
        End If
    End With
End Sub

Sub Styling()
    response.write "<style>"
    response.write " table {"
    response.write "     width: 65vw;"
    response.write "     font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande', 'Lucida Sans', Arial, sans-serif; "
    response.write "     border-collapse: collapse;"
    response.write " }"
    
    response.write " .container {"
    response.write "    display: flex"
    response.write " } "
    
    response.write " .myth, .mytd {"
    response.write "     border: 1px solid #ddd;"
    response.write "     padding: 10px;"
    response.write " }"
    
    response.write " .mytd {"
    response.write "     text-alig: 1px solid #ddd;"
    response.write "     padding: 8px;"
    response.write " }"
    
    response.write "  tr:nth-child(even) {"
    response.write "    background-color: #f9f9f9;"
    response.write " } "
    
    response.write " .myth {"
    response.write "     background-color: #c2c2c2;"
    response.write "     color: black;"
    response.write "     text-align: center; "
    response.write "     text-transform: uppercase; "
    response.write "     font-size: 18px;"
    response.write " }"
    
    response.write "  button {"
    response.write "     background-color: #0236c4;"
    response.write "     border-radius: 5px;"
    response.write "     border: none;"
    response.write "     margin-left: 50px;"
    response.write "     padding: 5px 20px;"
    response.write "     color: white;"
    response.write "     cursor: pointer;"
    response.write "  }"
    
    response.write "  #to, #from {"
    response.write "    padding: 5px;"
    response.write "    border-radius: 5px;"
    response.write "    cursor: pointer;"
    response.write "  }"
    
    response.write " .pagination {"
    response.write "    text-align: center;"
    response.write "    margin: 20px 0;"
    response.write " }"
    
    response.write " .pagination a {"
    response.write "    margin: 0 5px;"
    response.write "    padding: 10px 15px;"
    response.write "    background-color: #f1f1f1;"
    response.write "    border: 1px solid #ccc;"
    response.write "    text-decoration: none;"
    response.write "    color: #333;"
    response.write " }"
    
    response.write " .pagination a:hover {"
    response.write "    background-color: #ddd;"
    response.write " }"
    
    response.write "</style>"
End Sub
