'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim billMonth
billMonth = Request.QueryString("PrintFilter")

tableStyles
AddNewButton
dispSponsorArrears

Sub AddNewButton()
    response.write "<div style='display: flex; justify-content: flex-end'>"
        response.write "<button class='mybutton' onclick='openModal()'>Add New</button>"
    response.write "</div>"
    
    response.write "<div id='myModal' class='modal'>"
    response.write "    <div class='modal-content'>"
    response.write "        <span class='close' onclick='closeModal()'>&times;</span>"
    response.write "        <iframe id='modalIframe' style='width:100%; height: 65vh; border: none; scrolling=""no""'></iframe>"
    response.write "    </div>"
    response.write "</div>"
    
    response.write "<div id='myModal2' class='modal'>"
    response.write "    <div class='modal-content'>"
    response.write "        <span class='close' onclick='closeModal()'>&times;</span>"
    response.write "        <iframe id='modalIframe2' style='width:100%; height: 65vh; border: none; scrolling=""no""'></iframe>"
    response.write "    </div>"
    response.write "</div>"
    
    
    response.write vbCrLf & "<script>"
    response.write vbCrLf & "        function openModal() {"
    response.write vbCrLf & "            document.getElementById('myModal').style.display = 'block';"
    response.write vbCrLf & "            const iframe = document.getElementById('modalIframe');"
    response.write vbCrLf & "            iframe.onload = function() {"
    response.write vbCrLf & "                const innerDoc = iframe.contentDocument || iframe.contentWindow.document;"
    response.write vbCrLf & "                innerDoc.documentElement.style.overflow = 'hidden';"
    response.write vbCrLf & "                innerDoc.body.style.overflow = 'hidden';"
    response.write vbCrLf & "            };"
    response.write vbCrLf & "            iframe.src = 'http://192.168.5.11/rmchms01/wpgPrtPrintLayoutAll.asp?PrintLayoutName=SponsorArrearsForm&PositionForTableName=WorkingDay&WorkingDayID=';"
    response.write vbCrLf & "        }"
    response.write vbCrLf & "        function closeModal() {"
    response.write vbCrLf & "            document.getElementById('myModal').style.display = 'none';"
    response.write vbCrLf & "            document.getElementById('modalIframe').src = '';"
    response.write vbCrLf & "            location.reload(); "
    response.write vbCrLf & "        }"
    response.write vbCrLf & "        window.onclick = function(event) {"
    response.write vbCrLf & "            var modal = document.getElementById('myModal');"
    response.write vbCrLf & "            if (event.target == modal) {"
    response.write vbCrLf & "                closeModal();"
    response.write vbCrLf & "            }"
    response.write vbCrLf & "        }"
    
    response.write vbCrLf & "        function debitArrears() {"
    response.write vbCrLf & "            document.getElementById('myModal2').style.display = 'block';"
    response.write vbCrLf & "            const iframe = document.getElementById('modalIframe2');"
    response.write vbCrLf & "            iframe.onload = function() {"
    response.write vbCrLf & "                const innerDoc = iframe.contentDocument || iframe.contentWindow.document;"
    response.write vbCrLf & "                innerDoc.documentElement.style.overflow = 'hidden';"
    response.write vbCrLf & "                innerDoc.body.style.overflow = 'hidden';"
    response.write vbCrLf & "            };"
    response.write vbCrLf & "            iframe.src = 'https://www.livescore.com/';"
    response.write vbCrLf & "        }"
    response.write vbCrLf & "</script>"
End Sub

Sub dispSponsorArrears()
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT s.SponsorName, bm.BillMonthName, p.KeyPrefix "
    sql = sql & "FROM PerformVar16 p "
    sql = sql & "JOIN BillMonth bm "
    sql = sql & "ON p.Description = bm.BillMonthID "
    sql = sql & "JOIN Sponsor s "
    sql = sql & "ON p.PerformVar16Name = s.SponsorID "
    sql = sql & "WHERE bm.BillMonthID = '" & billMonth & "' "
    sql = sql & "ORDER BY p.Description DESC "
    
    response.write sql
    With rst
        .open qryPro.FltQry(sql), conn, 3, 4
        
        If .recordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Sponsor</th>"
            response.write "<th class='myth'>Bill Month</th>"
            response.write "<th class='myth'>Amount To Be Paid</th>"
            response.write "<th class='myth'>Amount Paid</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("SponsorName") & "</td>"
                response.write "<td class='mytd'>" & .fields("BillMonthName") & "</td>"
                response.write "<td class='mytd'>" & .fields("KeyPrefix") & "</td>"
                response.write "<td class='mytd'> <button onclick='debitArrears()'>Enter Amount Paid</button> </td>"
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
        
        response.write " .mybutton {"
        response.write "       width: 9.375rem;"
        response.write "       padding: 0.5rem;;"
        response.write "       border-radius: 10px;"
        response.write "       background-color: blue;"
        response.write "       font-weight: bold;"
        response.write "       color: white;"
        response.write "       border: none;"
        response.write "       cursor: pointer;"
        response.write "       transition: transform 0.1s ease, box-shadow 0.1s ease; "
        response.write "   }"
    
        response.write "  .mybutton:active {"
        response.write "     transform: scale(0.95);"
        response.write "     box-shadow: 0 4px 8px rgba(0, 0, 0, 0.3);"
        response.write "  }"
        
        response.write ".modal {"
        response.write "    display: none;"
        response.write "    position: fixed;"
        response.write "    z-index: 1000;"
        response.write "    left: 0;"
        response.write "    top: 0;"
        response.write "    width: 100%;"
        response.write "    height: 100%;"
        response.write "    background-color: rgba(0,0,0,0.4);"
        response.write "    overflow: hidden;"
        response.write "}"
        
        response.write ".modal-content {"
        response.write "    background-color: #fefefe;"
        response.write "    margin: 5% auto;"
        response.write "    padding: 0;"
        response.write "    border: 1px solid #888;"
        response.write "    width: 50%;"
        response.write "    overflow: hidden;"
        response.write "    max-width: 450px;"
        response.write "    box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2);"
        response.write "    position: relative;"
        response.write "}"
        
        response.write ".close {"
        response.write "    color: #aaa;"
        response.write "    float: right;"
        response.write "    font-size: 28px;"
        response.write "    font-weight: bold;"
        response.write "    position: absolute;"
        response.write "    right: 10px;"
        response.write "    top: 5px;"
        response.write "    cursor: pointer;"
        response.write "}"
        response.write ".close:hover, .close:focus {"
        response.write "    color: black;"
        response.write "    text-decoration: none;"
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
