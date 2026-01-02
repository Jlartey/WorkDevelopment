'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

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
    response.write "        <iframe id='modalIframe' style='width:100%; height: 80vh; border:none;'></iframe>"
    response.write "    </div>"
    response.write "</div>"
    
    response.write "<script>"
    response.write "        function openModal() {"
    response.write "            document.getElementById('myModal').style.display = 'block';"
    response.write "            document.getElementById('modalIframe').src = 'http://192.168.5.11/rmchms01/wpgPrtPrintLayoutAll.asp?PrintLayoutName=SponsorArrearsForm&PositionForTableName=WorkingDay&WorkingDayID=';"
    response.write "        }"
    response.write "        function closeModal() {"
    response.write "            document.getElementById('myModal').style.display = 'none';"
    response.write "            document.getElementById('modalIframe').src = '';"
    response.write "        }"
    response.write "        // Close modal if clicking outside the content area"
    response.write "        window.onclick = function(event) {"
    response.write "            var modal = document.getElementById('myModal');"
    response.write "            if (event.target == modal) {"
    response.write "                closeModal();"
    response.write "            }"
    response.write "        }"
    response.write "</script>"
End Sub

Sub dispSponsorArrears()
    Dim count, sql, rst
    count = 1

    Set rst = CreateObject("ADODB.Recordset")

    sql = "SELECT PerformVar16Name, Description, KeyPrefix From PerformVar16 ORDER BY Description DESC"
    
    With rst
        .open sql, conn, 3, 4
        
        If .recordCount > 0 Then
            
            response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
            response.write "<tr class='mytr'>"
            response.write "<th class='myth'>No.</th>"
            response.write "<th class='myth'>Sponsor</th>"
            response.write "<th class='myth'>Bill Month</th>"
            response.write "<th class='myth'>Amount Paid</th>"
            response.write "</tr class='mytr'>"

            Do While Not .EOF
                response.write "<tr class='mytr'>"
                response.write "<td class='mytd'>" & count & "</td>"
                response.write "<td class='mytd'>" & .fields("PerformVar16Name") & "</td>"
                response.write "<td class='mytd'>" & .fields("Description") & "</td>"
                response.write "<td class='mytd'>" & .fields("KeyPrefix") & "</td>"
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
        
        response.write ".modal {"
        response.write "    display: none;"
        response.write "    position: fixed;"
        response.write "    z-index: 1000;"
        response.write "    left: 0;"
        response.write "    top: 0;"
        response.write "    width: 100%;"
        response.write "    height: 100%;"
        response.write "    overflow: auto;"
        response.write "    background-color: rgba(0,0,0,0.4);"
        response.write "}"
        response.write ".modal-content {"
        response.write "    background-color: #fefefe;"
        response.write "    margin: 5% auto;"
        response.write "    padding: 0;"
        response.write "    border: 1px solid #888;"
        response.write "    width: 80%;"
        response.write "    max-width: 900px;"
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


