'Some slight modifications at this point
Sub displayStockLevel()
    
    Dim sql, periodStart, periodEnd, datePeriod, count, drugStoreIDs
    datePeriod = Trim(request.querystring("Dateperiod"))
    drugStoreIDs = Trim(request.querystring("DrugStoreID"))
    If datePeriod <> "" Then
        dateArr = Split(datePeriod, "||")
        periodStart = dateArr(0)
        periodEnd = dateArr(1)
    End If
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    
    If selectedDrugStoreIDs <> "" Then
        idsArr = Split(selectedDrugStoreIDs, ",")
        For Each id In idsArr
            formattedIDs = formattedIDs & "'" & Trim(id) & "',"
        Next
        ' Remove the trailing comma
        formattedIDs = Left(formattedIDs, Len(formattedIDs) - 1)
    End If
    
    If (periodStart <> "" And periodEnd <> "") Then
        sql = "select drugstoreid,drugstorename, stocklevel, convert(varchar(20),stockdate,103) stockdate "
        sql = sql & "from dbo.fn_displayStockLevel('" & periodStart & "','" & periodEnd & "') order by convert(DATE,StockDate) desc"
    Else
        sql = "select drugstoreid,drugstorename, stocklevel,convert(varchar(20),stockdate,103) stockdate "
        sql = sql & "from dbo.fn_displayStockLevel('2018-01-01', '2018-12-31') order by convert(DATE,StockDate) desc"
    End If
     
    response.write "    <form>"
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
    response.write "   </form>"
    response.write " <br />"
    response.write "<script> "
    response.write "    function updateUrl() { "
    response.write "        const fromDate = document.getElementById('from').value; "
    response.write "        const toDate = document.getElementById('to').value; "
    
    response.write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp'; "
    response.write "        const params = new URLSearchParams({ "
    response.write "            PrintLayoutName: 'displayStockLevel', "
    response.write "            PositionForTableName: 'WorkingDay', "
    response.write "            WorkingDayID: '' ,"
    response.write "            Dateperiod: fromDate + '||' + toDate"
    response.write "        }); "
    response.write "        const newUrl = baseUrl + '?' + params.toString(); "
    response.write "        console.log(newUrl); "
    response.write "        window.location.href = newUrl; "
    response.write "    } "
    response.write "</script> "
    
     response.write "<h2>Stock Levels From " & periodStart & " To " & periodEnd & " </h2>"
     
    With rst
        .open sql, conn, 3, 4
        
        If .recordCount > 0 Then

            
        response.write "<!DOCTYPE html>" & vbCrLf
        response.write "<html lang=""en"">" & vbCrLf
        response.write "<head>" & vbCrLf
        response.write "    <meta charset=""UTF-8"" />" & vbCrLf
        response.write "    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />" & vbCrLf
        response.write "    <script src=""https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js""></script>" & vbCrLf
        response.write "    <title>Document</title>" & vbCrLf
    
        response.write "</head>" & vbCrLf
        response.write "<body>" & vbCrLf
        response.write "    <div id=""countries-wrapper"">" & vbCrLf
        response.write "        <select name=""countries"" id=""countries"" multiple>" & vbCrLf
    '    Response.Write "            <option value=""1"">Afghanistan</option>" & vbCrLf
    '    Response.Write "            <option value=""2"">Australia</option>" & vbCrLf
    '    Response.Write "            <option value=""3"">Germany</option>" & vbCrLf
    '    Response.Write "            <option value=""4"">Canada</option>" & vbCrLf
    '    Response.Write "            <option value=""5"">Russia</option>" & vbCrLf
    '    Response.Write "            <option value=""6"">Ghana</option>" & vbCrLf
    '    Response.Write "            <option value=""7"">Togo</option>" & vbCrLf
    '    Response.Write "            <option value=""8"">Ivory Coast</option>" & vbCrLf
    '    Response.Write "            <option value=""9"">Burkina Faso</option>" & vbCrLf
    '    Response.Write "            <option value=""10"">Nigeria</option>" & vbCrLf
    Do Until .EOF
        response.write " <option value=""" & .fields("DrugStoreName") & """>" & .fields("DrugStoreName") & "</option>" & vbCrLf
        .MoveNext
    Loop
    response.write "        </select>" & vbCrLf
    response.write "    </div>" & vbCrLf
    response.write "    <script>" & vbCrLf
    response.write "     const arr = []; "
    response.write "        new MultiSelectTag('countries', {" & vbCrLf
    response.write "            rounded: true, // default true" & vbCrLf
    response.write "            shadow: true, // default false" & vbCrLf
    response.write "            placeholder: 'Search', // default Search..." & vbCrLf
    response.write "            tagColor: {" & vbCrLf
    response.write "                textColor: '#327b2c'," & vbCrLf
    response.write "                borderColor: '#92e681'," & vbCrLf
    response.write "                bgColor: '#eaffe6'," & vbCrLf
    response.write "            }," & vbCrLf
    response.write "            onChange: function (values) {" & vbCrLf
    response.write "               arr.push(values); console.log(arr);" & vbCrLf
    response.write "            }," & vbCrLf
    response.write "        });" & vbCrLf
    response.write "    </script>" & vbCrLf
    response.write "</body>" & vbCrLf
    response.write "</html>" & vbCrLf
    
    .movefirst
    response.write "<table width='100%' cellspacing='0' cellpadding='2' border='1' class='mytable'>"
    response.write "<tr>"
        response.write "<th class='myth'> No. </th>"
        response.write "<th class='myth'> Drug Store ID </th>"
        response.write "<th class='myth'> Drug Store Name </th>"
        response.write "<th class='myth'> Stock Level </th>"
        response.write "<th class='myth'> Stock Date</th>"
    response.write "</tr>"
    Do While Not .EOF
        count = count + 1
        response.write "<tr>"
            response.write "<td class='mytd' align='center'>" & count & "</td>"
            response.write "<td class='mytd'>" & .fields("DrugStoreID") & "</td>"
            response.write "<td class='mytd'>" & .fields("DrugStoreName") & "</td>"
            response.write "<td class='mytd' align='center'>" & .fields("StockLevel") & "</td>"
            response.write "<td class='mytd' align='center'>" & .fields("StockDate") & "</td>"
        response.write "</tr>"
        .MoveNext
    Loop
    response.write "</table>"
End If
        
      
    End With
End Sub