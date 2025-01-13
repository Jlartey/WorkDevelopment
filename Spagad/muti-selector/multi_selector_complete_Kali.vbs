Sub DrugStockLevelSub()
            Dim rst, sql, cnt
            Set rst = CreateObject("ADODB.RecordSet")
            
            sql = " select Drugstore.DrugStoreID , DrugStoreName, COUNT (DrugID) [Drug Stock], convert (varchar(20), StockDate1, 103) [Stock Date] from DrugStockLevel"
            sql = sql & " join drugstore on Drugstore.DrugStoreID = DrugStockLevel.DrugStoreID"
            'sql = sql & " where Drugstore.DrugStoreID = '" & drugSTid & "'"
            sql = sql & " and convert (date, StockDate1) "
             If (periodStart <> "" And periodEnd <> "") Then
                 sql = sql & "BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
            Else
                 sql = sql & "BETWEEN '2018-11-01' AND '2018-11-01'"
            End If
            sql = sql & " group by Drugstore.DrugStoreID, Drugstore.DrugStoreName,StockDate1"
            sql = sql & " order by [Drug Stock] desc ,[Stock Date] "
            
                    
                    

'Response.Write (sql)
    
   'Response.Write sql
  Response.Write "<div style='text-align: center;'><b>Drug Stock by  Pharmacy Stores Report for Date Period:</b></div>"
  Response.Write "<div style='text-align: center;'><br>" & datePeriod & "</div>"
    
    With rst
        .open sql, conn, 3, 4
        If .RecordCount > 0 Then
                Response.Write "<!DOCTYPE html>" & vbCrLf
                        Response.Write "<html lang=""en"">" & vbCrLf
                        Response.Write "<head>" & vbCrLf
                        Response.Write "    <meta charset=""UTF-8"" />" & vbCrLf
                        Response.Write "    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />" & vbCrLf
                        Response.Write "    <script src=""https://cdn.jsdelivr.net/gh/habibmhamadi/multi-select-tag@3.0.1/dist/js/multi-select-tag.js""></script>" & vbCrLf
                        Response.Write "    <title>Document</title>" & vbCrLf
                        Response.Write "    <style>" & vbCrLf
                        Response.Write "        .mult-select-tag {" & vbCrLf
                        Response.Write "            display: flex;" & vbCrLf
                        Response.Write "            width: 300px;" & vbCrLf
                        Response.Write "            flex-direction: column;" & vbCrLf
                        Response.Write "            align-items: center;" & vbCrLf
                        Response.Write "            position: relative;" & vbCrLf
                        Response.Write "            --tw-shadow: 0 1px 3px 0 rgb(0 0 0 / 0.1), 0 1px 2px -1px rgb(0 0 0 / 0.1);" & vbCrLf
                        Response.Write "            --tw-shadow-color: 0 1px 3px 0 var(--tw-shadow-color), 0 1px 2px -1px var(--tw-shadow-color);" & vbCrLf
                        Response.Write "            --border-color: rgb(218, 221, 224);" & vbCrLf
                        Response.Write "            font-family: Verdana, sans-serif;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .wrapper {" & vbCrLf
                        Response.Write "            width: 100%;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .body {" & vbCrLf
                        Response.Write "            display: flex;" & vbCrLf
                        Response.Write "            border: 1px solid var(--border-color);" & vbCrLf
                        Response.Write "            background: #fff;" & vbCrLf
                        Response.Write "            min-height: 2.15rem;" & vbCrLf
                        Response.Write "            width: 100%;" & vbCrLf
                        Response.Write "            min-width: 14rem;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .input-container {" & vbCrLf
                        Response.Write "            display: flex;" & vbCrLf
                        Response.Write "            flex-wrap: wrap;" & vbCrLf
                        Response.Write "            flex: 1 1 auto;" & vbCrLf
                        Response.Write "            padding: 0.1rem;" & vbCrLf
                        Response.Write "            align-items: center;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .input-body {" & vbCrLf
                        Response.Write "            display: flex;" & vbCrLf
                        Response.Write "            width: 100%;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .input {" & vbCrLf
                        Response.Write "            flex: 1;" & vbCrLf
                        Response.Write "            background: 0 0;" & vbCrLf
                        Response.Write "            border-radius: 0.25rem;" & vbCrLf
                        Response.Write "            padding: 0.45rem;" & vbCrLf
                        Response.Write "            margin: 10px;" & vbCrLf
                        Response.Write "            color: #2d3748;" & vbCrLf
                        Response.Write "            outline: 0;" & vbCrLf
                        Response.Write "            border: 1px solid var(--border-color);" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .btn-container {" & vbCrLf
                        Response.Write "            color: #e2ebf0;" & vbCrLf
                        Response.Write "            padding: 0.5rem;" & vbCrLf
                        Response.Write "            display: flex;" & vbCrLf
                        Response.Write "            border-left: 1px solid var(--border-color);" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag button {" & vbCrLf
                        Response.Write "            cursor: pointer;" & vbCrLf
                        Response.Write "            width: 100%;" & vbCrLf
                        Response.Write "            color: #718096;" & vbCrLf
                        Response.Write "            outline: 0;" & vbCrLf
                        Response.Write "            height: 100%;" & vbCrLf
                        Response.Write "            border: none;" & vbCrLf
                        Response.Write "            padding: 0;" & vbCrLf
                        Response.Write "            background: 0 0;" & vbCrLf
                        Response.Write "            background-image: none;" & vbCrLf
                        Response.Write "            -webkit-appearance: none;" & vbCrLf
                        Response.Write "            text-transform: none;" & vbCrLf
                        Response.Write "            margin: 0;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag button:first-child {" & vbCrLf
                        Response.Write "            width: 1rem;" & vbCrLf
                        Response.Write "            height: 90%;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .drawer {" & vbCrLf
                        Response.Write "            position: absolute;" & vbCrLf
                        Response.Write "            background: #fff;" & vbCrLf
                        Response.Write "            max-height: 15rem;" & vbCrLf
                        Response.Write "            z-index: 40;" & vbCrLf
                        Response.Write "            top: 98%;" & vbCrLf
                        Response.Write "            width: 100%;" & vbCrLf
                        Response.Write "            overflow-y: scroll;" & vbCrLf
                        Response.Write "            border: 1px solid var(--border-color);" & vbCrLf
                        Response.Write "            border-radius: 0.25rem;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag ul {" & vbCrLf
                        Response.Write "            list-style-type: none;" & vbCrLf
                        Response.Write "            padding: 0.5rem;" & vbCrLf
                        Response.Write "            margin: 0;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag ul li {" & vbCrLf
                        Response.Write "            padding: 0.5rem;" & vbCrLf
                        Response.Write "            border-radius: 0.25rem;" & vbCrLf
                        Response.Write "            cursor: pointer;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag ul li:hover {" & vbCrLf
                        Response.Write "            background: rgb(243 244 246);" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .item-container {" & vbCrLf
                        Response.Write "            display: flex;" & vbCrLf
                        Response.Write "            justify-content: center;" & vbCrLf
                        Response.Write "            align-items: center;" & vbCrLf
                        Response.Write "            padding: 0.2rem 0.4rem;" & vbCrLf
                        Response.Write "            margin: 0.2rem;" & vbCrLf
                        Response.Write "            font-weight: 500;" & vbCrLf
                        Response.Write "            border: 1px solid;" & vbCrLf
                        Response.Write "            border-radius: 9999px;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .item-label {" & vbCrLf
                        Response.Write "            max-width: 100%;" & vbCrLf
                        Response.Write "            line-height: 1;" & vbCrLf
                        Response.Write "            font-size: 0.75rem;" & vbCrLf
                        Response.Write "            font-weight: 400;" & vbCrLf
                        Response.Write "            flex: 0 1 auto;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .item-close-container {" & vbCrLf
                        Response.Write "            display: flex;" & vbCrLf
                        Response.Write "            flex: 1 1 auto;" & vbCrLf
                        Response.Write "            flex-direction: row-reverse;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .item-close-svg {" & vbCrLf
                        Response.Write "            width: 1rem;" & vbCrLf
                        Response.Write "            margin-left: 0.5rem;" & vbCrLf
                        Response.Write "            height: 1rem;" & vbCrLf
                        Response.Write "            cursor: pointer;" & vbCrLf
                        Response.Write "            border-radius: 9999px;" & vbCrLf
                        Response.Write "            display: block;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .hidden {" & vbCrLf
                        Response.Write "            display: none;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .shadow {" & vbCrLf
                        Response.Write "            box-shadow: var(--tw-ring-offset-shadow, 0 0 #0000), var(--tw-ring-shadow, 0 0 #0000), var(--tw-shadow);" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "        .mult-select-tag .rounded {" & vbCrLf
                        Response.Write "            border-radius: 0.375rem;" & vbCrLf
                        Response.Write "        }" & vbCrLf
                        Response.Write "    </style>" & vbCrLf
                        Response.Write "</head>" & vbCrLf
                        Response.Write "<body>" & vbCrLf
                        Response.Write "    <div id=""drugstoreids-wrapper"">" & vbCrLf
                        Response.Write "        <select name=""drugstoreids"" id=""drugstoreids"" multiple>" & vbCrLf
                        
                        Do Until .EOF
                            Response.Write "            <option value=""" & .fields("DrugStoreName") & """>" & .fields("DrugStoreName") & "</option>" & vbCrLf
                            .MoveNext
                        Loop
                        
                        Response.Write "        </select>" & vbCrLf
                        Response.Write "        <input type=""hidden"" name=""selectedDrugstores"" id=""selectedDrugstores"" />" & vbCrLf
                        Response.Write "        <button type=""submit"">Submit</button>" & vbCrLf
                        Response.Write "    </div>" & vbCrLf
                        Response.Write "    <script>" & vbCrLf
                        Response.Write "        document.querySelector('form').addEventListener('submit', function() {" & vbCrLf
                        Response.Write "            var selectedOptions = Array.from(document.querySelectorAll('#drugstoreids option:checked'));" & vbCrLf
                        Response.Write "            var selectedValues = selectedOptions.map(function(option) { return option.value; });" & vbCrLf
                        Response.Write "            document.getElementById('selectedDrugstores').value = selectedValues.join(',');" & vbCrLf
                        Response.Write "        });" & vbCrLf
                        Response.Write "        new MultiSelectTag('drugstoreids', {" & vbCrLf
                        Response.Write "            rounded: true, // default true" & vbCrLf
                        Response.Write "            shadow: true, // default false" & vbCrLf
                        Response.Write "            placeholder: 'Search', // default Search..." & vbCrLf
                        Response.Write "            tagColor: {" & vbCrLf
                        Response.Write "                textColor: '#327b2c'," & vbCrLf
                        Response.Write "                borderColor: '#92e681'," & vbCrLf
                        Response.Write "                bgColor: '#eaffe6'," & vbCrLf
                        Response.Write "            }," & vbCrLf
                        Response.Write "            onChange: function (values) {" & vbCrLf
                        Response.Write "                console.log(values);" & vbCrLf
                        Response.Write "            }," & vbCrLf
                        Response.Write "        });" & vbCrLf
                        Response.Write "    </script>" & vbCrLf
                        Response.Write "</body>" & vbCrLf
                        Response.Write "</html>" & vbCrLf

            
            
            

        
                        Response.Write "<form id='dateForm'> "
                        Response.Write "<div class='form-container'>"
                        Response.Write "    <div class='date-container'>"
                        Response.Write "        <div>"
                        Response.Write "            <label for='from'>From</label>"
                        Response.Write "            <input type='date' name='from' id='from'>"
                        Response.Write "        </div>"
                        Response.Write "        <div>"
                        Response.Write "            <label for='to'>To</label>"
                        Response.Write "            <input type='date' name='to' id='to'>"
                        Response.Write "        </div>"
                        Response.Write "    </div>"
                        Response.Write "    <div class='button-container'>"
                        Response.Write "        <button type='button' onclick='updateUrl()'>PROCESS</button>"
                        Response.Write "    </div>"
                        Response.Write "</div>"
                        Response.Write "</form> "
                        Response.Write "<script> "
                        Response.Write "    function updateUrl() { "
                        Response.Write "        const fromDate = document.getElementById('from').value; "
                        Response.Write "        const toDate = document.getElementById('to').value; "
                        Response.Write "        const baseUrl = 'http://192.168.5.11/thhms15/wpgPrtPrintLayoutAll.asp'; "
                        Response.Write "        const params = new URLSearchParams({ "
                        Response.Write "            PrintLayoutName: 'DrugStockLevelRPT', "
                        Response.Write "            PositionForTableName: 'WorkingDay', "
                        Response.Write "            WorkingDayID: '' ,"
                        Response.Write "            Dateperiod: ${fromDate}||${toDate}"
                        Response.Write "        }); "
                        Response.Write "        const newUrl = ${baseUrl}?${params.toString()}; "
                        Response.Write "        window.location.href = newUrl; "
                        Response.Write "    } "
                        Response.Write "</script> "
                        .movefirst
                
                   
                
                Response.Write "<table width='130%' cellspacing='0' cellpadding='2' border='1'>"
                
                Dim style

                style = "<style>"
                style = style & "body { font-family: Arial, sans-serif; background-color: #f4f4f4; margin: 0; padding: 20px; }"
                style = style & ".container { max-width: 600px; margin: auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); }"
                style = style & "h1 { text-align: center; color: #333; }"
                style = style & "label { display: block; margin-top: 10px; color: #555; }"
                style = style & "input[type='date'] { padding: 10px; margin-top: 5px; border: 1px solid #ddd; border-radius: 4px; }"
                style = style & "table { width: 80%; border-collapse: collapse; margin: 20px auto; border-radius: 8px; overflow: hidden; }"
                style = style & "th, td { padding: 8px; border: 1px solid #ddd; text-align: left; }"
                style = style & "th { background-color: #e0f7ff; }"
                style = style & "tr:nth-child(even) { background-color: #f2faff; }"
                style = style & "tr:hover { background-color: #d0e9ff; }"
                style = style & "caption { caption-side: top; font-size: 1.5em; padding: 10px; color: #333; }"
                style = style & ".form-container {"
                style = style & "    display: flex;"
                style = style & "    flex-direction: column;"
                style = style & "    justify-content: center;"
                style = style & "    align-items: center;"
                style = style & "    text-align: center;"
                style = style & "    margin-top: 20px;"
                style = style & "}"
                style = style & ".date-container {"
                style = style & "    display: flex;"
                style = style & "    justify-content: center;"
                style = style & "    align-items: center;"
                style = style & "    margin-bottom: 20px;"
                style = style & "}"
                style = style & ".date-container div { margin: 0 10px; }"
                style = style & "button { color: white; font-weight: 600; padding: 10px 15px; background-color: #094ebd; border: none; border-radius: 5px; cursor: pointer; }"
                style = style & "button:hover { background-color: #87CEFA; }"
                style = style & "#from, #to { border-color: #87CEFA; border-radius: 5px; padding: 8px 10px; }"
                style = style & "th.drug-names { width: 600px; }" ' Adjust the width here
                style = style & "</style>"
                
                Response.Write style
                
                
                
                
                Response.Write "<tr>"
                Response.Write "<th> Serial No.</th>"
                Response.Write "<th background-color='#043cd6'>DrugStore Name</th>"
                Response.Write "<th class='drug-names' > Drugstore ID </th>"
                Response.Write "<th> Drug Stock </th>"
                Response.Write "<th> Stock Date </th>"
                Response.Write "</tr>"
                
                
                
            Do While Not .EOF
            
                cnt = cnt + 1
                Response.Write "<tr>"
                Response.Write "<td>" & cnt & "</td>"
                Response.Write "<td>" & .fields("DrugStoreID") & "</td>"
                Response.Write "<td>" & .fields("DrugStoreName") & "</td>"
                Response.Write "<td>" & .fields("Drug Stock") & "</td>"
                Response.Write "<td>" & .fields("Stock Date") & "</td>"
                Response.Write "</tr>"
                .MoveNext
            Loop
            
                Response.Write "</table>"
        End If
        .Close
    End With
    
End Sub