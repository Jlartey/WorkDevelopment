'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
response.Clear
response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"
response.write "  <head>"
response.write "    <meta charset=""UTF-8"" />"
response.write "    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />"
response.write "    <link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/choices.js@11.1.0/public/assets/styles/choices.min.css"" />"
response.write "    <script src=""https://cdn.jsdelivr.net/npm/choices.js@11.1.0/public/assets/scripts/choices.min.js""></script>"
response.write "    <title>Sponsor Form</title>"
response.write "    <style>"
response.write "      .main {"
response.write "        display: flex;"
response.write "        justify-content: center;"
response.write "        align-items: center;"
response.write "        min-height: 100vh;"
response.write "        background-color: #f4f5f7;"
response.write "        margin: 0;"
response.write "        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;"
response.write "      }"
response.write "      .form-container {"
response.write "        background-color: #fff;"
response.write "        padding: 2rem;"
response.write "        border-radius: 0.5rem;"
response.write "        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);"
response.write "        max-width: 350px;"
response.write "        width: 100%;"
response.write "      }"
response.write "      form {"
response.write "        display: flex;"
response.write "        flex-direction: column;"
response.write "        gap: 1rem;"
response.write "      }"
response.write "      label {"
response.write "        font-size: 0.9rem;"
response.write "        font-weight: 500;"
response.write "        color: #374151;"
response.write "      }"
response.write "      .choices {"
response.write "        max-width: 100%;"
response.write "      }"
response.write "      .choices__inner {"
response.write "        padding: 0.5rem;"
response.write "        border: 1px solid #d1d5db;"
response.write "        border-radius: 0.375rem;"
response.write "        background-color: #fff;"
response.write "        font-size: 0.875rem;"
response.write "        line-height: 1.5rem;"
response.write "        min-height: 2.25rem;"
response.write "        width: 100%;"
response.write "        box-sizing: border-box;"
response.write "        transition: border-color 0.2s ease;"
response.write "      }"
response.write "      .choices__inner:hover, .choices__inner:focus, .choices__inner:focus-within {"
response.write "        border-color: #92e681;"
response.write "        width: 100% !important;"
response.write "      }"
response.write "      .choices__input {"
response.write "        background-color: #fff;"
response.write "        border: none;"
response.write "        outline: none;"
response.write "        font-size: 0.875rem;"
response.write "        line-height: 1.5rem;"
response.write "        padding: 0.25rem 0.5rem;"
response.write "        width: 100%;"
response.write "        box-sizing: border-box;"
response.write "      }"
response.write "      .choices__list--dropdown {"
response.write "        position: absolute;"
response.write "        z-index: 10;"
response.write "        border: 1px solid #d1d5db;"
response.write "        border-radius: 0.375rem;"
response.write "        background-color: #fff;"
response.write "        max-height: 15rem;"
response.write "        overflow: auto;"
response.write "        width: 100%;"
response.write "        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);"
response.write "      }"
response.write "      .choices__item--selectable {"
response.write "        padding: 0.5rem 0.75rem;"
response.write "        cursor: pointer;"
response.write "      }"
response.write "      .choices__item--selectable:hover {"
response.write "        background-color: #eaffe6;"
response.write "      }"
response.write "      .choices__placeholder {"
response.write "        opacity: 0.5;"
response.write "      }"
response.write "      input[type='number'] {"
response.write "        padding: 0.5rem;"
response.write "        border: 1px solid #d1d5db;"
response.write "        border-radius: 0.375rem;"
response.write "        font-size: 0.875rem;"
response.write "        line-height: 1.5rem;"
response.write "        width: 100%;"
response.write "        box-sizing: border-box;"
response.write "        transition: border-color 0.2s ease;"
response.write "      }"
response.write "      input[type='number']:focus {"
response.write "        outline: none;"
response.write "        border-color: #92e681;"
response.write "        box-shadow: 0 0 0 3px rgba(146, 230, 129, 0.2);"
response.write "      }"
response.write "      button {"
response.write "        padding: 0.5rem 1.5rem;"
response.write "        background-color: #327b2c;"
response.write "        color: #fff;"
response.write "        border: none;"
response.write "        border-radius: 0.375rem;"
response.write "        font-size: 0.875rem;"
response.write "        font-weight: 500;"
response.write "        cursor: pointer;"
response.write "        transition: background-color 0.2s ease, transform 0.1s ease;"
response.write "      }"
response.write "      button:hover {"
response.write "        background-color: #2a6724;"
response.write "        transform: translateY(-1px);"
response.write "      }"
response.write "      button:active {"
response.write "        transform: translateY(0);"
response.write "      }"
response.write "    </style>"
response.write "  </head>"
response.write "  <div class='main'>"


 sql = "select SponsorID, SponsorName from Sponsor"
    Set rstDropdown = CreateObject("ADODB.Recordset")
    rstDropdown.open qryPro.FltQry(sql), conn, 3, 4

    dropdownOptions = ""

    With rstDropdown
        If .recordCount > 0 Then
            .movefirst
            Do Until .EOF
                optionHTML = "<option value='" & .fields("SponsorID") & "'>" & .fields("SponsorName") & "</option>"
                dropdownOptions = dropdownOptions & optionHTML
                .MoveNext
            Loop
        End If
    End With

    rstDropdown.Close
    Set rstDropdown = Nothing
    
    
    sql2 = "select BillMonthID, BillMonthName from BillMonth ORDER BY BillMonthID DESC"
    Set rstDropdown2 = CreateObject("ADODB.Recordset")
    rstDropdown2.open qryPro.FltQry(sql2), conn, 3, 4


    dropdownOptions2 = ""

    With rstDropdown2
        If .recordCount > 0 Then
            .movefirst
            Do Until .EOF
                optionHTML2 = "<option value='" & .fields("BillMonthID") & "'>" & .fields("BillMonthName") & "</option>"
                dropdownOptions2 = dropdownOptions2 & optionHTML2
                .MoveNext
            Loop
        End If
    End With

    rstDropdown2.Close
    Set rstDropdown2 = Nothing


response.write "    <div class=""form-container"">"
response.write "      <form>"
response.write "        <label for=""sponsors"">Sponsor:</label>"
response.write "        <select name=""sponsors"" id=""sponsors"">"
response.write dropdownOptions
response.write "        </select>"
response.write "        <label for=""billMonth"">Bill Month:</label>"
response.write "        <select name=""billMonth"" id=""billMonth"">"
response.write dropdownOptions2
response.write "        </select>"
response.write "        <label for=""amount"">Amount:</label>"
response.write "        <input type=""number"" id=""amount"" name=""amount"" />"
response.write "        <button type=""button"" onclick=""sendData(event)"">Save</button>"

response.write vbCrLf & " <script>"
response.write vbCrLf & " function sendData(event) {"
response.write vbCrLf & "          const sponsorSelect = document.getElementById('sponsors');"
response.write vbCrLf & "          const selectedSponsor = sponsorSelect.selectedOptions[0] ? {"
response.write vbCrLf & "            id: sponsorSelect.value,"
response.write vbCrLf & "            value: sponsorSelect.selectedOptions[0].text"
response.write vbCrLf & "          } : null;"
response.write vbCrLf & "          const billMonthSelect = document.getElementById('billMonth');"
response.write vbCrLf & "          const selectedBillMonth = billMonthSelect.selectedOptions[0] ? {"
response.write vbCrLf & "            id: billMonthSelect.value,"
response.write vbCrLf & "            value: billMonthSelect.selectedOptions[0].text"
response.write vbCrLf & "          } : null;"
response.write vbCrLf & "          const amount = document.getElementById('amount').value;"
response.write vbCrLf & "          if (!selectedSponsor || !selectedBillMonth || !amount) {"
response.write vbCrLf & "            alert('Please select a sponsor, a bill month, and enter an amount.');"
response.write vbCrLf & "            return;"
response.write vbCrLf & "          }"
response.write vbCrLf & "        console.log(selectedSponsor, selectedBillMonth, amount)"

response.write vbCrLf & "       const xhr = new XMLHttpRequest();"
response.write vbCrLf & "       let url = ""wpgxmlhttp.asp?procedurename=InsertSponsorPayment"";"
response.write vbCrLf & "          url += '&sponsorID=' + selectedSponsor.id;"
response.write vbCrLf & "          url += '&billMonthID=' + selectedBillMonth.id;"
response.write vbCrLf & "          url += '&amount=' + amount;"

response.write vbCrLf & "        console.log(url)"

response.write vbCrLf & "          fetch(url)"
response.write vbCrLf & "            .then(response => {"
response.write vbCrLf & "              if (!response.ok) {"
response.write vbCrLf & "                throw new Error('Network response was not ok');"
response.write vbCrLf & "              }"
response.write vbCrLf & "              return response.json();"
response.write vbCrLf & "            })"
response.write vbCrLf & "            .then(data => {"
response.write vbCrLf & "              console.log('Server response:', data);"
response.write vbCrLf & "              if (data.success) {"
response.write vbCrLf & "                alert('Data saved successfully!');"
response.write vbCrLf & "                window.parent.postMessage({action: 'closeAndReload'}, '*');  
response.write vbCrLf & "              } else {"
response.write vbCrLf & "                alert('Save failed: ' + (data.message || 'Unknown error'));"
response.write vbCrLf & "              }"
response.write vbCrLf & "            })"

response.write vbCrLf & "}"
response.write vbCrLf & " </script>"

response.write vbCrLf & "      </form>"
response.write vbCrLf & "    </div>"


response.write vbCrLf & "    <script>"
response.write vbCrLf & "      try {"
response.write vbCrLf & "        const sponsorChoices = new Choices('#sponsors', {"
response.write vbCrLf & "          searchEnabled: true,"
response.write vbCrLf & "          placeholderValue: 'Search',"
response.write vbCrLf & "          searchPlaceholderValue: 'Search',"
response.write vbCrLf & "          itemSelectText: '',"
response.write vbCrLf & "          shouldSort: false,"
response.write vbCrLf & "        });"
response.write vbCrLf & "        const billMonthChoices = new Choices('#billMonth', {"
response.write vbCrLf & "          searchEnabled: true,"
response.write vbCrLf & "          placeholderValue: 'Search',"
response.write vbCrLf & "          searchPlaceholderValue: 'Search',"
response.write vbCrLf & "          itemSelectText: '',"
response.write vbCrLf & "          shouldSort: false,"
response.write vbCrLf & "        });"
response.write vbCrLf & "      } catch (error) {"
response.write vbCrLf & "        console.error('Choices.js initialization failed:', error);"
response.write vbCrLf & "      }"
response.write vbCrLf & "     </script>"


response.write vbCrLf & "  </div>"
response.write vbCrLf & "</html>"


'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
