'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
'Response.Write "Hello Joe"

Dim currentTime
currentTime = Now

Response.Write "<!DOCTYPE html>"
Response.Write "<html lang=""en"">"
Response.Write "  <head>"
Response.Write "    <meta charset=""UTF-8"" />"
Response.Write "    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />"
Response.Write "    <title>Treatment Flow Sheet</title>"
Response.Write "  </head>"
Response.Write "  <body>"
Response.Write "    <style>"
Response.Write "      body {"
Response.Write "        font-family: Arial, sans-serif;"
Response.Write "      }"
Response.Write "      .container {"
Response.Write "        padding: 20px 20px;"
Response.Write "        margin:  auto;"
Response.Write "        width: 600px;"
Response.Write "        line-height: 1.6;"
Response.Write "      }"
Response.Write "      .row {"
Response.Write "        display: flex;"
Response.Write "        margin-bottom: 10px;"
Response.Write "      }"
Response.Write "      label {"
Response.Write "        width: 200px;"
Response.Write "        margin-right: 20px;"
Response.Write "        text-align: right;"
Response.Write "      }"
Response.Write "      .value {"
Response.Write "        flex: 1;"
Response.Write "      }"
Response.Write "      button {"
Response.Write "        width: 75px;"
Response.Write "        margin: auto;"
Response.Write "        margin-left: 160px;"
Response.Write "        border-radius: 10px;"
Response.Write "        background-color: blue;"
Response.Write "        color: white;"
Response.Write "        padding: 10px 10px;"
Response.Write "        outline: none;"
Response.Write "        cursor: pointer;"
Response.Write "        font-family: Arial, sans-serif;"
Response.Write "      }"
Response.Write "    </style>"
Response.Write "  </head>"
Response.Write "  <body>"
Response.Write "    <div class=""container"">"
Response.Write "      <div class=""row"">"
Response.Write "        <label for=""treatment-date"">Treatment Date :</label>"
Response.Write "        " & currentTime & " "
Response.Write "      </div>"
Response.Write "      <div class=""row"">"
Response.Write "        <label for=""treatment-type"">Type :</label>"
Response.Write "        <select name=""treatment-type"" id=""treatment-type"">"
Response.Write "          <option value=""Therapeutic Exercises"">Therapeutic Exercises</option>"
Response.Write "          <option value=""Manual"">Manual</option>"
Response.Write "          <option value=""Other"">Other</option>"
Response.Write "        </select>"
Response.Write "      </div>"
Response.Write "      <div class=""row"">"
Response.Write "        <label for=""intervention"">Intervention :</label>"
Response.Write "        <input type=""text"" name=""intervention"" id=""intervention"" style=""width: 300px"" />"
Response.Write "      </div>"
Response.Write "      <div class=""row"">"
Response.Write "        <label for=""treatment-value"">Value :</label>"
Response.Write "        <input type=""text"" name=""treatment-value"" id=""treatment-value"" style=""width: 300px""/>"
Response.Write "      </div>"
Response.Write "      <button onclick=""updateTreatment()"">SAVE</button>"
Response.Write "    </div>"

Response.Write "<script>"
Response.Write "        function updateTreatment(){"
Response.Write "            const treamentDate = currentTime;"
Response.Write "            const treatmentType = document.getElementById(""treatment-type"");"
Response.Write "            const intervention = document.getElementById(""intervention"");"
Response.Write "            const treatmentValue = document.getElementById(""treatment-value"");"
Response.Write "            "
Response.Write "            let url = ""wpgXMLHTTP.asp?procedurename=updateTreatment"";"
Response.Write "            if(treatmentType?.value&&intervention?.value&&treatmentValue?.value){"
Response.Write "                url = url + ""&treatment-date="" + treamentDate.value;"
Response.Write "                url = url + ""&treatment-type="" + treatmentType.value;"
Response.Write "                url = url + ""&intervention="" + intervention.value;"
Response.Write "                url = url + ""&treatment-value="" + treatmentValue.value;"
'Response.Write "                url = url + ""&visitid="" + visitID;"
Response.Write "                fetch(url).then(response=>response.json()).then(data=>{"
Response.Write "                    window.location.reload()"
Response.Write "                })"
Response.Write "            }else{"
Response.Write "                alert('please enter all values');"
Response.Write "            }"
Response.Write "        }"
Response.Write "    </script>"
Response.Write "  </body>"
Response.Write "</html>"

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
