'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim labTestID
labTestID = Request.queryString("PrintFilter")

'Response.Write "<script>"
'Response.Write "        function updateTreatment(){"
'Response.Write "            const treamentDate = currentTime;"
'Response.Write "            const treatmentType = document.getElementById(""treatment-type"");"
'Response.Write "            const intervention = document.getElementById(""intervention"");"
'Response.Write "            const treatmentValue = document.getElementById(""treatment-value"");"
'Response.Write "            "
'Response.Write "            let url = ""wpgXMLHTTP.asp?procedurename=updateTreatment"";"
'Response.Write "            if(treatmentType?.value&&intervention?.value&&treatmentValue?.value){"
'Response.Write "                url = url + ""&treatment-date="" + treamentDate.value;"
'Response.Write "                url = url + ""&treatment-type="" + treatmentType.value;"
'Response.Write "                url = url + ""&intervention="" + intervention.value;"
'Response.Write "                url = url + ""&treatment-value="" + treatmentValue.value;"
''Response.Write "                url = url + ""&visitid="" + visitID;"
'Response.Write "                fetch(url).then(response=>response.json()).then(data=>{"
'Response.Write "                    window.location.reload()"
'Response.Write "                })"
'Response.Write "            }else{"
'Response.Write "                alert('please enter all values');"
'Response.Write "            }"
'Response.Write "        }"
'Response.Write "    </script>"

response.write "<script> "
response.write "   document.getElementById('callApiButton').addEventListener('click', function() {"
response.write "       var xhr = new XMLHttpRequest();"
response.write "      var url = 'wpgxmlhttp.asp?procedurename=UpdateLabByDoctorStatus&tablename=workingday&LabTestID='" & labTestID & "';"
response.write "       xhr.open('GET', url, true);"
response.write "       xhr.onreadystatechange = function() {"
response.write "           if (xhr.readyState === 4 && xhr.status === 200) {"
response.write "               console.log(xhr.responseText);"
response.write "           }"
response.write "        };"
response.write "       xhr.send();"
response.write "   });"
response.write " </script>"

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
