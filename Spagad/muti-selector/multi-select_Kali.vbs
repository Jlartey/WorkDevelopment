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
                    Response.Write "    <div id=""countries-wrapper"">" & vbCrLf
                    Response.Write "        <select name=""countries"" id=""countries"" multiple>" & vbCrLf
                    Do Until .EOF
                    Response.Write " <option value=""" & .fields("DrugStoreName") & """>" & .fields("DrugStoreName") & "</option>" & vbCrLf
                    .MoveNext
                    Loop
                    Response.Write "        </select>" & vbCrLf
                    Response.Write "    </div>" & vbCrLf
                    Response.Write "    <script>" & vbCrLf
                    Response.Write "        new MultiSelectTag('countries', {" & vbCrLf
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