 ' Set the response content type to JSON
    response.ContentType = "application/json"
    response.write jsonData

    ' Plotly chart creation using the JSON data
    ' Assuming the Plotly library is available in the environment
    ' You would need to create the appropriate JavaScript code to use Plotly for rendering the chart
    Dim script
    'script = "<script src=""https://cdn.plot.ly/plotly-latest.min.js""></script>"
    script = script & "<script>"
    script = script & "var data = " & jsonData & ";"
    script = script & "var traces = [];"
    script = script & "var quarters = [];"

    ' Prepare data for Plotly
    ' Add separate traces for each quarter
    script = script & "data.data.forEach(function(record) {"
    script = script & "if (!quarters.includes(record.Quarter)) {"
    script = script & "quarters.push(record.Quarter);"
    script = script & "}"
    script = script & "});"

    ' Create traces for each quarter
    script = script & "quarters.forEach(function(quarter) {"
    script = script & "var trace = {"
    script = script & "x: [],"
    script = script & "y: [],"
    script = script & "name: quarter,"
    script = script & "type: 'bar'"
    script = script & "};"
    script = script & "data.data.forEach(function(record) {"
    script = script & "if (record.Quarter === quarter) {"
    script = script & "trace.x.push(record.Year);"
    script = script & "trace.y.push(record.QuarterTotal);"
    script = script & "}"
    script = script & "});"
    script = script & "traces.push(trace);"
    script = script & "});"

    ' Plot the data
    script = script & "Plotly.newPlot('chart', traces, {"
    script = script & "title: 'Quarterly Insurance Type Analysis',"
    script = script & "xaxis: {title: 'Year'},"
    script = script & "yaxis: {title: 'Quarter Total'}"
    script = script & "});"
    script = script & "</script>"

    ' Add the script to the response
    response.write script

    ' Generate HTML table with DataTables integration
    Dim htmlTable
    'htmlTable = "<link rel='stylesheet' type='text/css' href='https://cdn.datatables.net/1.10.24/css/jquery.dataTables.min.css'>"
    htmlTable = htmlTable & "<script type='text/javascript' src='https://code.jquery.com/jquery-3.6.0.min.js'></script>"
    htmlTable = htmlTable & "<script type='text/javascript' src='https://cdn.datatables.net/1.10.24/js/jquery.dataTables.min.js'></script>"
    htmlTable = htmlTable & "<script type='text/javascript'>$(document).ready(function() {$('#insuranceTable').DataTable();});</script>"
    htmlTable = htmlTable & "<table id='quarterlyTable' class='display' width='100%'><thead><tr><th>InsuranceTypeName</th><th>Year</th><th>Quarter</th><th>QuarterTotal</th><th>PrevQuarterTote</th><th>DiffTote</th><th>QoQChange</th><th>QuarterPercentContToTote</th><th>YearPercentContToTote</th><th>OverallTotal</th></tr></thead><tbody>"
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            htmlTable = htmlTable & "<tr>"
            htmlTable = htmlTable & "<td>" & rst.Fields("InsuranceTypeName").Value & "</td>"
            htmlTable = htmlTable & "<td>" & rst.Fields("Year").Value & "</td>"
            htmlTable = htmlTable & "<td>" & rst.Fields("Quarter").Value & "</td>"
            htmlTable = htmlTable & "<td>" & rst.Fields("QuarterTotalF").Value & "</td>"
            htmlTable = htmlTable & "<td>" & rst.Fields("PrevQuarterTote").Value & "</td>"
            htmlTable = htmlTable & "<td>" & rst.Fields("DiffTote").Value & "</td>"
            htmlTable = htmlTable & "<td>" & rst.Fields("QoQChange").Value & "</td>"
            htmlTable = htmlTable & "<td>" & rst.Fields("QuarterPercentContToTote").Value & "</td>"
            htmlTable = htmlTable & "<td>" & rst.Fields("YearPercentContToTote").Value & "</td>"
            htmlTable = htmlTable & "<td>" & rst.Fields("OverallTotal").Value & "</td>"
            htmlTable = htmlTable & "</tr>"
            rst.MoveNext
        Loop
    End If
    
    htmlTable = htmlTable & "</tbody></table>"

    ' Add the HTML table to the response
    response.write htmlTable