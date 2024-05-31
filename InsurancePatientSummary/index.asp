CssStyles
InsuranceSummary

Sub InsuranceSummary
  response.write"<body>"
    response.write"  <h1>Corporate/Insurance Patient Summary For the Month Of:</h1> "
    response.write"  <table cellpadding='2' border='1' width='100%' cellspacing='0'>"
    response.write"    <thead>"
    response.write"      <tr>"
    response.write"        <th>No.</th>"
    response.write"        <th>Sponsor Names</th>"
    response.write"        <th>Number Of Sponsors</th>"
    response.write"        <th>Total Amount</th>"
    response.write"        <th>View Bill</th>"
    response.write"      </tr>"
    response.write"    </thead>"
    response.write"    <tbody>"
    response.write"      <tr>"
    response.write"        <td>1</td>"
    response.write"        <td>Spagad</td>"
    response.write"        <td>20</td>"
    response.write"        <td>GHc 3000.00</td>"
    response.write"        <td>"
    response.write"          <a href='detail.html' target='_blank'><button>View Details</button></a>"
    response.write"        </td>"
    response.write"      </tr>"
    response.write"      <tr>"
    response.write"        <td>2</td>"
    response.write"        <td>ECG</td>"
    response.write"        <td>50</td>"
    response.write"        <td>GHc 5000.00</td>"
    response.write"        <td>"
    response.write"          <a href='detail.html' target='_blank'><button>View Details</button></a">"
    response.write"        </td>"
    response.write"      </tr>"
    response.write"      <tr>"
    response.write"        <td>3</td>"
    response.write"        <td>CBI</td>"
    response.write"        <td>70</td>"
    response.write"        <td>GHc 7500.00</td>"
    response.write"        <td>"
    response.write"          <a href='detail.html' target='_blank'><button>View Details</button>>"
    response.write"        </td>"
    response.write"      </tr>"
    response.write"    </tbody>"
    response.write"  </table>"
  response.write"</body>"

End Sub

Sub CssStyles
  response.write"<style>"
  response.write"  body {"
  response.write"    width: 80%;"
  response.write"    padding: 50px 100px;"
  response.write"    margin: auto;"
  response.write"    font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande','Lucida Sans', Arial, sans-serif;"
  response.write"  }"

  response.write"  h1 {"
  response.write"    text-align: center;"
  response.write"  }"
  response.write"</style>"
End Sub
