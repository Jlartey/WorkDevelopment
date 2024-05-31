const http = require('http');
const fs = require('fs');

// Create a server
const server = http.createServer((req, res) => {
  // Set the content type to HTML
  res.writeHead(200, { 'Content-Type': 'text/html' });

  // Read the contents of your CSS file

  // Write the HTML code to the res
  res.write('<!DOCTYPE html>');
  res.write('<html lang="en">');
  res.write('<head>');
  res.write('<meta charset="UTF-8">');
  res.write(
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
  );
  res.write('<title>Your Title Here</title>');
  res.write('<style>');
  res.write(`
  * {
  box-sizing: border-box;
}

.container {
  width: 950px;
  margin: auto;
  font-family: 'Trebuchet MS', 'Lucida Sans Unicode', 'Lucida Grande',
    'Lucida Sans', Arial, sans-serif;
}

.heading {
  text-align: center;
  margin-bottom: 0px;
  font-size: 1.4rem;
}

@media (max-width: 360px) {
  .heading {
    margin-bottom: 0px;
    font-size: 0.625rem;
  }
}

#form-info {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 0.625rem;
}

#claim-form {
  margin-bottom: -15px;
}

/* I am done with this section */
#claims-info input {
  width: 80px;
}

.client-information {
  display: flex;
  align-items: center;
  justify-content: space-between;
  /* border: 1px solid red; */
}

.surname-others-dob-hosp input {
  width: 120px;
}

.age-gender-member-claims {
  display: flex;
  flex-direction: column;
  text-align: right;
}

#services-provided {
  display: flex;
  margin-bottom: 10px;
  align-items: space-between;
  margin-bottom: 30px;
  gap: 20px;
}

#service-type {
  border: 1px solid black;
  width: 60%;
  margin-right: auto;
  padding: 5px 5px;
}

#provision-date {
  border: 1px solid black;
  width: 35%;
  padding: 15px 15px;
}

#attendance-type {
  border: 1px solid black;
  padding: 5px 5px;
  margin-bottom: 30px;
}

#physician-details {
  margin-top: 10px;
  border: 1px solid black;
  width: 100%;
  padding-top: 15px;
  padding-left: 15px;
  margin-bottom: 30px;
}

#physician-details input {
  border: none;
} `);
  res.write('</style>');
  res.write('</head>');
  res.write(
    "<body style='margin-top: 50px; margin-bottom: 50px' class='container'>"
  );
  res.write("    <h2 class='heading'>NATIONAL HEALTH INSURANCE SCHEME</h2> ");
  res.write("    <section id='form-info'> ");
  res.write("      <div class='form-and-regulation' style='display: flex'> ");
  res.write('        <img ');
  res.write("          style='margin-right: 0.625rem' ");
  res.write(
    "          src='https://c8.alamy.com/comp/KYBMDA/national-health-insurance-scheme-ghana-nhis-logo-KYBMDA.jpg' "
  );
  res.write("          alt='National Health Insurance Logo' ");
  res.write("          height='70' ");
  res.write("         width='70' ");
  res.write('        /> ');
  res.write("        <div style='display: flex; flex-direction: column'> ");
  res.write("          <p id='claim-form'>Claim Form</p> ");
  res.write('          <p>(Regulation 62)</p> ');
  res.write('        </div> ');
  res.write('      </div> ');

  res.write("      <div class='form-no-and-hi-code' style='display: block'> ");
  res.write("        <label for='form-no'>Form No.:</label> ");
  res.write("        <input type='text' id='form-no' name='form-no' /><br /> ");

  res.write(
    "        <label for='hi-code' style='margin-right: 0.6785rem'>HI Code:</label> "
  );
  res.write("        <input type='number' id='hi-code' name='hi-code' /> ");
  res.write('      </div> ');
  res.write('    </section> ');

  res.write('    <section ');
  res.write("      id='claims-info' ");
  res.write(
    "      style='display: flex; justify-content: space-between; align-items: center' "
  );
  res.write('    > ');
  res.write("      <div style='display: flex; flex-direction: column'> ");
  res.write("        <div style='display: block'> ");
  res.write('          <label ');
  res.write("            for='claim-code' ");
  res.write("            class='claims-info-label' ");
  res.write("            style='margin-right: 14px' ");
  res.write('            >Claim Code:</label ');
  res.write('          > ');
  res.write(
    "          <input type='text' id='claim-code' name='claim-code' /> "
  );
  res.write('        </div> ');

  res.write("        <div style='display: block'> ");
  res.write("          <label for='scheme-code' class='claims-info-label' ");
  res.write('            >Scheme Code:</label ');
  res.write('          > ');
  res.write(
    "          <input type='text' id='scheme-code' name='scheme-code' /> "
  );
  res.write('        </div> ');

  res.write("        <div style='display: block'> ");
  res.write('          <label ');
  res.write("            for='referral-no' ");
  res.write("            class='claims-info-label' ");
  res.write("            style='margin-right: 14px' ");
  res.write('            >Referral No:</label ');
  res.write('          > ');
  res.write(
    "          <input type='text' id='referral-no' name='referral-no' /> "
  );
  res.write('        </div> ');
  res.write('      </div> ');

  res.write('      <div> ');
  res.write("        <label for='month-claim'>Month of Claim:</label> ");
  res.write('        <input ');
  res.write("          type='text' ");
  res.write("          id='month-claim' ");
  res.write("          name='month-claim' ");
  res.write("          style='margin-right: 5px' ");
  res.write('        /> ');
  res.write('      </div> ');
  res.write('      <div> ');
  res.write("        <label for='date-claim'>Date of Claim: </label> ");
  res.write("        <input type='text' id='date-claim' name='date-claim' /> ");
  res.write('      </div> ');
  res.write('    </section> ');

  res.write('    <hr /> ');

  res.write('    <h4>Client Information</h4> ');
  res.write("    <div class='client-information'> ");
  res.write('      <div ');
  res.write("        class='surname-others-dob-hosp' ");
  res.write("        style='display: flex; flex-direction: column' ");
  res.write('      > ');
  res.write("        <div style='display: flex'> ");
  res.write(
    "          <label for='surname' style='margin-right: 77.5px'>Surname:</label> "
  );
  res.write("          <input type='text' id='surname' name='surname' /> ");
  res.write('        </div> ');

  res.write("        <div style='display: flex'> ");
  res.write(
    "          <label for='other-names' style='margin-right: 42px'>Other Names:</label> "
  );
  res.write(
    "          <input type='text' id='other-names' name='other-names' /> "
  );
  res.write('        </div> ');

  res.write("        <div style='display: flex'> ");
  res.write(
    "          <label for='dob' style='margin-right: 43px'>Date of Birth:</label> "
  );
  res.write("          <input type='date' id='dob' name='dob' /> ");
  res.write('        </div> ');

  res.write("        <div style='display: flex'> ");
  res.write(
    "          <label for='hospital-record' style='margin-right: 5px'>Hospital Record No:</label> "
  );
  res.write(
    "          <input type='text' id='hospital-record' name='hospital-record' /> "
  );
  res.write('        </div> ');
  res.write('      </div> ');

  res.write('      <div> ');
  res.write("        <label for='age'>Age:</label> ");
  res.write("        <input type='number' id='age' name='age' /> ");
  res.write('      </div> ');

  res.write("      <div class='age-gender-member-claims'> ");
  res.write('        <div> ');
  res.write("          <label for='gender'>Gender</label> ");
  res.write("          <input type='text' id='gender' name='gender' /> ");
  res.write('        </div> ');

  res.write('        <div> ');
  res.write("          <label for='member-no'>Member No.:</label> ");
  res.write("          <input type='text' id='member-no' name='member-no' /> ");
  res.write('        </div> ');

  res.write("        <div style='display: flex'> ");
  res.write("          <label for='claims-code'>Claims Check Code:</label> ");
  res.write(
    "          <input type='text' id='claims-code' name='claims-code' /> "
  );
  res.write('        </div> ');
  res.write('      </div> ');
  res.write('    </div> ');

  res.write('    <hr /> ');

  res.write('    <p> ');
  res.write('      <strong>Services Provided</strong> ');
  res.write('      <em>(to be filled by health care providers)</em> ');
  res.write('    </p> ');

  res.write("    <section id='services-provided'> ");
  res.write("      <div id='service-type'> ");
  res.write('        <p><em>Type of Services</em></p> ');
  res.write('        (a) ');
  res.write(
    "        <input type='checkbox' id='outpatient' name='outpatient' /> "
  );
  res.write("        <label for='outpatient'>Outpatient</label> ");
  res.write(
    "        <input type='checkbox' id='inpatient' name='inpatient' /> "
  );
  res.write("        <label for='inpatient'>Inpatient</label> ");
  res.write("        <input type='checkbox' id='pharmacy' name='pharmacy' /> ");
  res.write("        <label for='pharmacy'>Pharmacy</label> ");
  res.write(
    "        <p style='margin-left: 100px; margin-bottom: -2px'>Investigation</p> "
  );
  res.write('        <hr /> ');
  res.write('        (b) ');
  res.write('        <input ');
  res.write("          type='checkbox' ");
  res.write("          id='all-inclusive' ");
  res.write("          name='all-inclusive' ");
  res.write("          value='1' ");
  res.write('        /> ');
  res.write("        <label for='all-inclusive'>All Inclusive</label> ");
  res.write(
    "        <input type='checkbox' id='unbundled' name='unbundled' /> "
  );
  res.write("        <label for='unbundled'>Unbundled</label> ");
  res.write('        <hr /> ');
  res.write('        <p><em>Outcome</em></p> ');
  res.write(
    "        <input type='checkbox' id='discharged' name='discharged' /> "
  );
  res.write("        <label for='discharged'>Discharged</label> ");
  res.write(
    "        <input type='checkbox' id='died' name='died' value='1' /> "
  );
  res.write("        <label for='died'>Died</label> ");
  res.write("        <input type='checkbox' id='transfer' name='transfer' /> ");
  res.write("        <label for='transfer'>Transfer</label> ");
  res.write('        <br /> ');
  res.write(
    "        <input type='checkbox' id='absconded' name='absconded' /> "
  );
  res.write(
    "        <label for='absconded'>Absconded/Discharged against medical advice</label> "
  );
  res.write('      </div> ');
  res.write("      <div id='provision-date'> ");
  res.write('        <h4>Date(s) of Services Provision</h4> ');
  res.write(
    "        <label for='first-visit' style='margin-right: 40px'>1st Visit/Admission</label> "
  );
  res.write(
    "        <input type='date' id='first-visit' name='first-visit' /><br /> "
  );
  res.write(
    "        <label for='second-visit' style='margin-right: 35px'>2nd Visit/Admission</label> "
  );
  res.write(
    "        <input type='date' id='second-visit' name='second-visit' /><br /> "
  );
  res.write(
    "        <label for='third-visit' style='margin-right: 117px'>3rd Visit</label> "
  );
  res.write(
    "        <input type='date' id='third-visit' name='third-visit' /><br /> "
  );
  res.write(
    "        <label for='fourth-visit' style='margin-right: 38px'>4th Visit/Admission</label> "
  );
  res.write(
    "        <input type='date' id='fourth-visit' name='fourth-visit' /><br /> "
  );
  res.write(
    "        <label for='duration-spell'>Duration of Spell(days)</label> "
  );
  res.write('        <input ');
  res.write("          type='number' ");
  res.write("          id='duration-spell' ");
  res.write("          name='duration-spell' ");
  res.write("          style='width: 50px; margin-left: 77px' ");
  res.write('        /><br /> ');
  res.write('      </div> ');
  res.write('    </section> ');
  res.write("    <section id='attendance-type'> ");
  res.write('      <p><em>Type of Attendance</em></p> ');
  res.write("      <input type='checkbox' id='follow-up' name='follow-up' /> ");
  res.write("      <label for='follow-up'>Chronic Follow-up</label> ");
  res.write("      <input type='checkbox' id='emergency' name='emergency' /> ");
  res.write("      <label for='emergency'>Emergency</label> ");
  res.write(
    "      <input type='checkbox' id='acute-episode' name='acute-episode' /> "
  );
  res.write("      <label for='acute-episode'>Acute Episode</label> ");
  res.write("      <input type='checkbox' id='anc' name='anc' /> ");
  res.write("      <label for='anc'>ANC</label> ");
  res.write('    </section> ');
  res.write("    <section id='physician-details'> ");
  res.write(
    "      <label for='physician-name'>Physician/Clinician Name: </label> "
  );
  res.write('      <input ');
  res.write("        type='text' ");
  res.write("        id='physician-name' ");
  res.write("        name='physician-name' ");
  res.write('      /><br /><br /> ');
  res.write(
    "      <label for='physician-Id' style='margin-right: 26px'>Physician/Clinician ID: </label> "
  );
  res.write(
    "      <input type='text' id='physician-Id' name='physician-Id' /><br /><br /> "
  );
  res.write(
    "      <label for='speciality-code' style='margin-right: 65px'>Specialty Code: </label> "
  );
  res.write('      <input ');
  res.write("        type='text' ");
  res.write("        id='speciality-code' ");
  res.write("        name='speciality-code' ");
  res.write('      /><br /><br /> ');
  res.write('    </section> ');
  res.write(
    "    <table cellpadding='2' border='1' width='100%' cellspacing='0'> "
  );
  res.write('      <thead> ');
  res.write('        <tr> ');
  res.write("          <td colspan='4'> ");
  res.write("            <strong style='padding: 0px 2px'>Diagnosis</strong ");
  res.write("            ><em style='font-size: 12px' ");
  res.write(
    '              >(to be filled by health care providers who have provided out and '
  );
  res.write('              in-patient services)</em ');
  res.write('            > ');
  res.write('          </td> ');
  res.write('        </tr> ');
  res.write('      </thead> ');
  res.write('      <thead> ');
  res.write('        <tr> ');
  res.write('          <th></th> ');
  res.write('          <th>Description</th> ');
  res.write('          <th>ICD-10</th> ');
  res.write('          <th>G-DRG</th> ');
  res.write('        </tr> ');
  res.write('      </thead> ');
  res.write('      <tbody> ');
  res.write('        <tr> ');
  res.write('          <td>1</td> ');
  res.write('          <td>MALARIA</td> ');
  res.write('          <td>B54</td> ');
  res.write('          <td>OPDC06C</td> ');
  res.write('        </tr> ');
  res.write('        <tr> ');
  res.write('          <td>2</td> ');
  res.write('          <td>ANAEMIA (UNSPECIFIED)</td> ');
  res.write('          <td>D64.9</td> ');
  res.write('          <td>OPDC06C</td> ');
  res.write('        </tr> ');
  res.write('      </tbody> ');
  res.write('    </table> ');
  res.write('  </body> ');
  res.write('</html>');

  // End the res
  res.end();
});

// Set the port number
const port = 3000;

// Start the server
server.listen(port, () => {
  console.log(`Server running at http://localhost:${port}/`);
});
