let outcome;
let intervention;

function updateOutcome() {
  const respSympYes = document.querySelector(
    'input[name="EMR050^^EMR05030Column21"]:checked'
  );
  const expHistory1Yes = document.querySelector(
    'input[name="EMR050^^EMR05024Column2"]:checked'
  );
  const expHistory2Yes = document.querySelector(
    'input[name="EMR050^^EMR05024Column5"]:checked'
  );

  if (
    (expHistory1Yes && expHistory1Yes.value === 'PP0102') ||
    (expHistory2Yes && expHistory2Yes.value === 'PP0102')
  ) {
    outcome = 'High Risk';
    intervention =
      'Immediate masking, isolation, notify infection control, initiate referral';
  } else if (respSympYes && respSympYes.value === 'RESP01') {
    outcome = 'Moderate Risk';
    intervention =
      'Mask, fast-track to consultation, advise home isolation if needed';
  } else {
    outcome = 'No Risk';
    intervention = 'Proceed with standard OPD visit';
  }

  console.log(outcome);

  document.getElementById('EMR050^^EMR05027Column2').value = outcome;
  document.getElementById('EMR050^^EMR05027Column5').value = intervention;
}

document
  .querySelectorAll(
    'input[name="EMR050^^EMR05030Column21"], input[name="EMR050^^EMR05024Column2"], input[name="EMR050^^EMR05024Column5"]'
  )
  .forEach((radio) => {
    radio.addEventListener('change', updateOutcome);
  });
updateOutcome();
