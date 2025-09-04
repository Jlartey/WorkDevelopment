TDProp**3**Left**Top%%Default**<script>

let outcome;

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
  } else if (respSympYes && respSympYes.value === 'RESP01') {
    outcome = 'Moderate Risk';
  } else {
    outcome = 'No Risk';
  }

  console.log(outcome);

  document.getElementById('EMR050^^EMR05027Column2').value = outcome; 
}


document
  .querySelectorAll(
    'input[name="EMR050^^EMR05030Column21"], input[name="EMR050^^EMR05024Column2"], input[name="EMR050^^EMR05024Column5"]'
  )
  .forEach((radio) => {
    radio.addEventListener('change', updateOutcome);
  });


updateOutcome();

</script>