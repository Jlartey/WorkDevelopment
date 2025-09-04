window.onload = function () {
  const ageRadios = document.getElementsByName('NUR006^^NUR00607Column3');
  let age = document.getElementById('NUR006^^NUR00607Column6');

  for (let i = 0; i < ageRadios.length; i++) {
    ageRadios[i].addEventListener('change', function () {
      if (ageRadios[i].checked) {
        age.value = i;
        console.log(i);
      }
      updateTotalScore();
    });
  }

  const fHistoryRadios = document.getElementsByName('NUR006^^NUR00608Column21');
  const fallHistory = document.getElementById('NUR006^^NUR00608Column6');

  for (const radio of fHistoryRadios) {
    radio.addEventListener('change', () => {
      fallHistory.value = radio.checked ? 5 : 0;
      updateTotalScore();
    });
  }

  //Elimination, Bowel and Urination
  const emuRadios = document.getElementsByName('NUR006^^NUR00609Column3');
  const emu = document.getElementById('NUR006^^NUR00609Column6');

  for (let i = 0; i < emuRadios.length; i++) {
    emuRadios[i].addEventListener('change', function () {
      if (emuRadios[1].checked) {
        emu.value = i + 1;
        console.log(i);
      } else {
        emu.value = 0;
      }
      updateTotalScore();
    });
  }

  //Medication
  const medicationRadios = document.getElementsByName(
    'NUR006^^NUR00611Column2'
  );
  const medication = document.getElementById('NUR006^^NUR00611Column6');

  for (let i = 0; i < medicationRadios.length; i++) {
    medicationRadios[i].addEventListener('change', function () {
      if (medicationRadios[i].checked) {
        medication.value = 2 * i + 1;
        console.log(i);
      }
      if (medicationRadios[0].checked) {
        medication.value = 0;
      }
      updateTotalScore();
    });
  }

  const pceRadios = document.getElementsByName('NUR006^^NUR00613Column2');
  const pce = document.getElementById('NUR006^^NUR00613Column6');

  for (let i = 0; i < pceRadios.length; i++) {
    pceRadios[i].addEventListener('change', function () {
      if (pceRadios[i].checked) {
        pce.value = i;
        console.log(i);
      }
      updateTotalScore();
    });
  }

  // Check that 'mobilityRadios' selects radio button elements
  const mobilityRadios = [
    document.querySelector("input[name='NUR006^^NUR00615Column21']"),
    document.querySelector("input[name='NUR006^^NUR00615Column22']"),
    document.querySelector("input[name='NUR006^^NUR00615Column23']"),
  ];

  const mobility = document.getElementById('NUR006^^NUR00615Column6');

  function calculateSum() {
    let sum = 0;

    for (let i = 0; i < mobilityRadios.length; i++) {
      if (mobilityRadios[i] && mobilityRadios[i].checked) {
        sum += 2; // Assuming each checkbox contributes a value of 2
      }
    }
    mobility.value = sum; // Set the calculated sum to the input field
    console.log(`Sum: ${sum}`);
    updateTotalScore();
  }

  // Add event listeners to each radio button
  for (let i = 0; i < mobilityRadios.length; i++) {
    if (mobilityRadios[i]) {
      mobilityRadios[i].addEventListener('change', calculateSum);
    }
  }

  //Cognition
  const cognitionOptions = [
    {
      element: document.getElementsByName('NUR006^^NUR00617Column21')[0],
      value: 1,
    },
    {
      element: document.getElementsByName('NUR006^^NUR00617Column22')[0],
      value: 2,
    },
    {
      element: document.getElementsByName('NUR006^^NUR00617Column23')[0],
      value: 4,
    },
  ];

  const cognition = document.getElementById('NUR006^^NUR00617Column6');

  function calculateSum2() {
    let sum2 = 0;

    // Loop through all checkbox elements and add values if checked
    for (let i = 0; i < cognitionOptions.length; i++) {
      if (cognitionOptions[i].element.checked) {
        sum2 += cognitionOptions[i].value;
      }
    }
    cognition.value = sum2; //
    console.log(`Sum: ${sum2}`); // Log the result
    updateTotalScore();
  }

  for (let i = 0; i < cognitionOptions.length; i++) {
    cognitionOptions[i].element.addEventListener('change', calculateSum2);
  }

  const totalScoreEl = document.getElementById('NUR006^^NUR00618Column6');

  function getNumericValue(element) {
    return parseInt(element.value) || 0; // Convert to integer or use 0 if falsy
  }

  function updateTotalScore() {
    const totalScore =
      getNumericValue(age) +
      getNumericValue(fallHistory) +
      getNumericValue(emu) +
      getNumericValue(medication) +
      getNumericValue(pce) +
      getNumericValue(mobility) +
      getNumericValue(cognition);

    console.log(`Total Score: ${totalScore}`);
    totalScoreEl.value = totalScore;
  }
};
