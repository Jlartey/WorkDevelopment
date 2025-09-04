window.onload = () => {
  const falls = document.getElementsByName('EMR050^^FRAS01.1Column2');
  let mfs_total = document.getElementById('EMR050^^FRAS01.4Column2');

  const riskLevel = document.getElementById('EMR050^^FRAS01.4Column5');

  let fallsSum;

  for (let i = 0; i < falls.length; i++) {
    falls[i].addEventListener('change', () => {
      if (falls[i].checked) {
        console.log('falls has been clicked!');
        const result = [0, 0, 15, 25];
        fallsSum = result[i];
      }
      updateTotalScore();
    });
  }

  const amb_aid = document.getElementsByName('EMR050^^FRAS01.1Column5');
  let ambAidSum;

  for (let i = 0; i < amb_aid.length; i++) {
    amb_aid[i].addEventListener('change', () => {
      if (amb_aid[i].checked) {
        const result = [0, 0, 15, 30];
        ambAidSum = result[i];
      }
      updateTotalScore();
    });
  }

  const gait = document.getElementsByName('EMR050^^FRAS01.2Column2');
  let gaitSum;

  for (let i = 0; i < gait.length; i++) {
    gait[i].addEventListener('change', () => {
      if (gait[i].checked) {
        const result = [0, 0, 10, 20];
        gaitSum = result[i];
      }
      updateTotalScore();
    });
  }

  const mental = document.getElementsByName('EMR050^^FRAS01.2Column5');
  let mentalSum;

  for (let i = 0; i < mental.length; i++) {
    mental[i].addEventListener('change', () => {
      if (mental[i].checked) {
        const result = [0, 0, 15];
        mentalSum = result[i];
      }
      updateTotalScore();
    });
  }

  const review = document.getElementsByName('EMR050^^FRAS01.3Column2');
  let reviewSum;

  for (let i = 0; i < review.length; i++) {
    review[i].addEventListener('change', () => {
      if (review[i].checked) {
        const result = [0, 0, 10];
        reviewSum = result[i];
      }
      updateTotalScore();
    });
  }

  //Adapted Humpty Dumpty
  const ageRange = document.getElementsByName('EMR050^^FRAS02.1Column2');

  let ageRangeSum;
  for (let i = 0; i < ageRange.length; i++) {
    ageRange[i].addEventListener('change', () => {
      if (ageRange[i].checked) {
        const result = [0, 4, 3, 2, 1];
        ageRangeSum = result[i];
      }
      updateTotalScore_adhd();
    });
  }

  const gender = document.getElementsByName('EMR050^^FRAS02.1Column5');
  let genderSum;
  for (let i = 0; i < gender.length; i++) {
    gender[i].addEventListener('change', () => {
      if (gender[i].checked) {
        const result = [0, 2, 1];
        genderSum = result[i];
      }
      updateTotalScore_adhd();
    });
  }

  const diagnosis = document.getElementsByName('EMR050^^FRAS02.2Column2');
  let diagnosisSum;
  for (let i = 0; i < diagnosis.length; i++) {
    diagnosis[i].addEventListener('change', () => {
      if (diagnosis[i].checked) {
        const result = [0, 4, 2, 1];
        diagnosisSum = result[i];
      }
      updateTotalScore_adhd();
    });
  }

  const cogImp = document.getElementsByName('EMR050^^FRAS02.2Column5');
  let cogImpSum;
  for (let i = 0; i < cogImp.length; i++) {
    cogImp[i].addEventListener('change', () => {
      if (cogImp[i].checked) {
        const result = [0, 3, 2, 1];
        cogImpSum = result[i];
      }
      updateTotalScore_adhd();
    });
  }

  const envFac = document.getElementsByName('EMR050^^FRAS02.3Column2');
  let envFacSum;
  for (let i = 0; i < envFac.length; i++) {
    envFac[i].addEventListener('change', () => {
      if (envFac[i].checked) {
        const result = [0, 4, 2, 1];
        envFacSum = result[i];
      }
      updateTotalScore_adhd();
    });
  }

  const respSed = document.getElementsByName('EMR050^^FRAS02.3Column5');
  let respSedSum;
  for (let i = 0; i < respSed.length; i++) {
    respSed[i].addEventListener('change', () => {
      if (respSed[i].checked) {
        const result = [0, 3, 2, 1];
        respSedSum = result[i];
      }
      updateTotalScore_adhd();
    });
  }

  const medUse = document.getElementsByName('EMR050^^FRAS02.4Column2');
  let medUseSum;
  for (let i = 0; i < medUse.length; i++) {
    medUse[i].addEventListener('change', () => {
      if (medUse[i].checked) {
        const result = [0, 3, 2, 1];
        medUseSum = result[i];
      }
      updateTotalScore_adhd();
    });
  }

  function getNumericValue(element) {
    return parseInt(element) || 0;
  }

  function updateTotalScore() {
    const totalScore =
      getNumericValue(fallsSum) +
      getNumericValue(ambAidSum) +
      getNumericValue(gaitSum) +
      getNumericValue(mentalSum) +
      getNumericValue(reviewSum);

    console.log(`Total Score: ${totalScore}`);
    mfs_total.value = totalScore;

    if (totalScore < 25) {
      riskLevel.value = 'Low Risk';
    } else if (totalScore < 45) {
      riskLevel.value = 'Moderate Risk';
    } else {
      riskLevel.value = 'High Risk';
    }
  }

  function updateTotalScore_adhd() {
    const ahdt_total = document.getElementById('EMR050^^FRAS02.5Column2');
    const riskLevelAhdt = document.getElementById('EMR050^^FRAS02.5Column5');
    const totalScore =
      getNumericValue(ageRangeSum) +
      getNumericValue(genderSum) +
      getNumericValue(diagnosisSum) +
      getNumericValue(cogImpSum) +
      getNumericValue(envFacSum) +
      getNumericValue(respSedSum) +
      getNumericValue(medUseSum);

    ahdt_total.value = totalScore;

    if (totalScore < 12) {
      riskLevelAhdt.value = 'Low Risk';
    } else if (totalScore < 18) {
      riskLevelAhdt.value = 'Moderate Risk';
    } else {
      riskLevelAhdt.value = 'High Risk';
    }
  }
};
