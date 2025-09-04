let groupScores = {
  FALS: 0,
  AMB: 0,
  GAIT: 0,
  MENT: 0,
  MEDIR: 0,
};

let groupScores2 = {
  AgeRange01: 0,
  GEND01: 0,
  Diag: 0,
  CogniImp: 0,
  EnvFac: 0,
  RespSed: 0,
  MedUse: 0,
};

function updateFinalScore() {
  let total = 0;
  for (let group of Object.keys(groupScores)) {
    total += groupScores[group];
  }

  // Select total score element
  const totalElement = document.querySelector(
    '#EMR050\\^\\^EMR05035Column11EMRVAR2B-FRAS01\\^\\^FRAS01\\.4Column2'
  );
  if (totalElement) {
    totalElement.value = total;
    totalElement.setAttribute('value', total);
    ['input', 'change'].forEach((eventType) => {
      let event;
      if (typeof Event === 'function') {
        event = new Event(eventType, { bubbles: true });
      } else {
        event = document.createEvent('Event');
        event.initEvent(eventType, true, true);
      }
      totalElement.dispatchEvent(event);
    });
  } else {
    console.error('Total score input not found for mfs.');
  }

  // Select risk level element
  const riskLevelElement = document.querySelector(
    '#EMR050\\^\\^EMR05035Column11EMRVAR2B-FRAS01\\^\\^FRAS01\\.4Column5'
  );
  if (riskLevelElement) {
    const riskThresholds = { low: 25, moderate: 45 };
    riskLevelElement.value =
      total < riskThresholds.low
        ? 'Low Risk'
        : total < riskThresholds.moderate
        ? 'Moderate Risk'
        : 'High Risk';
    riskLevelElement.setAttribute('value', riskLevelElement.value);
    ['input', 'change'].forEach((eventType) => {
      let event;
      if (typeof Event === 'function') {
        event = new Event(eventType, { bubbles: true });
      } else {
        event = document.createEvent('Event');
        event.initEvent(eventType, true, true);
      }
      riskLevelElement.dispatchEvent(event);
    });
  } else {
    console.error('Risk level input not found for mfs.');
  }

  // Select action element
  const actionElement = document.querySelector(
    '#EMR050\\^\\^EMR05035Column11EMRVAR2B-FRAS01\\^\\^FRAS01\\.5Column2'
  );
  if (actionElement) {
    const riskThresholds = { low: 25, moderate: 45 };
    actionElement.value =
      total < riskThresholds.low
        ? 'Routine Care'
        : total < riskThresholds.moderate
        ? 'Provide verbal warning; consider mobility aid review'
        : 'Implement fall precautions, flag in record, consider referral for assessment';
    actionElement.setAttribute('value', actionElement.value);
    ['input', 'change'].forEach((eventType) => {
      let event;
      if (typeof Event === 'function') {
        event = new Event(eventType, { bubbles: true });
      } else {
        event = document.createEvent('Event');
        event.initEvent(eventType, true, true);
      }
      actionElement.dispatchEvent(event); // Fixed to dispatch to actionElement
    });
  } else {
    console.error('Action element not found for mfs.');
  }

  console.log('mfs Final Score: ' + total);
}

function updateFinalScore2() {
  let total = 0;
  for (let group of Object.keys(groupScores2)) {
    total += groupScores2[group];
  }

  // Select total score element
  const totalElement = document.querySelector(
    '#EMR050\\^\\^EMR05035Column12EMRVAR2B-FRAS02\\^\\^FRAS02\\.5Column2'
  );
  if (totalElement) {
    totalElement.value = total;
    totalElement.setAttribute('value', total);
    ['input', 'change'].forEach((eventType) => {
      let event;
      if (typeof Event === 'function') {
        event = new Event(eventType, { bubbles: true });
      } else {
        event = document.createEvent('Event');
        event.initEvent(eventType, true, true);
      }
      totalElement.dispatchEvent(event);
    });
  } else {
    console.error('Total score input not found for adhd.');
  }

  // Select risk level element
  const riskLevelElement = document.getElementById(
    'EMR050^^EMR05035Column12EMRVAR2B-FRAS02^^FRAS02.5Column5'
  );
  if (riskLevelElement) {
    const riskThresholds = { low: 12, moderate: 18 };
    riskLevelElement.value =
      total < riskThresholds.low
        ? 'Low Risk'
        : total < riskThresholds.moderate
        ? 'Moderate Risk'
        : 'High Risk';
    riskLevelElement.setAttribute('value', riskLevelElement.value);
    ['input', 'change'].forEach((eventType) => {
      let event;
      if (typeof Event === 'function') {
        event = new Event(eventType, { bubbles: true });
      } else {
        event = document.createEvent('Event');
        event.initEvent(eventType, true, true);
      }
      riskLevelElement.dispatchEvent(event);
    });
  } else {
    console.error('Risk level input not found for adhd.');
  }

  const actionElement = document.getElementById(
    'EMR050^^EMR05035Column12EMRVAR2B-FRAS02^^FRAS02.6Column2'
  );
  if (actionElement) {
    const riskThresholds = { low: 12, moderate: 18 };
    actionElement.value =
      total < riskThresholds.low
        ? 'Basic safety education, standard observation'
        : total < riskThresholds.moderate
        ? 'Flag chart, caregiver alert, assist child as needed'
        : 'Enhanced monitoring, escort if needed, fall precautions';
    actionElement.setAttribute('value', actionElement.value);
    ['input', 'change'].forEach((eventType) => {
      let event;
      if (typeof Event === 'function') {
        event = new Event(eventType, { bubbles: true });
      } else {
        event = document.createEvent('Event');
        event.initEvent(eventType, true, true);
      }
      actionElement.dispatchEvent(event); // Fixed to dispatch to actionElement
    });
  } else {
    console.error('Action element not found for mfs.');
  }
  console.log('adhd Final Score: ' + total);
}

document.addEventListener('click', function (e) {
  if (e.target.matches('input[type="radio"]')) {
    let score = 0;
    let group = null;
    switch (e.target.value) {
      // mfs section (groupScores)
      case 'FALS01':
        score = 0;
        group = 'FALS';
        break;
      case 'FALS02':
        score = 25;
        group = 'FALS';
        break;
      case 'FALS03':
        score = 15;
        group = 'FALS';
        break;
      case 'GAIT01':
        score = 0;
        group = 'GAIT';
        break;
      case 'GAIT02':
        score = 10;
        group = 'GAIT';
        break;
      case 'GAIT03':
        score = 20;
        group = 'GAIT';
        break;
      case 'MENT01':
        score = 0;
        group = 'MENT';
        break;
      case 'MENT02':
        score = 15;
        group = 'MENT';
        break;
      case 'AMB01':
        score = 0;
        group = 'AMB';
        break;
      case 'AMB02':
        score = 15;
        group = 'AMB';
        break;
      case 'AMB03':
        score = 30;
        group = 'AMB';
        break;
      case 'MEDIR01':
        score = 0;
        group = 'MEDIR';
        break;
      case 'MEDIR02':
        score = 10;
        group = 'MEDIR';
        break;
      // adhd section (groupScores2)
      case 'AgeRange01.1':
        score = 4;
        group = 'AgeRange01';
        break;
      case 'AgeRange01.2':
        score = 3;
        group = 'AgeRange01';
        break;
      case 'AgeRange01.3':
        score = 2;
        group = 'AgeRange01';
        break;
      case 'AgeRange01.4':
        score = 1;
        group = 'AgeRange01';
        break;
      case 'GEND01.1':
        score = 1;
        group = 'GEND01';
        break;
      case 'GEND01.2':
        score = 2;
        group = 'GEND01';
        break;
      case 'Diag01':
        score = 4;
        group = 'Diag';
        break;
      case 'Diag02':
        score = 2;
        group = 'Diag';
        break;
      case 'Diag03':
        score = 1;
        group = 'Diag';
        break;
      case 'CogniImp01':
        score = 3;
        group = 'CogniImp';
        break;
      case 'CogniImp02':
        score = 2;
        group = 'CogniImp';
        break;
      case 'CogniImp03':
        score = 1;
        group = 'CogniImp';
        break;
      case 'EnvFac01':
        score = 4;
        group = 'EnvFac';
        break;
      case 'EnvFac02':
        score = 2;
        group = 'EnvFac';
        break;
      case 'EnvFac03':
        score = 1;
        group = 'EnvFac';
        break;
      case 'RespSed01':
        score = 3;
        group = 'RespSed';
        break;
      case 'RespSed02':
        score = 2;
        group = 'RespSed';
        break;
      case 'RespSed03':
        score = 1;
        group = 'RespSed';
        break;
      case 'MedUse01':
        score = 3;
        group = 'MedUse';
        break;
      case 'MedUse02':
        score = 2;
        group = 'MedUse';
        break;
      case 'MedUse03':
        score = 1;
        group = 'MedUse';
        break;
      default:
        score = 0;
        group = null;
        break; // Handle empty or unknown values
    }
    if (group) {
      if (['FALS', 'GAIT', 'MENT', 'AMB', 'MEDIR'].includes(group)) {
        groupScores[group] = score;
        updateFinalScore();
      } else {
        groupScores2[group] = score;
        updateFinalScore2();
      }
    }
  }
});
