let groupScores2 = {
  AgeRange01: 0,
  GEND01: 0,
  Diag: 0,
  CogniImp: 0,
  EnvFac: 0,
  RespSed: 0,
  MedUse: 0,
};

// function updateFinalScore() {
//   let total = 0;
//   for (let group of Object.keys(groupScores)) {
//     total += groupScores[group];
//   }

//   // Select total score element (escaping invalid characters)
//   const totalElement = document.querySelector("#EMR050\\^\\^EMR05035Column11EMRVAR2B-FRAS01\\^\\^FRAS01\\.4Column2");
//   if (totalElement) {
//     totalElement.value = total;
//     totalElement.setAttribute("value", total);
//     ["input", "change"].forEach(eventType => {
//       let event;
//       if (typeof Event === "function") {
//         event = new Event(eventType, { bubbles: true });
//       } else {
//         event = document.createEvent("Event");
//         event.initEvent(eventType, true, true);
//       }
//       totalElement.dispatchEvent(event);
//     });
//   } else {
//     console.error("Total score input not found. Check the element ID.");
//   }

//   // Select risk level element
//   const riskLevelElement = document.querySelector("#EMR050\\^\\^EMR05035Column11EMRVAR2B-FRAS01\\^\\^FRAS01\\.4Column5");
//   if (riskLevelElement) {
//     // Set risk level based on thresholds
//     console.log('Working well')
//     const riskThresholds = { low: 25, moderate: 45 };
//     riskLevelElement.value =
//       total < riskThresholds.low
//         ? "Low Risk"
//         : total < riskThresholds.moderate
//         ? "Moderate Risk"
//         : "High Risk";
//     riskLevelElement.setAttribute("value", riskLevelElement.value);
//     ["input", "change"].forEach(eventType => {
//       let event;
//       if (typeof Event === "function") {
//         event = new Event(eventType, { bubbles: true });
//       } else {
//         event = document.createEvent("Event");
//         event.initEvent(eventType, true, true);
//       }
//       riskLevelElement.dispatchEvent(event);
//     });
//   } else {
//     console.error("Risk level input not found. Check the element ID.");
//   }

//   const actionElement = document.getElementById("EMR050^^EMR05035Column11EMRVAR2B-FRAS01^^FRAS01.5Column2");
//   if (actionElement) {
//     // Set risk level based on thresholds
//     console.log('Working well')
//     const riskThresholds = { low: 25, moderate: 45 };
//     actionElement.value =
//       total < riskThresholds.low
//         ? "Routine Care"
//         : total < riskThresholds.moderate
//         ? "Provide verbal warning; consider mobility aid review"
//         : "Implement fall precautions, flag in record, consider referral for assessment";
//     actionElement.setAttribute("value", actionElement.value);
//     ["input", "change"].forEach(eventType => {
//       let event;
//       if (typeof Event === "function") {
//         event = new Event(eventType, { bubbles: true });
//       } else {
//         event = document.createEvent("Event");
//         event.initEvent(eventType, true, true);
//       }
//       riskLevelElement.dispatchEvent(event);
//     });
//   } else {
//     console.error("Risk level input not found. Check the element ID.");
//   }
//   console.log("Final Score: " + total);
// }

function updateFinalScore2() {
  let total = 0;
  for (let group of Object.keys(groupScores)) {
    total += groupScores[group];
  }

  // Select total score element (escaping invalid characters)
  const totalElement = document.getElementById(
    'EMR050^^EMR05035Column12EMRVAR2B-FRAS02^^FRAS02.5Column2'
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
    console.error('Total score input not found. Check the element ID.');
  }
  console.log('Final Score: ' + total);
}

// document.addEventListener("click", function (e) {
//   if (e.target.matches('input[type="radio"]')) {
//     let score = 0;
//     let group = null;
//     switch (e.target.value) {
//       case "": score = 0; group = "FALS"; break;
//       case "FALS01": score = 0; group = "FALS"; break;
//       case "FALS02": score = 25; group = "FALS"; break;
//       case "FALS03": score = 15; group = "FALS"; break;
//       case "": score = 0; group = "GAIT"; break;
//       case "GAIT01": score = 0; group = "GAIT"; break;
//       case "GAIT02": score = 10; group = "GAIT"; break;
//       case "GAIT03": score = 20; group = "GAIT"; break;
//       case "": score = 0; group = "MENT"; break;
//       case "MENT01": score = 0; group = "MENT"; break;
//       case "MENT02": score = 15; group = "MENT"; break;
//       case "": score = 0; group = "AMB"; break;
//       case "AMB01": score = 0; group = "AMB"; break;
//       case "AMB02": score = 15; group = "AMB"; break;
//       case "AMB03": score = 30; group = "AMB"; break;
//       case "": score = 0; group = "MEDIR"; break;
//       case "MEDIR01": score = 0; group = "MEDIR"; break;
//       case "MEDIR02": score = 10; group = "MEDIR"; break;
//     }
//     if (group) {
//       groupScores[group] = score;
//       updateFinalScore();
//     }
//   }
// });

document.addEventListener('click', function (e) {
  if (e.target.matches('input[type="radio"]')) {
    let score = 0;
    let group = null;
    switch (e.target.value) {
      case '':
        score = 0;
        group = 'AgeRange01';
        break;
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
      // case "": score = 0; group = "GAIT"; break;
      // case "GAIT01": score = 0; group = "GAIT"; break;
      // case "GAIT02": score = 10; group = "GAIT"; break;
      // case "GAIT03": score = 20; group = "GAIT"; break;
      // case "": score = 0; group = "MENT"; break;
      // case "MENT01": score = 0; group = "MENT"; break;
      // case "MENT02": score = 15; group = "MENT"; break;
      // case "": score = 0; group = "AMB"; break;
      // case "AMB01": score = 0; group = "AMB"; break;
      // case "AMB02": score = 15; group = "AMB"; break;
      // case "AMB03": score = 30; group = "AMB"; break;
      // case "": score = 0; group = "MEDIR"; break;
      // case "MEDIR01": score = 0; group = "MEDIR"; break;
      // case "MEDIR02": score = 10; group = "MEDIR"; break;
    }
    if (group) {
      groupScores2[group] = score;
      updateFinalScore2();
    }
  }
});
