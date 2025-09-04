window.onload = () => {
  const sensoryRadios = document.getElementsByName('NUR012^^NUR012003Column3');
  const sensory = document.getElementById('NUR012^^NUR012003Column4');

  for (let i = 0; i < sensoryRadios.length; i++) {
    sensoryRadios[i].addEventListener('change', () => {
      if (sensoryRadios[i].checked) {
        sensory.value = i;
        console.log(i);
      }
      updateTotalScore();
    });
  }

  //Moisture
  const moistureRadios = document.getElementsByName('NUR012^^NUR012004Column3');
  const moisture = document.getElementById('NUR012^^NUR012004Column4');

  for (let i = 0; i < moistureRadios.length; i++) {
    moistureRadios[i].addEventListener('change', () => {
      if (moistureRadios[i].checked) {
        moisture.value = i;
        console.log(i);
      }
      updateTotalScore();
    });
  }

  //Activity
  const activityRadios = document.getElementsByName('NUR012^^NUR012005Column3');
  const activity = document.getElementById('NUR012^^NUR012005Column4');

  for (let i = 0; i < activityRadios.length; i++) {
    activityRadios[i].addEventListener('change', () => {
      if (activityRadios[i].checked) {
        activity.value = i;
        console.log(i);
      }
      updateTotalScore();
    });
  }

  //Mobility
  const mobilityRadios = document.getElementsByName('NUR012^^NUR012006Column3');
  const mobility = document.getElementById('NUR012^^NUR012006Column4');

  for (let i = 0; i < mobilityRadios.length; i++) {
    mobilityRadios[i].addEventListener('change', () => {
      if (mobilityRadios[i].checked) {
        mobility.value = i;
        console.log(i);
      }
      updateTotalScore();
    });
  }

  //Nutrition
  const nutritionRadios = document.getElementsByName(
    'NUR012^^NUR012007Column3'
  );
  const nutrition = document.getElementById('NUR012^^NUR012007Column4');

  for (let i = 0; i < nutritionRadios.length; i++) {
    nutritionRadios[i].addEventListener('change', () => {
      if (nutritionRadios[i].checked) {
        nutrition.value = i;
        console.log(i);
      }
      updateTotalScore();
    });
  }

  //Friction
  const frictionRadios = document.getElementsByName('NUR012^^NUR012008Column3');
  const friction = document.getElementById('NUR012^^NUR012008Column4');

  for (let i = 0; i < frictionRadios.length; i++) {
    frictionRadios[i].addEventListener('change', () => {
      if (frictionRadios[i].checked) {
        friction.value = i;
        console.log(i);
      }
      updateTotalScore();
    });
  }

  const totalScoreEl = document.getElementById('NUR012^^NUR012009Column4');

  function getNumericValue(element) {
    return parseInt(element.value) || 0;
  }

  function updateTotalScore() {
    const totalScore =
      getNumericValue(sensory) +
      getNumericValue(moisture) +
      getNumericValue(activity) +
      getNumericValue(mobility) +
      getNumericValue(nutrition) +
      getNumericValue(friction);

    console.log(`Total Score: ${totalScore}`);
    totalScoreEl.value = totalScore;
  }
};
