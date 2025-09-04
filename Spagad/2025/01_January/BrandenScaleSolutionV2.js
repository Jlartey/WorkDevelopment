window.onload = () => {
  const radioGroups = [
    { name: 'NUR012^^NUR012003Column3', outputId: 'NUR012^^NUR012003Column4' },
    { name: 'NUR012^^NUR012004Column3', outputId: 'NUR012^^NUR012004Column4' },
    { name: 'NUR012^^NUR012005Column3', outputId: 'NUR012^^NUR012005Column4' },
    { name: 'NUR012^^NUR012006Column3', outputId: 'NUR012^^NUR012006Column4' },
    { name: 'NUR012^^NUR012007Column3', outputId: 'NUR012^^NUR012007Column4' },
    { name: 'NUR012^^NUR012008Column3', outputId: 'NUR012^^NUR012008Column4' },
  ];

  function setupRadioGroup(radioName, outputId) {
    const radios = document.getElementsByName(radioName);
    const output = document.getElementById(outputId);

    radios.forEach((radio, i) => {
      radio.addEventListener('change', () => {
        if (radio.checked) {
          output.value = i;
          console.log(i);
        }
        updateTotalScore();
      });
    });
  }

  radioGroups.forEach((group) => setupRadioGroup(group.name, group.outputId));

  const totalScoreEl = document.getElementById('NUR012^^NUR012009Column4');

  function getNumericValue(element) {
    return parseInt(element.value) || 0;
  }

  function updateTotalScore() {
    const totalScore = radioGroups.reduce((sum, group) => {
      const element = document.getElementById(group.outputId);
      return sum + getNumericValue(element);
    }, 0);

    console.log(`Total Score: ${totalScore}`);
    totalScoreEl.value = totalScore;
  }
};
