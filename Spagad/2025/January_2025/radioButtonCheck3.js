// Check that 'mobilityRadios' selects radio button elements
const mobilityRadios = [
  document.querySelector("input[name='NUR006^^NUR00615Column21']"),
  document.querySelector("input[name='NUR006^^NUR00616Column22']"),
  document.querySelector("input[name='NUR006^^NUR00617Column23']"),
];
const mobility = document.getElementById('NUR006^^NUR00615Column6');

function calculateSum() {
  let sum = 0;

  for (let i = 0; i < mobilityRadios.length; i++) {
    if (mobilityRadios[i] && mobilityRadios[i].checked) {
      sum += 2; //
    }
  }
  mobility.value = sum;
  console.log(`Sum: ${sum}`);
}

for (let i = 0; i < mobilityRadios.length; i++) {
  if (mobilityRadios[i]) {
    mobilityRadios[i].addEventListener('change', calculateSum);
  }
}
