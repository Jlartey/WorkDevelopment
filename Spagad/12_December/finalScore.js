document.body.addEventListener('mousemove', () => {
  //console.log(optionSelected.length);
  const totalScore = document.querySelector('[id="RES017^^RES017033Column3"]');
  total = 0;
  total += parseFloat(
    document.querySelector('[id="RES017^^RES017007Column3"]')?.value || '0'
  );
  total += parseFloat(
    document.querySelector('[id="RES017^^RES017011Column3"]')?.value || '0'
  );
  total += parseFloat(
    document.querySelector('[id="RES017^^RES017014Column6"]')?.value || '0'
  );
  total += parseFloat(
    document.querySelector('[id="RES017^^RES017018Column3"]')?.value || '0'
  );
  total += parseFloat(
    document.querySelector('[id="RES017^^RES017022Column6"]')?.value || '0'
  );
  total += parseFloat(
    document.querySelector('[id="RES017^^RES017025Column3"]')?.value || '0'
  );
  total += parseFloat(
    document.querySelector('[id="RES017^^RES017028Column3"]')?.value || '0'
  );
  total += parseFloat(
    document.querySelector('[id="RES017^^RES017032Column6"]')?.value || '0'
  );
  totalScore.value = (total / optionSelected.length).toFixed(2);
});
