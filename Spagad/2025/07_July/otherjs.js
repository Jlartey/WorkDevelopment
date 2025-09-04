Default**<script>
const inpFlds = ["EMR050^^EMR05001Column2", "EMR050^^EMR05001Column4", "EMR050^^EMR05002Column2", 
                 "EMR050^^EMR05002Column4", "EMR050^^EMR05003Column2", "EMR050^^EMR05003Column4", 
                 "EMR050^^EMR05001Column6", "EMR050^^EMR05002Column6", "EMR050^^EMR05008Column3"
                ];
inpFlds.forEach( e => {
    document.getElementById(e).setAttribute("required", true);
} );



window.onload = () => {
  const falls = document.getElementsByName(
    'EMR050^^EMR05017Column11EMRVAR2B-FRAS01^^FRAS01.1Column2'
  );

  for (let i = 0; i < falls.length; i++) {
    falls[i].addEventListener('change', () => {
      if (falls[i].checked) {
        console.log('falls has been cliked!');
      }
    });
  }
};

</script>