TDProp**3**Left**Top%%Default**
<script> 
  let groupScores = {};
function updateFinalScore() {
    let total = 0;
    for (let group of Object.keys(groupScores)) {
        total += groupScores[group];
    }
    let target = document.getElementById("EMR050^^EMR05035Column11EMRVAR2B-FRAS01^^FRAS01.4Column2");
    if (target) {
        target.value = total;
        target.setAttribute("value", total);
        ["input", "change"].forEach(eventType => {
            let event;
            if (typeof(Event) === "function") {
                event = new Event(eventType, { bubbles: true });
            } else {
                event = document.createEvent("Event");
                event.initEvent(eventType, true, true);
            }
            target.dispatchEvent(event);
        });
    } else {
        console.error("Target input not found. Check the element ID.");
    }
    console.log("Final Score: " + total);
}
document.addEventListener('click', function(e) {
    if (e.target.matches('input[type="radio"]')) {
        let score = 0;
        let group = null;
        switch (e.target.value) {
            case "": score = 0; group = 'FALS'; break;
            case 'FALS01': score = 0; group = 'FALS'; break;
            case 'FALS02': score = 25; group = 'FALS'; break;
            case 'FALS03': score = 15; group = 'FALS'; break;
            case "": score = 0;  group = 'GAIT'; break;
            case 'GAIT01': score = 0;  group = 'GAIT'; break;
            case 'GAIT02': score = 10; group = 'GAIT'; break;
            case 'GAIT03': score = 20; group = 'GAIT'; break;
            case "": score = 0;  group = 'MENT'; break;
            case 'MENT01': score = 0;  group = 'MENT'; break;
            case 'MENT02': score = 15; group = 'MENT'; break;
            case "": score = 0; group = 'AMB'; break;
            case 'AMB01': score = 0; group = 'AMB'; break;
            case 'AMB02': score = 15; group = 'AMB'; break;
            case 'AMB03': score = 30; group = 'AMB'; break;
            case "": score = 0;  group = 'MEDIR'; break;
            case 'MEDIR01': score = 0;  group = 'MEDIR'; break;
            case 'MEDIR02': score = 10; group = 'MEDIR'; break;
        }
        if (group) {
            groupScores[group] = score;
            updateFinalScore();
        }
    }
});
</script>