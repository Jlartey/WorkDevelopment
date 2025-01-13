<script>
const radios = document.querySelectorAll('input[type="radio"]');
const subCat = {
    sub1: [
        "RES004^^RES004003Column2",
        "RES004^^RES004003Column5",
    ],
    sub2: [ "RES004^^RES004005Column2",
        "RES004^^RES004005Column5",],

    sub3:[ "RES004^^RES004007Column2","RES004^^RES004007Column5",],

    sub4: [ "RES004^^RES004009Column2",
    "RES004^^RES004011Column2",],

    sub5: ["RES004^^RES004011Column5",],

    sub6: ["RES004^^RES004012Column2",
    "RES004^^RES004014Column2",
    "RES004^^RES004014Column5",],

    sub7: [
        "RES004^^RES004016Column2",
    "RES004^^RES004016Column5",],

    sub8: ["RES004^^RES004018Column2",
    "RES004^^RES004018Column5",],

    sub9: [ "RES004^^RES004020Column2",
        "RES004^^RES004020Column5",
        "RES004^^RES004021Column2",
        "RES004^^RES004021Column5",
        "RES004^^RES004022Column2",],

    sub10: ["RES004^^RES004024Column2",],

    sub11: ["RES004^^RES004026Column2",
    "RES004^^RES004026Column5"]
}

radios.forEach(el=>{
    el.onchange = () =>{
        let sub1Score = 0;
        let sub2Score = 0;
        let sub3Score = 0;
        let sub4Score = 0;
        let sub5Score = 0;
        let sub6Score = 0;
        let sub7Score = 0;
        let sub8Score = 0;
        let sub9Score = 0;
        let sub10Score = 0;
        let sub11Score = 0;
        radios.forEach(el=>{
            if(el.checked){
                if(subCat.sub1.includes(el.getAttribute('name'))){
                    const tempscore = [...document.querySelectorAll(`[name="${el.getAttribute('name')}"]`)]
                    .map((value,index)=>{
                        if(!!el.value&&(value.value == el.value)){
                            return Math.abs(index-5)
                        }
                    }).filter(el=>typeof el === 'number')[0]
                    sub1Score += tempscore || 0
                }else if(subCat.sub2.includes(el.getAttribute('name'))){
                    const tempscore = [...document.querySelectorAll(`[name="${el.getAttribute('name')}"]`)]
                    .map((value,index)=>{
                        if(!!el.value&&(value.value == el.value)){
                            return Math.abs(index-5)
                        }
                    }).filter(el=>typeof el === 'number')[0]
                    sub2Score += tempscore || 0
                    
                }else if(subCat.sub3.includes(el.getAttribute('name'))){
                    const tempscore = [...document.querySelectorAll(`[name="${el.getAttribute('name')}"]`)]
                    .map((value,index)=>{
                        if(!!el.value&&(value.value == el.value)){
                            return Math.abs(index-5)
                        }
                    }).filter(el=>typeof el === 'number')[0]
                    sub3Score += tempscore || 0
                }else if(subCat.sub4.includes(el.getAttribute('name'))){
                    const tempscore = [...document.querySelectorAll(`[name="${el.getAttribute('name')}"]`)]
                    .map((value,index)=>{
                        if(!!el.value&&(value.value == el.value)){
                            return Math.abs(index-5)
                        }
                    }).filter(el=>typeof el === 'number')[0]
                    sub4Score += tempscore || 0

                }else if(subCat.sub5.includes(el.getAttribute('name'))){
                    const tempscore = [...document.querySelectorAll(`[name="${el.getAttribute('name')}"]`)]
                    .map((value,index)=>{
                        if(!!el.value&&(value.value == el.value)){
                            return Math.abs(index-5)
                        }
                    }).filter(el=>typeof el === 'number')[0]
                    sub5Score += tempscore || 0
                }
                }else if(subCat.sub6.includes(el.getAttribute('name'))){
                    const tempscore = [...document.querySelectorAll(`[name="${el.getAttribute('name')}"]`)]
                    .map((value,index)=>{
                        if(!!el.value&&(value.value == el.value)){
                            return Math.abs(index-5)
                        }
                    }).filter(el=>typeof el === 'number')[0]
                    sub6Score += tempscore || 0
                }else if(subCat.sub7.includes(el.getAttribute('name'))){
                    const tempscore = [...document.querySelectorAll(`[name="${el.getAttribute('name')}"]`)]
                    .map((value,index)=>{
                        if(!!el.value&&(value.value == el.value)){
                            return Math.abs(index-5)
                        }
                    }).filter(el=>typeof el === 'number')[0]
                    sub7Score += tempscore || 0
                }else if(subCat.sub8.includes(el.getAttribute('name'))){
                    const tempscore = [...document.querySelectorAll(`[name="${el.getAttribute('name')}"]`)]
                    .map((value,index)=>{
                        if(!!el.value&&(value.value == el.value)){
                            return Math.abs(index-5)
                        }
                    }).filter(el=>typeof el === 'number')[0]
                    sub8Score += tempscore || 0
                }else if(subCat.sub9.includes(el.getAttribute('name'))){
                    const tempscore = [...document.querySelectorAll(`[name="${el.getAttribute('name')}"]`)]
                    .map((value,index)=>{
                        if(!!el.value&&(value.value == el.value)){
                            return Math.abs(index-5)
                        }
                    }).filter(el=>typeof el === 'number')[0]
                    sub9Score += tempscore || 0
                } else if(subCat.sub10.includes(el.getAttribute('name'))){
                    const tempscore = [...document.querySelectorAll(`[name="${el.getAttribute('name')}"]`)]
                    .map((value,index)=>{
                        if(!!el.value&&(value.value == el.value)){
                            return Math.abs(index-5)
                        }
                    }).filter(el=>typeof el === 'number')[0]
                    sub10Score += tempscore || 0
                } else if(subCat.sub11.includes(el.getAttribute('name'))){
                    const tempscore = [...document.querySelectorAll(`[name="${el.getAttribute('name')}"]`)]
                    .map((value,index)=>{
                        if(!!el.value&&(value.value == el.value)){
                            return Math.abs(index-5)
                        }
                    }).filter(el=>typeof el === 'number')[0]
                    sub11Score += tempscore || 0
                }
                    document.querySelector('[id="RES004^^RES004002Column6"]').value = ((((sub1Score||1)-1)/8)*100).toFixed(2);
                    document.querySelector('[id="RES004^^RES004004Column6"]').value = ((((sub2Score||1)-1)/8)*100).toFixed(2);
                    document.querySelector('[id="RES004^^RES004006Column6"]').value = ((((sub3Score||1)-1)/8)*100).toFixed(2);
                    document.querySelector('[id="RES004^^RES004008Column6"]').value = ((((sub4Score||1)-1)/8)*100).toFixed(2);
                    document.querySelector('[id="RES004^^RES004010Column6"]').value = ((((sub5Score||1)-1)/8)*100).toFixed(2);
                    document.querySelector('[id="RES004^^RES004013Column6"]').value = ((((sub6Score||1)-1)/8)*100).toFixed(2);
                    document.querySelector('[id="RES004^^RES004015Column6"]').value = ((((sub7Score||1)-1)/8)*100).toFixed(2);
                    document.querySelector('[id="RES004^^RES004017Column6"]').value = ((((sub8Score||1)-1)/8)*100).toFixed(2);
                    document.querySelector('[id="RES004^^RES004019Column6"]').value = ((((sub9Score||1)-1)/8)*100).toFixed(2);
                    document.querySelector('[id="RES004^^RES004023Column6"]').value = ((((sub10Score||1)-1)/8)*100).toFixed(2);
                    document.querySelector('[id="RES004^^RES004025Column6"]').value = ((((sub11Score||1)-1)/8)*100).toFixed(2);
                    document.querySelector('[id="RES004^^RES004001Column2"]').value = ((((((sub1Score||1)-1)/8)*100) + sub2Score +
                    sub3Score + sub4Score + sub6Score + sub7Score + sub8Score + sub9Score + sub10Score + sub11Score
                        )/(12)).toFixed(2) ;
            
            })
        }
    })
    </script>