Default**<script>    
    let selector = '[name="RES017^^RES017003Column3"],[name="RES017^^RES017004Column3"]'
    selector += ',[name="RES017^^RES017005Column3"],[name="RES017^^RES017006Column3"]'
    selector += ',[name="RES017^^RES017003Column6"],[name="RES017^^RES017004Column6"]'
    selector += ',[name="RES017^^RES017005Column6"],[name="RES017^^RES017006Column6"],[name="RES017^^E000266Column3"],[name="RES017^^E000266Column6"]'
    const scoreAllocation = {
            ["RES023007001"]:0,
            ["RES023007002"]:50,
            ["RES023007003"]:100,
    }

    const radios = document.querySelectorAll(selector);
    const phyScore = document.querySelector('[id="RES017^^RES017007Column3"]')
    radios.forEach(el=>{
        el.onchange = () =>{
            let accScore = 0
            radios.forEach(el=>{
                if(el.checked){
                    accScore += scoreAllocation[el.value] ?? 0;
                    phyScore.value = accScore
                }
            })
        }
    })
</script>