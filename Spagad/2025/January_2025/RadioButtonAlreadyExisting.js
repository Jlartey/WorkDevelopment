default**
<script>
document.addEventListener('DOMContentLoaded', function() {
document.getElementById('cmdSave').addEventListener('click', function(event) {
            const fields = ['NUR006\^\^NUR00607Column6', 'NUR006\^\^NUR00608Column6', 'NUR006\^\^NUR00609Column6', 'NUR006\^\^NUR00611Column6', 'NUR006\^\^NUR00615Column6', 'NUR006\^\^NUR00613Column6', 'NUR006\^\^NUR00617Column6', 'NUR006\^\^NUR00618Column6'];
            let allFilled = true;
            fields.forEach(field => {
                const inputElement = document.getElementById(field);
                if (!inputElement.value.trim()) {
                    inputElement.classList.add('error');
                    allFilled = false;
                } else {
                    inputElement.classList.remove('error');
                }
            });
            if (!allFilled) {
                event.preventDefault();
                alert("Please fill in all the required fields.");
            }
        });
   });
</script>