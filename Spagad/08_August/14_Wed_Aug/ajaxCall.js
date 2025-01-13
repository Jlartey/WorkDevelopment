<script>
    document.getElementById('callApiButton').addEventListener('click', function() {
        var xhr = new XMLHttpRequest();
        var url = 'wpgxmlhttp.asp?procedurename=generatepatientinvoice&tablename=labbydoctor&KeyPrefix=E100002747-TH060&LabtestID=L324||L405';
        xhr.open('GET', url, true);
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4 && xhr.status === 200) {
                console.log(xhr.responseText);
            }
        };
        xhr.send();
    });
</script>