function sendData(event) {
  event.preventDefault();
  try {
    const sponsorSelect = document.getElementById('sponsors');
    const selectedSponsor = sponsorSelect.selectedOptions[0]
      ? {
          id: sponsorSelect.value,
          value: sponsorSelect.selectedOptions[0].text,
        }
      : null;
    const billMonthSelect = document.getElementById('billMonth');
    const selectedBillMonth = billMonthSelect.selectedOptions[0]
      ? {
          id: billMonthSelect.value,
          value: billMonthSelect.selectedOptions[0].text,
        }
      : null;
    const amount = document.getElementById('amount').value;
    if (!selectedSponsor || !selectedBillMonth || !amount) {
      alert('Please select a sponsor, a bill month, and enter an amount.');
      return;
    }

    let url = 'wpgXMLHTTP.asp?procedurename=InsertSponsorPayment';
    url += '&sponsorID=' + selectedSponsor.id;
    url += '&billMonthID=' + selectedBillMonth.id;
    url += '&amount=' + amount;
    console.log('AJAX URL:', url);

    fetch(url)
      .then((response) => {
        if (!response.ok) {
          throw new Error('Network response was not ok');
        }
        return response.json();
      })
      .then((data) => {
        console.log('Server response:', data);
        if (data.success) {
          alert('Data saved successfully!');
          window.location.reload();
        } else {
          alert('Save failed: ' + (data.message || 'Unknown error'));
        }
      })
      .catch((error) => {
        console.error('AJAX Error:', error);
        alert('Failed to save data: ' + error.message);
      });
  } catch (error) {
    console.error('sendData failed:', error);
    alert('An unexpected error occurred.');
  }
}
