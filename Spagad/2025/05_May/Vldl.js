window.onload = () => {
  const triglyceride = document.getElementById('LL2017^^L0156Column2');
  console.log(triglyceride.value);
  const vldl = document.getElementById('LL2017^^L0021Column2');
  vldl.value = triglyceride / 5;
  console.log(vldl.value);

  const cholesterol = document.getElementById('LL2017^^L0155Column2');
  console.log(cholesterol.value);
  const hdl = document.getElementById('LL2017^^L0160Column2');
  const coronaryRisk = document.getElementById('LL2017^^L0022Column2');
  coronaryRisk.value = cholesterol.value / hdl.value;
  console.log(coronaryRisk.value);
};

// Coronary Risk
window.onload = () => {
  const cholesterol = document.getElementById('LL2017^^L0155Column2');
  console.log(cholesterol.value);
  const hdl = document.getElementById('LL2017^^L0160Column2');
  const coronaryRisk = document.getElementById('LL2017^^L0022Column2');
  coronaryRisk.value = cholesterol.value / hdl.value;
  console.log(coronaryRisk.value);
};
