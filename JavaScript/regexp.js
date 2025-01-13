const regex = /pattern/;
const inputString = 'input string';

const matches = inputString.match(regex);

if (matches) {
  // Process the matches
  for (const match of matches) {
    console.log(match);
  }
} else {
  console.log('No matches found.');
}

const regex2 = /pattern/;
const inputString2 = 'input string';

if (regex2.test(inputString2)) {
  console.log('Pattern found in input string.');
} else {
  console.log('Pattern not found in input string.');
}

response.write " fetch(url)"
response.write "   .then((response) => response.json())"
response.write "   .then((results) => {"
response.write "     if (results.success) {"
response.write "       const tableContainer = document.querySelector('.mytable>tbody');"
response.write "       const tr = document.createElement('tr');"
response.write "       tr.innerHTML = `<td class='mytd'>${"
response.write "         tableContainer.querySelectorAll('tr').length + 1"
response.write "       }</td>"
response.write "                       <td class='mytd'>${"
response.write "                         results.data.KeyPrefix.split('||')[0]"
response.write "                       }</td>"
response.write "                       <td class='mytd'>${results.data.PerformVarName}</td>"
response.write "                       <td class='mytd'>${results.data.Description}</td>"
response.write "                       <td class='mytd'>${"
response.write "                         results.data.KeyPrefix.split('||')[1]"
response.write "                       }</td>`;"
response.write "       tableContainer.appendChild(tr);"
response.write "     }"
response.write "   })"
response.write "   .catch((error) => console.error('Fetch error:', error));"
response.write "    const tds = tableContainer.querySelectorAll('[id=""count""]');"
response.write "    tds.forEach((el,index) => {"
response.write "        index++;"
response.write "        tds.innerText = index;"
response.write "    });"