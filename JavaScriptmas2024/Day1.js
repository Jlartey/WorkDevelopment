/* ⚠️ You need to add code to this function! ⚠️*/
let itemText = itemInput.value;
itemText = itemText.trim().replace(/\s{2,}/g, ' ');

if (listArr.includes(itemText)) {
  alert('The item has already been added!');
  return;
}

listArr.push(itemText);
renderList();

const itemText = itemInput.value;
const normalisedText = itemText.trim().replace(/\s{2,}/g, '');
if (listItem.includes(normalisedText)) {
  alert('The item has already been added');
}
listArr.push(normalisedText);
