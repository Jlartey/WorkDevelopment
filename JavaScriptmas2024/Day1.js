// Get references to DOM elements
const itemInput = document.getElementById('item-input');
const addItemButton = document.getElementById('add-item-button');
const shoppingList = document.getElementById('shopping-list');
const listArr = []; // Stores original case items
const normalizedArr = []; // Stores normalized versions for duplicate checking

// Function to check if item is a duplicate
function checkDuplicate() {
  let itemText = itemInput.value.trim().replace(/\s+/g, ' ');
  if (!itemText) return; // Prevent empty input

  let normalizedText = itemText.toLowerCase();

  if (normalizedArr.includes(normalizedText)) {
    alert('The item has already been added');
    return;
  }

  listArr.push(itemText); // Store original case version
  normalizedArr.push(normalizedText); // Store normalized version

  renderList();
}

// Function to render the shopping list
function renderList() {
  shoppingList.innerHTML = '';
  listArr.forEach((gift, index) => {
    const listItem = document.createElement('li');
    listItem.textContent = gift;

    // Create Edit button
    const editButton = document.createElement('button');
    editButton.textContent = 'Edit';
    editButton.onclick = () => editItem(index);

    // Create Delete button
    const deleteButton = document.createElement('button');
    deleteButton.textContent = 'Delete';
    deleteButton.onclick = () => deleteItem(index);

    listItem.appendChild(editButton);
    listItem.appendChild(deleteButton);
    shoppingList.appendChild(listItem);
  });
  itemInput.value = ''; // Clear input field
}

// Function to delete an item
function deleteItem(index) {
  listArr.splice(index, 1);
  normalizedArr.splice(index, 1);
  renderList();
}

// Function to edit an item
function editItem(index) {
  let newValue = prompt('Edit the item:', listArr[index]);
  if (newValue) {
    let normalizedNew = newValue.trim().replace(/\s+/g, ' ').toLowerCase();
    if (
      normalizedArr.includes(normalizedNew) &&
      normalizedNew !== normalizedArr[index]
    ) {
      alert('This edited item already exists.');
      return;
    }
    normalizedArr[index] = normalizedNew; // Update normalized array
    listArr[index] = newValue.trim().replace(/\s+/g, ' ');
    renderList();
  }
}

// Add event listeners
addItemButton.addEventListener('click', checkDuplicate);
itemInput.addEventListener('keypress', (event) => {
  if (event.key === 'Enter') {
    checkDuplicate();
  }
});
