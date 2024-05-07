// const userDetails = {
//   firstName: '',
//   lastName: '',
//   age: null,
// };

// document.addEventListener('DOMContentLoaded', () => {
//   const form = document.getElementById('basic-form');

//   form.addEventListener('submit', (event) => {
//     event.preventDefault();

//     // grab the inputs
//     const firstName = document.getElementById('firstName').value.trim();
//     const lastName = document.getElementById('lastName').value.trim();
//     const age = document.getElementById('age').value.trim();

//     if (firstName === '' && typeof firstName !== 'string') {
//       alert('Please enter a valid first name');

//       return;
//     }

//     if (lastName === '' && typeof lastName !== 'string') {
//       alert('Please enter a valid last name');
//       return;
//     }

//     if (age === '' || isNaN(age) || age < 1 || age > 100) {
//       alert('Please enter a valid age between 1 and 100');
//       return;
//     }

//     alert('Form has been submitted successfully!');
//     userDetails.firstName = firstName;
//     userDetails.lastName = lastName;
//     userDetails.age = +age;
//     console.log(userDetails);
//     console.log(typeof firstName);
//   });
// });

// Alternative Approach

// Function to validate the form inputs
// Function to validate the form inputs
function validateForm(firstName, lastName, age) {
  if (firstName === '') {
    alert('Please enter a valid first name');
    return false;
  }

  if (!isNaN(firstName)) {
    alert('First name cannot be a number');
    return false;
  }

  if (lastName === '') {
    alert('Please enter a valid last name');
    return false;
  }

  if (!isNaN(lastName)) {
    alert('Last name cannot be a number');
    return false;
  }

  if (age === '' || isNaN(age) || age < 1 || age > 100) {
    alert('Please enter a valid age between 1 and 100');
    return false;
  }

  return true;
}

// Function to handle form submission
function handleSubmit(event) {
  event.preventDefault(); // Prevent form submission
  const form = event.target; // Get the form element

  const firstName = form.querySelector('.firstName').value.trim();
  const lastName = form.querySelector('.lastName').value.trim();
  const age = form.querySelector('.age').value.trim();

  if (validateForm(firstName, lastName, age)) {
    alert('Form submitted successfully!');
    console.log(firstName, lastName, age);
  }
}

// Attach event listener to the form submission event for the form with id "basic-form"
const form = document.querySelector('#basic-form');
form.addEventListener('submit', handleSubmit);
