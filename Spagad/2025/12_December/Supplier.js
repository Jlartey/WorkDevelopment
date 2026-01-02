// <!DOCTYPE html>
// <html lang="en">
// <head>
//     <meta charset="UTF-8">
//     <meta name="viewport" content="width=device-width, initial-scale=1.0">
//     <title>Dynamic Block Label</title>
//     <style>
//         .block-container {
//             font-family: Arial, sans-serif;
//             margin: 20px;
//             font-size: 18px;
//         }
//         input {
//             padding: 8px;
//             font-size: 16px;
//             margin-left: 10px;
//         }
//     </style>
// </head>
// <body>

<div id="container"></div>;

{
  /* <script> */
}
// Create the BLOCK [YES/NO] label and textbox dynamically
function createBlockInput() {
  const container = document.getElementById('container');

  const label = document.createElement('label');
  label.textContent = 'BLOCK [YES/NO]';
  label.htmlFor = 'blockTextbox';
  label.style.fontWeight = 'bold';

  const input = document.createElement('input');
  input.type = 'text';
  input.id = 'blockTextbox';
  input.placeholder = 'Type YES or NO';
  input.style.marginLeft = '10px';
  input.style.padding = '8px';
  input.style.fontSize = '16px';

  input.addEventListener('input', function () {
    const value = this.value.toUpperCase();
    if (!['YES', 'NO'].includes(value)) {
      this.setCustomValidity('Please enter YES or NO');
    } else {
      this.setCustomValidity('');
    }
  });

  container.appendChild(label, input);
}

// Call the function to create the elements
createBlockInput();

// </script>

// </body>
// </html>
