// // The keyboard has been rendered for you
// import { renderKeyboard } from '/keyboard'
// document.getElementById('keyboard-container').addEventListener('click', checkGuess)

// // Some useful elements
// const guessContainer = document.getElementById('guess-container')
// const snowmanParts = document.getElementsByClassName('snowman-part')

// /*
// Challenge
// 1. Your challenge is to build a Christmas take on the classic game "Hangman" where a player attempts to guess a word by selecting letters to save a snowman from melting.
// - The snowman is made up of 6 parts: hat, arm, nose, scarf, head, and body. These are separate images and have been positioned with CSS.
// - At the start of the game, a player can see a number of dashes, with a dash for each letter of the word. So if the word was TREE the player would see - - - -
// - The player selects a letter.
// - If that letter is in the word, that letter replaces the dash in the corresponding position. For the word "TREE", if the player has selected the letter E, they will see --EE.
// - If the selected letter does not appear in the word, one part of the snowman gets removed.
// - If the player guesses the entire word, they win!
//     - any removed parts of the snowman are reinstated.
//     - the snowman gets sunglasses
//     - the message "You Win!" is displayed in the "guess-container" div.
// -If the player guesses wrong 6 times:
//     - only a puddle remains.
//     - the message "You Lose!" is displayed in the "guess-container" div.

// *** Stretch Goals ***

// - Disable the letter button once a letter has been used.
// - Add a "New Game" button that appears at the end of a game and resets the app. (You will need to create an array of words to guess)
// */

// // Set the word to guess
// const word = "gift"
// // 6 guesses for the 6 parts of the snowman
// let guesses = 6

// function checkGuess() {

// }

// renderKeyboard()

// The keyboard has been rendered for you
import { renderKeyboard } from '/keyboard';

// Some useful elements
const guessContainer = document.getElementById('guess-container');
const snowmanParts = document.getElementsByClassName('snowman-part');
const sunglasses = document.querySelector('.sunglasses');
const puddle = document.querySelector('.puddle');

// Array of words for the game
const words = ['gift', 'snow', 'tree', 'jolly', 'santa', 'elf'];
let word = words[Math.floor(Math.random() * words.length)]; // Random word to guess
let guesses = 6; // 6 guesses for 6 snowman parts
let guessArr = []; // Array to track guessed letters/dashes

// Initialize the game
function start() {
  word = words[Math.floor(Math.random() * words.length)]; // Pick a new random word
  guesses = 6;
  guessArr = [];
  for (let i = 0; i < word.length; i++) {
    guessArr.push('-');
  }
  // Reset snowman and UI elements
  for (let part of snowmanParts) {
    part.style.visibility = 'visible';
  }
  sunglasses.style.visibility = 'hidden';
  puddle.style.zIndex = '-2';
  // Remove any existing "New Game" button
  const existingButton = document.querySelector('.new-game-btn');
  if (existingButton) existingButton.remove();
  // Re-enable all letter buttons
  const letterButtons = document.querySelectorAll('.letter');
  letterButtons.forEach((button) => {
    button.disabled = false;
    button.style.backgroundColor = '#c03a2b'; // Reset to original color
    button.style.color = 'white';
  });
  renderGuess(); // Render initial dashes
  renderKeyboard(); // Re-render keyboard
}
start();

// Render the guess state
function renderGuess() {
  const guessHtml = guessArr.map((char) => {
    return `<div class="guess-char">${char}</div>`;
  });
  guessContainer.innerHTML = guessHtml.join('');
}

// Check if the game is over
function checkGameOver() {
  const allGuessed = guessArr.join('') === word;
  const noGuessesLeft = guesses === 0;

  if (allGuessed) {
    // Win condition
    guessContainer.innerHTML = '<div class="message">You Win!</div>';
    sunglasses.style.visibility = 'visible'; // Add sunglasses
    for (let part of snowmanParts) {
      part.style.visibility = 'visible'; // Reinstate all parts
    }
    addNewGameButton();
  } else if (noGuessesLeft) {
    // Lose condition
    guessContainer.innerHTML = '<div class="message">You Lose!</div>';
    puddle.style.zIndex = '2'; // Show puddle
    for (let part of snowmanParts) {
      part.style.visibility = 'hidden'; // Hide all parts
    }
    addNewGameButton();
  }
  return allGuessed || noGuessesLeft;
}

// Add a "New Game" button
function addNewGameButton() {
  const newGameButton = document.createElement('button');
  newGameButton.classList.add('new-game-btn');
  newGameButton.textContent = 'New Game';
  newGameButton.style.margin = '20px auto';
  newGameButton.style.display = 'block';
  newGameButton.style.padding = '10px 20px';
  newGameButton.style.backgroundColor = '#f0c419';
  newGameButton.style.border = 'none';
  newGameButton.style.borderRadius = '5px';
  newGameButton.style.fontSize = '1.2em';
  newGameButton.style.color = 'black';
  newGameButton.style.cursor = 'pointer';
  newGameButton.addEventListener('click', start);
  guessContainer.appendChild(newGameButton);
}

// Handle a letter guess
function checkGuess(event) {
  if (event.target.classList.contains('letter')) {
    const letter = event.target.id.toLowerCase();
    const button = event.target;

    // Disable the button after clicking
    button.disabled = true;
    button.style.backgroundColor = '#d9f7f7'; // Match the active state color
    button.style.color = 'black';

    let correctGuess = false;

    // Check if the letter is in the word
    for (let i = 0; i < word.length; i++) {
      if (word[i] === letter) {
        guessArr[i] = letter;
        correctGuess = true;
      }
    }

    // Update game state
    if (!correctGuess && guesses > 0) {
      guesses--;
      const partsToRemove = Array.from(snowmanParts);
      const partToHide = partsToRemove[guesses]; // Remove parts in order
      if (partToHide) partToHide.style.visibility = 'hidden';
    }

    // Render updated guess
    renderGuess();

    // Check if game is over
    checkGameOver();
  }
}

// Add event listener for keyboard clicks
document
  .getElementById('keyboard-container')
  .addEventListener('click', checkGuess);

// Render the keyboard
renderKeyboard();
