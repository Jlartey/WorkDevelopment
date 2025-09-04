/* 
This Christmas, you’ve been tasked with running an anagram quiz at 
the family gathering.

You have been given a list of anagrams, but you suspect that some 
of the anagram pairs might be incorrect.

Your job is to write a JavaScript function to loop through the array
and filter out any pairs that aren’t actually anagrams.

For this challenge, spaces will be ignored, so "Be The Helm" would 
be considered a valid anagram of "Bethlehem".
*/

let anagrams = [
  ['Can Assault', 'Santa Claus'],
  ['Refreshed Erudite Londoner', 'Rudolf the Red Nose Reindeer'],
  ['Frosty The Snowman', 'Honesty Warms Front'],
  ['Drastic Charms', 'Christmas Cards'],
  ['Congress Liar', 'Carol Singers'],
  ['The Tin Glints', 'Silent Night'],
  ['Be The Helm', 'Betlehem'],
  ['Is Car Thieves', 'Christmas Eve'],
];

function findAnagrams(array) {
  for (let i = 0; i < array.length; i++) {
    return (array = array.filter((pair) => checkAnagrams(pair)));
  }
}

function checkAnagrams([word1, word2]) {
  const formattedStr = (str) =>
    str.replace(/\s/g, '').toLowerCase().split('').sort().join('');
  return formattedStr(word1) === formattedStr(word2);
}

console.log(findAnagrams(anagrams));

// Alternate Solution
function findAnagrams(array) {
  // Helper function to normalize a string: remove spaces, convert to lowercase, and sort letters
  function normalizeString(str) {
    return str
      .replace(/\s/g, '') // Remove all spaces
      .toLowerCase() // Convert to lowercase
      .split('') // Split into array of characters
      .sort() // Sort alphabetically
      .join(''); // Join back into a string
  }

  // Filter the array to keep only valid anagram pairs
  return array.filter((pair) => {
    let [str1, str2] = pair; // Destructure the pair into two strings
    return normalizeString(str1) === normalizeString(str2); // Compare normalized strings
  });
}

// Test the function
console.log(findAnagrams(anagrams));
