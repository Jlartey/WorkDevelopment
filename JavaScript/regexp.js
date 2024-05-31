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
