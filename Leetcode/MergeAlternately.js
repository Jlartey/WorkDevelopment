function mergeAlternately(word1, word2) {
  const alternateWord = [];

  const difference = Math.abs(word1.length - word2.length);

  let longestWord;
  let shortestWord;

  if (word1.length < word2.length) {
    shortestWord = word1;
    longestWord = word2;
  } else {
    shortestWord = word2;
    longestWord = word1;
  }

  for (i = 0; i < shortestWord.length; i++) {
    alternateWord.push(word1[i] + word2[i]);
  }
  return difference == 0
    ? alternateWord.join('')
    : alternateWord.join('') + longestWord.slice(-difference);
}

const result = mergeAlternately('Joe', 'Lartey');

console.log(result);

function mergeAlternately01(word1, word2) {
  let result = '';
  const len1 = word1.length;
  const len2 = word2.length;
  const minLen = Math.min(len1, len2);

  for (let i = 0; i < minLen; i++) {
    result += word1[i] + word2[i];
  }

  return result + word1.slice(minLen) + word2.slice(minLen);
}

const result2 = mergeAlternately('Joe', 'Lartey');

console.log(result2);
