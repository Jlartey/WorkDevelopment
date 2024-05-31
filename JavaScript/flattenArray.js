const nestedArray = [
  [1, 2],
  [3, 4],
  [5, 6],
];
const flattenedArray = nestedArray.reduce((acc, curr) => acc.concat(curr), []);
console.log(flattenedArray); // Output: [1, 2, 3, 4, 5, 6]
