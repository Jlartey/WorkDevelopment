const numbers = [1, 2, 3, 4, 5];
const sum = numbers.reduce((acc, curr) => acc + curr, 0);
console.log(sum); // Output: 15

const nestedArray = [
  [1, 2],
  [3, 4],
  [5, 6],
];
const flattenedArray = nestedArray.reduce((acc, curr) => acc.concat(curr), []);
console.log(flattenedArray); // Output: [1, 2, 3, 4, 5, 6]

const people = [
  { name: 'Alice', age: 25 },
  { name: 'Bob', age: 30 },
  { name: 'Charlie', age: 25 },
  { name: 'Dave', age: 30 },
];

const groupedByAge = people.reduce((acc, curr) => {
  if (!acc[curr.age]) {
    acc[curr.age] = [];
  }
  acc[curr.age].push(curr);
  return acc;
}, {});

console.log(groupedByAge);
/*
Output:
{
  '25': [{ name: 'Alice', age: 25 }, { name: 'Charlie', age: 25 }],
  '30': [{ name: 'Bob', age: 30 }, { name: 'Dave', age: 30 }]
}
*/
const fruits = [
  'apple',
  'banana',
  'apple',
  'orange',
  'banana',
  'apple',
  'orange',
  'guava',
];

const fruitCounts = fruits.reduce((acc, curr) => {
  acc[curr] = (acc[curr] || 0) + 1;
  return acc;
}, {});

console.log(fruitCounts);

const products = [
  { id: 1, name: 'Laptop', price: 999 },
  { id: 2, name: 'Phone', price: 699 },
  { id: 3, name: 'Tablet', price: 499 },
];

const productMap = products.reduce((acc, curr) => {
  acc[curr.id] = curr;
  return acc;
}, {});

console.log(productMap);
/*
Output:
{
  '1': { id: 1, name: 'Laptop', price: 999 },
  '2': { id: 2, name: 'Phone', price: 699 },
  '3': { id: 3, name: 'Tablet', price: 499 }
}
*/

// Accessing a product by ID
const laptop = productMap[1];
console.log(laptop); // Output: { id: 1, name: 'Laptop', price: 999 }

const add5 = (x) => x + 5;
const multiply3 = (x) => x * 3;
const subtract2 = (x) => x - 2;

const composedFunctions = [add5, multiply3, subtract2];

const result = composedFunctions.reduce((acc, curr) => curr(acc), 10);
console.log(result); // Output: 43

// --------------------

const initialState = {
  count: 0,
  todos: [],
};

const actions = [
  { type: 'INCREMENT_COUNT' },
  { type: 'ADD_TODO', payload: 'Learn Array.reduce()' },
  { type: 'INCREMENT_COUNT' },
  { type: 'ADD_TODO', payload: 'Master TypeScript' },
];

const reducer = (state, action) => {
  switch (action.type) {
    case 'INCREMENT_COUNT':
      return { ...state, count: state.count + 1 };
    case 'ADD_TODO':
      return { ...state, todos: [...state.todos, action.payload] };
    default:
      return state;
  }
};

const finalState = actions.reduce(reducer, initialState);
console.log(finalState);
/*
Output:
{
  count: 2,
  todos: ['Learn Array.reduce()', 'Master TypeScript']
}
*/

const numbers3 = [1, 2, 3, 2, 4, 3, 5, 1, 6];

const uniqueNumbers = numbers3.reduce((acc, curr) => {
  if (!acc.includes(curr)) {
    acc.push(curr);
  }
  return acc;
}, []);

console.log(uniqueNumbers); // Output: [1, 2, 3, 4, 5, 6]

// -------------------------
const grades = [85, 90, 92, 88, 95];

const average = grades.reduce((acc, curr, index, array) => {
  acc += curr;
  if (index === array.length - 1) {
    return acc / array.length;
  }
  return acc;
}, 0);

console.log(average); // Output: 90
