function gcdOfStrings(str1, str2) {
  if (str1 + str2 !== str2 + str1) {
    return '';
  }

  function gcd(a, b) {
    while (b) {
      [a, b] = [b, a % b];
    }
    return a;
  }

  const len = gcd(str1.length, str2.length);
  return str1.substring(0, len);
}

function gcd(a, b) {
  while (b) {
    [a, b] = [b, a % b];
    console.log([a, b]);
  }
  return a;
}

const result = gcd(4, 6);
console.log(result);
