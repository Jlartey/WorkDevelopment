// function getDrugCodes(code) {
//   const drugs = [];

//   for (const result of code.split('||||||||||||||~~')) {
//     drugs.push(result.substring(0, 16));
//   }
//   return drugs.toString(' ');
// }

// console.log(
//   getDrugCodes(
//     'EDD-3513-0500-00||||Oral||BID||||||||||||||~~CVD-1003-0010-00||||Oral||Daily||||||||||||||~~CVD-1803-0050-00||||Oral||Nocte||||||||||||||~~EWD-4513-0600-00||||Oral||TID||||||||||||||~~LRD-1013-0020-00||||||||||||||||||||'
//   )
// );

/**
 * Extracts drug codes from a string, where codes are separated by a delimiter.
 * Each code is assumed to be the first 16 characters of a segment.
 * @param {string} input - The input string containing drug codes.
 * @param {Object} [options] - Optional configuration.
 * @param {string} [options.delimiter='||||||||||||||~~'] - The delimiter separating codes.
 * @param {number} [options.codeLength=16] - The expected length of each drug code.
 * @returns {string[]} An array of extracted drug codes.
 * @throws {Error} If the input is invalid.
 */
function getDrugCodes(input, options = {}) {
  // Default options
  const { delimiter = '||||||||||||||~~', codeLength = 16 } = options;

  // Input validation
  if (typeof input !== 'string' || input.length === 0) {
    throw new Error('Input must be a non-empty string');
  }

  // Split input and extract codes
  const drugCodes = input
    .split(delimiter)
    .map((segment) => segment.trim().substring(0, codeLength))
    .filter((code) => code.length === codeLength); // Ensure valid code length

  return drugCodes;
}

// Example usage
try {
  const codes = getDrugCodes(
    'EDD-3513-0500-00||||Oral||BID||||||||||||||~~CVD-1003-0010-00||||Oral||Daily||||||||||||||~~CVD-1803-0050-00||||Oral||Nocte||||||||||||||~~EWD-4513-0600-00||||Oral||TID||||||||||||||~~LRD-1013-0020-00||||||||||||||||||||'
  );
  console.log(codes); // ['EDD-3513-0500-00', 'CVD-1003-0010-00', 'CVD-1803-0050-00', 'EWD-4513-0600-00', 'LRD-1013-0020-00']
  console.log(codes.join(' ')); // 'EDD-3513-0500-00 CVD-1003-0010-00 CVD-1803-0050-00 EWD-4513-0600-00 LRD-1013-0020-00'
} catch (error) {
  console.error('Error:', error.message);
}
