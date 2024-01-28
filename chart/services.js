/**
 *
 * @param {number} n - starts with 1
 * @returns {string}
 */
function getColName(n) {
  const abc = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  n--;
  if (n < 26) {
    return abc[n];
  }

  return getColName(Math.floor(n / 26)) + abc[n % 26];
}

function getRowNames (data) {
  const rowNames = [...new Set(Object.values(data).reduce((memo, columnData) => memo.concat(...Object.keys(columnData)), []))];
  return rowNames;
}

module.exports = {
  getColName,
  getRowNames,
};
