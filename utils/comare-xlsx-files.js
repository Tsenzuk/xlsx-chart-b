const fs = require('fs');
const { promisify } = require('util');
const JSZip = require('jszip');
const xml2js = require('xml2js');
const { transform, isEqual, isArray, isObject } = require('lodash');


/**
 * Find difference between two objects
 * @link https://davidwells.io/snippets/get-difference-between-two-objects-javascript
 * @param  {object} origObj - Source object to compare newObj against
 * @param  {object} newObj  - New object with potential changes
 * @return {object} differences
 */
function difference(origObj, newObj) {
  function changes(newObj, origObj) {
    let arrayIndexCounter = 0;
    return transform(newObj, function (result, value, key) {
      if (!isEqual(value, origObj[key])) {
        const resultKey = isArray(origObj) ? arrayIndexCounter++ : key;
        if (isObject(value) && isObject(origObj[key])) {
          result[resultKey] = changes(value, origObj[key]);
        } else {
          // const origValue = isObject(origObj[key]) ? JSON.stringify(origObj[key]) : origObj[key];
          const origValue = JSON.stringify(origObj[key])?.replace(/"/g, '\'');
          // const newValue = isObject(value) ? JSON.stringify(value) : value;
          const newValue = JSON.stringify(value)?.replace(/"/g, '\'');
          result[resultKey] = [`-${origValue}`, `+${newValue}`];
        }
      }
    });
  }
  return changes(newObj, origObj);
}


const file1 = process.argv[2];
const file2 = process.argv[3];
const fileToShowDetails = process.argv[4];

async function main(file1, file2, fileToShowDetails) {
  const p1 = promisify(fs.readFile)(file1);
  const p2 = promisify(fs.readFile)(file2);

  const [data1, data2] = await Promise.all([p1, p2]);

  const pz1 = JSZip.loadAsync(data1);
  const pz2 = JSZip.loadAsync(data2);

  const [zip1, zip2] = await Promise.all([pz1, pz2]);

  const files1 = [];
  zip1.forEach((relativePath, zipEntry) => {
    files1.push(relativePath);
  });

  const files2 = [];
  zip2.forEach((relativePath, zipEntry) => {
    files2.push(relativePath);
  });

  const files = files1.concat(files2);
  const uniqueFiles = [...new Set(files)];

  const promises = uniqueFiles.map(async (file) => {
    if (file.endsWith('/')) {
      return;
    }
    const data1 = await zip1.file(file)?.async('string');
    const data2 = await zip2.file(file)?.async('string');

    if (!data1) {
      console.log(`File ${file} is missing in ${file1}`);
      return;
    }

    if (!data2) {
      console.log(`File ${file} is missing in ${file2}`);
      return;
    }

    if (data1 !== data2) {
      console.log(`File ${file} is different`);
      const json1 = await xml2js.parseStringPromise(data1);
      const json2 = await xml2js.parseStringPromise(data2);

      if (file === fileToShowDetails) {
        console.log(JSON.stringify(difference(json1, json2), null, 2));
      }
    }
  });

  await Promise.all(promises);

}

main(file1, file2, fileToShowDetails);
