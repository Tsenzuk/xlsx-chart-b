const {
  getColName,
  getRowNames,
} = require('../services');
const CONTENT_TYPES = require('../ContentTypes');
const RELATION_TYPES = require('../relationTypes');

let worksheetCounter = 1;

class Worksheet {
  constructor() {
    this.contentType = CONTENT_TYPES.worksheet;
    this.relationType = RELATION_TYPES.worksheet;
    this.id = worksheetCounter;
    this.fileName = `sheet${worksheetCounter}.xml`;
    this.name = `${+new Date()}`;
    this.content = JSON.parse(JSON.stringify(require('../../template/xl/worksheets/sheet1.xml.json')));

    worksheetCounter++;
  }

  /**
   *
   * @param {string} name
   */
  setWorksheetName(name) {
    this.name = name;
  }

  /**
   *
   * @param {Object.<string, Object.<string, number|string>>} data
   */
  setWorksheetData(data, rowOffset = 0, columnOffset = 0) {
    this._data = data;

    const columnNames = Object.keys(data);

    const rowNames = getRowNames(data);

    this.content.worksheet.sheetData.push({
      row: [
        {
          $: {
            r: 1 + rowOffset,
          },
          c: columnNames.map((columnName, columnIndex) => ({
            $: {
              r: `${getColName(columnIndex + 2 + columnOffset)}1`,
              t: 'str',
            },
            v: [
              columnName,
            ],
          })),
        },
        ...rowNames.map((rowName, rowIndex) => ({
          $: {
            r: rowIndex + 2 + rowOffset,
          },
          c: [
            {
              $: {
                r: `${getColName(1 + columnOffset)}${rowIndex + 2 + rowOffset}`,
                t: 'str',
              },
              v: [
                rowName,
              ],
            },
            ...columnNames.map((columnName, columnIndex) => ({
              $: {
                r: `${getColName(columnIndex + 2 + columnOffset)}${rowIndex + 2 + rowOffset}`,
              },
              v: [
                data[columnName][rowName] || '',
              ],
            })),
          ],
        })),
      ],
    });
  }
}

module.exports = Worksheet;
