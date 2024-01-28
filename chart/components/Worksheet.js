const {
  getColName,
  getRowNames,
} = require('../services');
const CONTENT_TYPES = require('../ContentTypes');
const RELATION_TYPES = require('../relationTypes');

let worksheetCounter = 1;

class Worksheet {
  constructor(name) {
    this.contentType = CONTENT_TYPES.worksheet;
    this.relationType = RELATION_TYPES.worksheet;
    this.id = worksheetCounter;
    this.fileName = `sheet${worksheetCounter}.xml`;
    this.name = `${+new Date()}`;
    this.content = JSON.parse(JSON.stringify(require('../../template/xl/worksheets/sheet1.xml.json')));
    this.name = name;

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

    this.content.worksheet.sheetData.row = this.content.worksheet.sheetData.row || [];

    this.content.worksheet.sheetData.row.push({
      $: {
        r: 1 + rowOffset,
      },
      c: columnNames.map((columnName, columnIndex) => ({
        $: {
          r: `${getColName(columnIndex + 2 + columnOffset)}${1 + rowOffset}`,
          t: 'str',
        },
        v: [
          columnName,
        ],
      })),
    });

    rowNames.forEach((rowName, rowIndex) => this.content.worksheet.sheetData.row.push({
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
            data[columnName][rowName].value || '',
          ],
        })),
      ],
    }));
  }

  getRelationship(relationshipId) {
    return {
      $: {
        'sheetId': `${this.id}`,
        'name': this.name,
        'state': 'visible',
        'r:id': relationshipId,
      },
    };
  }
}

module.exports = Worksheet;
